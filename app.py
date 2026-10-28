"""Документы качества 8.0: бесплатный локальный OCR для кириллицы."""
from __future__ import annotations

import hmac
import io
import ipaddress
import logging
import os
import re
import secrets
import socket
import time
import tempfile
from collections import OrderedDict, defaultdict, deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import PurePosixPath
from threading import Lock, BoundedSemaphore
from urllib.parse import urljoin, urlsplit

import fitz
import requests
from bs4 import BeautifulSoup
from flask import Flask, jsonify, render_template, request, send_file, Response
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from extractor import DOC_TYPES, norm_space
from local_ocr import RecognitionError, inspect_document, recognition_status

app=Flask(__name__)
app.json.ensure_ascii=False
app.config['MAX_CONTENT_LENGTH']=26*1024*1024
log=logging.getLogger('quality-docs')
logging.basicConfig(level=logging.INFO)
job_semaphore=BoundedSemaphore(1)
rate_cache=defaultdict(deque)
rate_lock=Lock()
file_cache=OrderedDict()
file_lock=Lock()
search_lock=Lock()
search_cache=OrderedDict()
DOC_SELECT={'all':'Все документы',**DOC_TYPES}
DOC_KEYWORDS={
    'certificate':['сертификат соответствия'],
    'declaration':['декларация о соответствии'],
    'passport':['паспорт качества'],
    'test':['протокол испытаний'],
    'refusal':['отказное письмо'],
}
UA={'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36',
    'Accept':'application/pdf,text/html,*/*', 'Accept-Language':'ru-RU,ru;q=0.9,en;q=0.7'}
MAX_BYTES=25*1024*1024

@app.before_request
def before_request():
    password=os.environ.get('SITE_PASSWORD','')
    if password:
        auth=request.authorization
        if not auth or not hmac.compare_digest(password,auth.password or ''):
            return Response('Введите пароль доступа',401,{'WWW-Authenticate':'Basic realm="Документы качества"'})
    policy={'/api/search':8,'/api/upload':12,'/api/export':25}
    if request.path not in policy:return None
    now=time.monotonic(); key=(request.remote_addr,request.path)
    with rate_lock:
        q=rate_cache[key]
        while q and now-q[0]>900:q.popleft()
        if len(q)>=policy[request.path]:return jsonify({'error':'Слишком много запросов. Повторите позднее.'}),429
        q.append(now)
    return None

@app.after_request
def headers(response):
    response.headers['X-Content-Type-Options']='nosniff'
    response.headers['X-Frame-Options']='DENY'
    response.headers['Referrer-Policy']='no-referrer'
    if request.path.startswith('/api/'):response.headers['Cache-Control']='no-store'
    return response

def public_url(url):
    try:
        p=urlsplit(url)
        if p.scheme not in ('https','http') or not p.hostname or p.username or p.password or p.port not in (None,80,443):return False
        if p.hostname.lower() in ('localhost','localhost.localdomain') or p.hostname.endswith(('.local','.internal')):return False
        try:return ipaddress.ip_address(p.hostname).is_global
        except ValueError:return True
    except ValueError:return False

def verify_host(url):
    if not public_url(url):raise ValueError('Непубличный URL запрещён')
    p=urlsplit(url)
    addresses=socket.getaddrinfo(p.hostname,p.port or (443 if p.scheme=='https' else 80),type=socket.SOCK_STREAM)
    if not addresses or any(not ipaddress.ip_address(a[4][0]).is_global for a in addresses):
        raise ValueError('Ссылка указывает на недопустимый сетевой адрес')

def fetch(url,max_bytes=MAX_BYTES,timeout=10):
    current=url
    for _ in range(5):
        verify_host(current)
        r=None
        try:
            r=requests.get(current,headers=UA,stream=True,timeout=(4,timeout),allow_redirects=False)
            if r.status_code in (301,302,303,307,308):
                nxt=urljoin(current,r.headers.get('Location',''))
                if nxt==current:raise ValueError('Зацикленное перенаправление')
                current=nxt
                continue
            r.raise_for_status()
            length=r.headers.get('Content-Length','')
            if length.isdecimal() and int(length)>max_bytes:raise ValueError('PDF больше 25 МБ')
            data=bytearray()
            for chunk in r.iter_content(128*1024):
                data.extend(chunk)
                if len(data)>max_bytes:raise ValueError('Ответ превышает лимит')
            return bytes(data),r.headers.get('Content-Type','').lower(),current
        finally:
            if r is not None:r.close()
    raise ValueError('Слишком много перенаправлений')

def file_id_and_save(data: bytes) -> str:
    """Keep short-lived PDFs on temporary disk instead of consuming Render RAM.

    Render filesystem is ephemeral: files only exist in this running instance.
    """
    fid=secrets.token_urlsafe(16)
    with tempfile.NamedTemporaryFile(prefix='quality_pdf_',suffix='.pdf',delete=False) as out:
        out.write(data)
        path=out.name
    now=time.monotonic()
    with file_lock:
        for k in list(file_cache):
            old_path,old_at,old_size=file_cache[k]
            if now-old_at>2400:
                file_cache.pop(k,None)
                try:os.unlink(old_path)
                except OSError:pass
        file_cache[fid]=(path,now,len(data))
        file_cache.move_to_end(fid)
        while sum(v[2] for v in file_cache.values())>150*1024*1024:
            _,(old_path,_,_)=file_cache.popitem(last=False)
            try:os.unlink(old_path)
            except OSError:pass
    return fid


def search_queries(material,dtype,mfr,standard):
    """Two non-identical searches: specific request and wider product document.

    A certificate can cover a range of diameters, not the exact queried diameter.
    """
    clean=norm_space(re.sub(r'[«»"“”]', ' ', material))
    extras=norm_space(' '.join(filter(None,[mfr,standard])))
    selected=(DOC_KEYWORDS.get(dtype) or DOC_KEYWORDS['certificate'])[0]
    broad=re.sub(r'(?:[Øø∅]|\b(?:d|диаметр)\s*)\s*\d{1,3}\s*(?:мм)?\b',' ',clean,flags=re.I)
    broad=norm_space(broad) or clean
    second_type = selected if dtype!='all' else 'сертификат паспорт декларация'
    return list(dict.fromkeys((norm_space(q) for q in (
        f'{clean} {selected} pdf {extras}',
        f'{broad} {second_type} pdf {extras}',
    ))))


def search_hits(queries):
    """Search sources are not proof that a document or its URL works."""
    found=[]; problems=[]; attempts=0; successful=0
    api_key=os.environ.get('BRAVE_SEARCH_API_KEY','').strip()
    for q in queries[:2]:
        items=[]
        if api_key:
            attempts+=1
            try:
                resp=requests.get('https://api.search.brave.com/res/v1/web/search',
                    params={'q':q,'count':15,'search_lang':'ru'},
                    headers={'X-Subscription-Token':api_key,'Accept':'application/json'},timeout=(4,10))
                resp.raise_for_status()
                items=[{'href':v.get('url'),'title':v.get('title'),'body':v.get('description')}
                       for v in resp.json().get('web',{}).get('results',[])]
                successful+=1
            except Exception as exc:
                problems.append(type(exc).__name__)
                log.warning('Search API error: %s',repr(exc)[:240])
        else:
            from ddgs import DDGS
            # Never make a single backend mandatory. Shared cloud IPs can be blocked.
            for backend in ('auto','google'):
                attempts+=1
                try:
                    with DDGS(timeout=8) as ddgs:
                        items=ddgs.text(q,region='ru-ru',safesearch='moderate',
                                       max_results=12,backend=backend) or []
                    successful+=1
                    if items:break
                except Exception as exc:
                    problems.append(type(exc).__name__)
                    log.warning('Search %s error: %s',backend,repr(exc)[:240])
        for v in items:
            url=v.get('href') or v.get('url')
            if public_url(url or ''):
                found.append({'url':url,'title':norm_space(v.get('title') or ''),
                              'snippet':norm_space(v.get('body') or v.get('description') or '')})
    unique={}
    for hit in found:
        unique.setdefault(hit['url'].split('#')[0],hit)
    warnings=[]
    if not successful:
        warnings.append('Поисковые системы не ответили. Это не означает, что документов нет.')
    elif not unique:
        warnings.append('Поисковые системы ответили, но ссылок по запросу не вернули.')
    elif problems:
        warnings.append('Часть поисковых источников недоступна; использованы ответы остальных.')
    return list(unique.values())[:25],warnings


def expand_page(hit):
    # Look for PDF anchors on a verified public HTML page, not unrelated external domains.
    try:
        content,kind,base=fetch(hit['url'],max_bytes=1_300_000,timeout=8)
        if content.lstrip()[:5]==b'%PDF-':
            return [{'url':base,'title':hit['title'],'snippet':hit['snippet']}]
        if 'html' not in kind and not content.lstrip().startswith((b'<!',b'<ht')):return []
        page=BeautifulSoup(content,'html.parser')
        links=[]
        for a in page.select('a[href]'):
            url=urljoin(base,a['href'])
            if public_url(url) and re.search(r'\.pdf(?:[?#]|$)',url,re.I):
                caption=norm_space(a.get_text(' ',strip=True))
                links.append({'url':url,'title':caption or hit['title'],'snippet':hit['snippet']})
            if len(links)>=25:break
        return links
    except Exception:return None


def pertinent(text,material):
    """Conservative product relevance; no diameter requirement for range certificates."""
    def norm(x):
        return x.lower().replace('a500','а500').replace('a400','а400').replace('a600','а600')
    clean=norm(text)
    words={w[:4] for w in re.findall(r'[а-яa-z0-9]+',norm(material)) if len(w)>=4
           and w not in ('документ','паспорт','сертификат','качества')}
    return not words or any(w in clean for w in words)


def link_problem(exc):
    """A 403/timeout means 'unverified', NEVER 'broken'."""
    if isinstance(exc, requests.HTTPError) and exc.response is not None:
        code=exc.response.status_code
        if code in (404,410):return 'gone'
        if code in (401,403,429):return 'restricted'
        return 'http_'+str(code)
    if isinstance(exc, requests.Timeout):return 'timeout'
    if isinstance(exc, requests.ConnectionError):return 'connection'
    if isinstance(exc, ValueError):return 'invalid'
    return type(exc).__name__.lower()


def candidate_link(hit, state, note='', fid=''):
    return {'url':hit['url'],'title':(hit.get('title') or
              PurePosixPath(urlsplit(hit['url']).path).name or 'Источник документа')[:190],
            'state':state, 'note':note[:170], 'id':fid}


def material_doc_record(info,source_id,source_url):
    return {'material':info['material'] or 'Не удалось извлечь',
            'number':info['number'] or 'б/н',
            'date':info['date'] or '—',
            'manufacturer':info['manufacturer'] or '—',
            'type':DOC_TYPES.get(info['type'],'Вид документа не определён'),
            'id':source_id,'source_url':source_url,
            'has_text':info['has_text'],
            'ocr_note':info['ocr_note']}


def perform_search(material,dtype,manufacturer,standard):
    """Never silently erase search results because Render cannot download them.

    Verified PDF and unverified source links are two independent result categories.
    Unverified links are NOT given invented material/date/manufacturer fields.
    Definitively dead HTTP 404/410 links are not displayed.
    """
    queries=search_queries(material,dtype,manufacturer,standard)
    hits,warnings=search_hits(queries)
    direct=[h for h in hits if re.search(r'\.pdf(?:[?#]|$)',h['url'],re.I)
            and not re.search(r'(?i)(?:katalog|catalog|каталог|pricelist|price[-_]?list)',h['url']+' '+h['title'])]
    pages=[h for h in hits if h not in direct][:8]
    # HTML landing pages can contain PDF buttons, even with non-.pdf URLs.
    # Keep them as possible source pages if there was no extractable PDF.
    page_links=[]
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures={pool.submit(expand_page,h):h for h in pages}
        for f in as_completed(futures):
            hit=futures[f]
            try:
                expanded=f.result()
                direct.extend(expanded)
                if expanded is not None and not expanded and pertinent(hit['title']+' '+hit['snippet'],material):
                    page_links.append(hit)
            except Exception:
                page_links.append(hit)
    unique={}
    for h in direct:
        if public_url(h['url']):unique.setdefault(h['url'].split('#')[0],h)
    candidates=sorted(unique.values(), key=lambda h: (pertinent(h['title']+' '+h['snippet'],material),
        any(x in h['title'].lower() for x in ('сертификат','декларация','паспорт'))), reverse=True)[:12]
    items=[]; sources=[]; reasons={}; ignored=0; scanned=0
    def note(reason):
        reasons[reason]=reasons.get(reason,0)+1
    def probe(hit):
        try:
            data,kind,actual=fetch(hit['url'],max_bytes=6*1024*1024,timeout=10)
            if not data.lstrip()[:1024].startswith(b'%PDF-'):
                return (hit,None,None,None),'not_pdf'
            filename=PurePosixPath(urlsplit(actual).path).name or 'document.pdf'
            info=inspect_document(data,material,filename,use_ocr=False)
            return (hit,data,info,actual),None
        except Exception as exc:
            reason=link_problem(exc)
            log.warning('PDF check %s: %s (%s)',urlsplit(hit['url']).hostname,reason,str(exc)[:125])
            return (hit,None,None,None),reason

    checked=[]
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures=[pool.submit(probe,h) for h in candidates]
        for f in as_completed(futures):
            item,problem=f.result()
            hit,data,info,actual=item
            if problem:
                note(problem)
                if problem not in ('gone','invalid','not_pdf'):
                    sources.append(candidate_link(hit,'unverified','Не удалось проверить PDF на сервере: '+{'restricted':'доступ ограничен','timeout':'время ожидания истекло','connection':'сбой соединения'}.get(problem,problem)+'. Можно попробовать открыть ссылку в браузере.'))
            else:
                checked.append(item)
    checked.sort(key=lambda x:(bool(x[2].get('type')),
                 any(k in x[0]['title'].lower() for k in ('сертификат','декларация','паспорт'))),reverse=True)
    ocr_budget=max(0,min(2,int(os.getenv('OCR_SEARCH_LIMIT','1'))))
    for hit,data,info,actual in checked:
        filename=info['filename']
        if re.search(r'(?i)(?:katalog|catalog|каталог|pricelist|price[-_]?list)',filename):
            ignored+=1
            continue
        if not info['type'] or not info['material']:
            if ocr_budget:
                ocr_budget-=1
                try:info=inspect_document(data,material,filename,use_ocr=True)
                except RecognitionError as exc:
                    log.warning('OCR search candidate: %s',str(exc)[:130])
                    note('ocr')
        fid=file_id_and_save(data)
        if info.get('type') and (dtype=='all' or info['type']==dtype) and \
           info.get('material') and pertinent(info.get('text','')+' '+info['material'],material) and \
           (not manufacturer or not info.get('manufacturer') or
            info['manufacturer']=='—' or manufacturer.lower() in info['manufacturer'].lower()):
            items.append(material_doc_record(info,fid,actual))
        else:
            # Actual PDF was downloaded, but automatic classification is uncertain.
            # Make the known-good PDF accessible without claiming it matches.
            if pertinent(hit['title']+' '+hit['snippet'],material) or info.get('type'):
                sources.append(candidate_link(hit,'downloadable',
                    'PDF скачан, но его принадлежность запросу не подтверждена.',fid))
            else:ignored+=1
        if len(items)>=10:break
    # Landing pages are clearly separated from PDFs and never labeled verified.
    for hit in page_links[:6]:
        if pertinent(hit['title']+' '+hit['snippet'],material):
            sources.append(candidate_link(hit,'page','Страница из поисковой выдачи; PDF не извлечён.'))
    distinct={}
    for h in sources:
        if public_url(h['url']):distinct.setdefault(h['url'].split('#')[0],h)
    sources=list(distinct.values())[:16]
    total=len(candidates)
    if not items and sources:
        warnings.append(f'Поиск нашёл {total} PDF-кандидатов, но сервер не подтвердил реквизиты. '
                        'Ссылки для самостоятельной проверки приведены отдельно.')
    if reasons:
        view={'timeout':'тайм-аут','restricted':'доступ ограничен (401/403/429)',
              'connection':'сбой соединения','gone':'удалённая ссылка (404/410)',
              'not_pdf':'ссылка ведёт не на PDF', 'ocr':'сбой OCR'}
        warnings.append('Проверка файлов: '+', '.join(f'{view.get(k,k)} — {v}' for k,v in sorted(reasons.items()))+'.')
    if ignored:warnings.append(f'Каталоги или неподходящие документы исключены: {ignored}.')
    return items,warnings,sources,{'found_links':len(hits),'pdf_candidates':total,
                                  'verified_pdf':len(items),'unverified_sources':len(sources)}

@app.get('/')
def index():return render_template('index.html',types=DOC_SELECT)

@app.get('/health')
def health():return jsonify({'ok':True,'version':'quality-docs-v9-source-visibility','recognition':recognition_status()})

@app.post('/api/search')
def api_search():
    d=request.get_json(silent=True) or {}
    material=norm_space(d.get('material',''))[:160]
    dtype=str(d.get('type','all'))
    manufacturer=norm_space(d.get('manufacturer',''))[:100]
    standard=norm_space(d.get('standard',''))[:100]
    if len(material)<3:return jsonify({'error':'Укажите наименование материала.'}),400
    if dtype not in DOC_SELECT:return jsonify({'error':'Некорректный вид документа'}),400
    if not job_semaphore.acquire(blocking=False):return jsonify({'error':'Поиск занят. Попробуйте повторить запрос.'}),429
    try:
        results,warnings,sources,stats=perform_search(material,dtype,manufacturer,standard)
    except Exception:
        log.exception('search failed')
        return jsonify({'error':'Поиск не выполнен из-за ошибки сервера. Повторите попытку.'}),503
    finally:job_semaphore.release()
    return jsonify({'results':results,'warnings':warnings,'sources':sources,'stats':stats,'search_status':('unavailable' if any('Поисковые системы не ответили' in w for w in warnings) else 'completed')})

@app.post('/api/upload')
def api_upload():
    file=request.files.get('file')
    if file is None or not file.filename:return jsonify({'error':'Выберите PDF-файл.'}),400
    b=file.stream.read(MAX_BYTES+1)
    if len(b)>MAX_BYTES:return jsonify({'error':'Файл больше 25 МБ'}),413
    if not job_semaphore.acquire(blocking=False):return jsonify({'error':'Анализ занят. Повторите попытку.'}),429
    try:
        info=inspect_document(b,norm_space(request.form.get('material',''))[:160],file.filename,use_ocr=True)
    except Exception as exc:
        return jsonify({'error':'Не удалось обработать файл: '+str(exc)[:150]}),422
    finally:job_semaphore.release()
    # Show result, even if no document type: user explicitly uploaded the file.
    fid=file_id_and_save(b)
    record=material_doc_record(info,fid,'')
    record['warning']=('Это не распознанный сертификат, декларация или паспорт качества.' if not info['type'] else '')
    if info.get('ocr_note') and ('недоступен' in info['ocr_note'] or 'не обнаружило' in info['ocr_note']):
        record['warning'] += ' '+info['ocr_note']
    return jsonify({'result':record})

@app.get('/api/file/<file_id>')
def api_file(file_id):
    with file_lock:
        record=file_cache.get(file_id)
    if record is None or time.monotonic()-record[1]>2400 or not os.path.isfile(record[0]):
        return jsonify({'error':'Временная копия недоступна. Повторите поиск.'}),404
    return send_file(record[0], mimetype='application/pdf',as_attachment=True,download_name='document.pdf')

@app.post('/api/export')
def api_export():
    rows=(request.get_json(silent=True) or {}).get('results',[])
    if not isinstance(rows,list) or len(rows)>100:return jsonify({'error':'Некорректная таблица'}),400
    workbook=Workbook(); ws=workbook.active;ws.title='Документы качества'
    headers=['Наименование материала','Номер документа','Дата документа','Производитель']
    keys=['material','number','date','manufacturer']
    ws.append(headers);ws.freeze_panes='A2'
    for row in rows:
        if not isinstance(row,dict):continue
        values=[]
        for k in keys:
            val=str(row.get(k,'') or '')[:500]
            if val.lstrip().startswith(('=','+','-','@')):val="'"+val
            values.append(val)
        ws.append(values)
    ws.auto_filter.ref=f'A1:D{max(2,ws.max_row)}'
    for cell in ws[1]:
        cell.font=Font(name='Calibri',bold=True,color='FFFFFF')
        cell.fill=PatternFill('solid',fgColor='183557')
        cell.alignment=Alignment(wrap_text=True)
    for idx,width in enumerate((62,31,40,55),1):ws.column_dimensions[get_column_letter(idx)].width=width
    for row in ws.iter_rows(min_row=2):
        for cell in row:cell.alignment=Alignment(wrap_text=True,vertical='top')
        ws.row_dimensions[row[0].row].height=44
    stream=io.BytesIO(); workbook.save(stream);stream.seek(0)
    return send_file(stream,as_attachment=True,download_name='dokumenty_kachestva.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__=='__main__':
    app.run(host='0.0.0.0',port=int(os.getenv('PORT',10000)))
