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
    now=time.monotonic()
    fid=secrets.token_urlsafe(16)
    with file_lock:
        for k in list(file_cache):
            if now-file_cache[k][1]>2400: file_cache.pop(k,None)
        file_cache[fid]=(data,now)
        file_cache.move_to_end(fid)
        while sum(len(x[0]) for x in file_cache.values()) > 42*1024*1024:
            file_cache.popitem(last=False)
    return fid


def search_queries(material,dtype,mfr,standard):
    # Two focused queries keep external search traffic and free-Render latency low.
    extras=' '.join(filter(None,[mfr,standard])).strip()
    clean=re.sub(r'[«»"“”]', ' ', material)
    clean=norm_space(clean)
    selected=(DOC_KEYWORDS.get(dtype) or DOC_KEYWORDS['certificate'])[0]
    phrases=[f'{clean} {selected} pdf {extras}',
             f'{clean} сертификат паспорт качества файл pdf {extras}']
    if dtype=='all':
        phrases[1]=f'{clean} декларация соответствия паспорт качества pdf {extras}'
    return list(dict.fromkeys(norm_space(q) for q in phrases))[:2]


def search_hits(queries):
    """Try lightweight public search sources, without changing the OCR engine.

    DDGS is an unofficial web-search client. Hosting providers may experience
    CAPTCHA, rate limits or connection timeouts; no availability is guaranteed.
    """
    found=[]
    api_key=os.environ.get('BRAVE_SEARCH_API_KEY','').strip()
    problems=[]
    had_success=False
    for q in queries[:2]:
        items=[]
        if api_key:
            try:
                resp=requests.get('https://api.search.brave.com/res/v1/web/search',
                    params={'q':q,'count':12,'country':'RU','search_lang':'ru'},
                    headers={'X-Subscription-Token':api_key,'Accept':'application/json'},timeout=(4,9))
                resp.raise_for_status()
                items=[{'href':v.get('url'),'title':v.get('title'),'body':v.get('description')}
                       for v in resp.json().get('web',{}).get('results',[])]
                had_success=True
            except Exception as exc:
                problems.append(type(exc).__name__)
                log.warning('Brave Search API failed: %s',type(exc).__name__)
        else:
            from ddgs import DDGS
            # Explicit backends are independent attempts. Stop once useful hits
            # have been returned; avoid 5-6 searches for each document type.
            for backend in ('bing','brave'):
                try:
                    with DDGS(timeout=7) as ddgs:
                        items=ddgs.text(q,region='ru-ru',safesearch='moderate',
                                       max_results=12,backend=backend) or []
                    had_success=True
                    if items:break
                except Exception as exc:
                    problems.append(type(exc).__name__)
                    log.warning('Search provider %s failed: %s',backend,type(exc).__name__)
        for v in items:
            url=v.get('href') or v.get('url')
            if public_url(url or ''):
                found.append({'url':url,'title':norm_space(v.get('title','')),
                              'snippet':norm_space(v.get('body',''))})
        if found:break
    unique={}
    for v in found:unique.setdefault(v['url'].split('#')[0],v)
    warnings=[]
    if not had_success:
        warnings.append('Поисковые источники не ответили. Это сбой интернет-поиска, '
                        'а не подтверждение отсутствия документов. Попробуйте кнопки '
                        '«Поиск через браузер» или позже повторите запрос.')
    elif problems and not unique:
        warnings.append('Поиск выполнен не всеми источниками; подходящие ссылки не получены.')
    return list(unique.values()),warnings


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
    except Exception:return []


def pertinent(text,material):
    # Only words materially relevant to the product, never borrow the query as a PDF field.
    qwords=[w[:4] for w in re.findall(r'[а-яa-z0-9]+',material.lower()) if len(w)>=5]
    if not qwords:return True
    relevant=sum(1 for w in set(qwords) if w in text.lower())
    return relevant>=min(2,len(set(qwords)))


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
    queries=search_queries(material,dtype,manufacturer,standard)
    hits,warnings=search_hits(queries)
    direct=[x for x in hits if re.search(r'\.pdf(?:[?#]|$)',x['url'],re.I)]
    other=[x for x in hits if x not in direct][:8]
    with ThreadPoolExecutor(max_workers=3) as pool:
        futures=[pool.submit(expand_page,h) for h in other]
        for f in as_completed(futures):
            try:direct.extend(f.result())
            except Exception:pass
    unique={}
    for x in direct:unique.setdefault(x['url'].split('#')[0],x)
    # Limit transient PDF memory: a free Render instance has only 512 MB.
    candidates=list(unique.values())[:12]
    items=[]; unreadable=0; invalid=0; inappropriate=0
    # Test the actual bytes. Only return a link if a PDF was fully downloaded and parsed.
    def probe(hit):
        try:
            data,kind,actual=fetch(hit['url'],max_bytes=6*1024*1024,timeout=13)
            if not data.lstrip()[:1024].startswith(b'%PDF-'):return None,'not_pdf'
            label=PurePosixPath(urlsplit(actual).path).name or 'document.pdf'
            info=inspect_document(data,material,label,use_ocr=False)
            return (hit,data,info,actual),None
        except Exception as exc:
            log.info('Candidate inaccessible %s (%s)',urlsplit(hit['url']).hostname,type(exc).__name__)
            return None,'broken'
    checked=[]
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures=[pool.submit(probe,h) for h in candidates]
        for f in as_completed(futures):
            item,problem=f.result()
            if item:checked.append(item)
            elif problem=='broken':unreadable+=1
            else:invalid+=1
    # Search candidates must be real PDFs and actual quality documents.
    # OCR is local and serial; no external AI usage or charge.
    ocr_budget = max(0,min(6,int(os.getenv('OCR_SEARCH_LIMIT','3'))))
    checked.sort(key=lambda item:(bool(item[2].get('type')),
                     'сертификат' in item[0]['title'].lower(),
                     'паспорт' in item[0]['title'].lower()),reverse=True)
    ocr_count = 0
    for hit,data,info,actual in checked:
        filename=info['filename']
        # Explicitly prevent catalogues from being presented as quality docs.
        if re.search(r'(?i)(?:katalog|catalog|каталог|pricelist|price[-_]?list)',filename):
            inappropriate+=1
            continue
        if not info['type'] or not info['material']:
            if ocr_budget<=0:
                inappropriate+=1
                continue
            ocr_budget-=1
            ocr_count+=1
            try:
                info=inspect_document(data,material,filename,use_ocr=True)
            except RecognitionError as exc:
                warnings.append('OCR недоступен: '+str(exc)[:130])
                unreadable+=1
                continue
        if not info.get('type') or (dtype!='all' and info['type']!=dtype):
            inappropriate+=1
            continue
        if not info.get('material') or not pertinent(info.get('text','')+' '+info['material'],material):
            inappropriate+=1
            continue
        if manufacturer and info['manufacturer'] not in ('—','') and manufacturer.lower() not in info['manufacturer'].lower():
            inappropriate+=1
            continue
        fid=file_id_and_save(data)
        items.append(material_doc_record(info,fid,actual))
        if len(items)>=15:break
    if ocr_count and ocr_budget==0 and len(checked)>ocr_count:
        warnings.append('Лимит OCR-проверок достигнут. Уточните материал или изготовителя для более точного поиска.')
    if unreadable:
        warnings.append(f'Скрыты недоступные PDF: {unreadable}.')
    if inappropriate:
        warnings.append(f'Исключены каталоги и документы другого вида/материала: {inappropriate}.')
    if invalid:
        warnings.append(f'Исключены ссылки без PDF: {invalid}.')
    return items,warnings

@app.get('/')
def index():return render_template('index.html',types=DOC_SELECT)

@app.get('/health')
def health():return jsonify({'ok':True,'version':'quality-docs-v8.3-search-fallback','recognition':recognition_status()})

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
        results,warnings=perform_search(material,dtype,manufacturer,standard)
    except Exception:
        log.exception('search failed')
        return jsonify({'error':'Поиск не выполнен из-за ошибки сервера. Повторите попытку.'}),503
    finally:job_semaphore.release()
    return jsonify({'results':results,'warnings':warnings,'search_status':('unavailable' if any('Поисковые источники не ответили' in w for w in warnings) else 'completed')})

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
    if record is None or time.monotonic()-record[1]>2400:
        return jsonify({'error':'Срок хранения временной копии истёк. Повторите поиск.'}),404
    return send_file(io.BytesIO(record[0]), mimetype='application/pdf',as_attachment=True,download_name='document.pdf')

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
