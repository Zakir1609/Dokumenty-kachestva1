"""Tests with local fixtures; no claim of testing third-party sites or actual Render."""
import sys, types, io
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
# Flask is not installed in the test sandbox, emulate decorator plumbing only.
try:
 import flask
except ImportError:
 class DummyApp:
  def __init__(self,*a,**kw):
   self.json=types.SimpleNamespace(ensure_ascii=True)
   self.config={}
  def __getattr__(self, name):
   if name in ('before_request','after_request','get','post'):
    return lambda *args,**kwargs: (args[0] if args and callable(args[0]) else (lambda fn:fn))
   raise AttributeError(name)
 fake=types.ModuleType('flask')
 fake.Flask=DummyApp
 for name in ('jsonify','render_template','request','send_file','Response'):
  setattr(fake,name,lambda *a,**k:None)
 sys.modules['flask']=fake
import fitz, requests
import app

checks=0
def ok(cond, msg):
 global checks
 assert cond,msg
 checks+=1

# Real PDF parser, embedded Russian text (not mocked OCR)
doc=fitz.open();page=doc.new_page()
font='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
page.insert_font(fontname='Cyr',fontfile=font)
page.insert_text((45,60),'СЕРТИФИКАТ СООТВЕТСТВИЯ',fontname='Cyr',fontsize=14)
page.insert_text((45,86),'Продукция: Арматура А500С 18 мм',fontname='Cyr',fontsize=11)
page.insert_text((45,107),'Сертификат соответствия № RU.TEST.2026.12345',fontname='Cyr',fontsize=11)
page.insert_text((45,128),'Дата выдачи: 20.05.2026',fontname='Cyr',fontsize=11)
page.insert_text((45,149),'Изготовитель: ООО «МеталлПром»',fontname='Cyr',fontsize=11)
pdf=doc.tobytes();doc.close()
from extractor import inspect_bytes
info=inspect_bytes(pdf,'Арматура А500 Ø18 мм','cert.pdf')
ok(info['type']=='certificate',f'type={info["type"]}')
ok('Арматура' in info['material'],f'material={info["material"]}')
ok('20.05.2026' in info['date'],f'date={info["date"]}')
ok('МеталлПром' in info['manufacturer'],f'manufacturer={info["manufacturer"]}')

h1={'url':'https://example.org/cert.pdf','title':'Арматура А500С сертификат','snippet':'арматура а500'}
h2={'url':'https://example.org/restricted.pdf','title':'Арматура сертификат','snippet':'арматура'}
h3={'url':'https://example.org/missing.pdf','title':'Арматура сертификат','snippet':'арматура'}
h4={'url':'https://example.org/slow.pdf','title':'Арматура сертификат','snippet':'арматура'}
original_fetch=app.fetch
original_hits=app.search_hits
original_expand=app.expand_page

def failure(code):
 response=requests.Response();response.status_code=code;response.url='https://example.org/'+str(code)
 return requests.HTTPError(f'HTTP {code}',response=response)

def fake_fetch(url,**kw):
 if url==h1['url']:return pdf,'application/pdf',url
 if url==h2['url']:raise failure(403)
 if url==h3['url']:raise failure(404)
 if url==h4['url']:raise requests.ConnectTimeout('simulated')
 raise AssertionError(url)
app.search_hits=lambda q:([h1,h2,h3,h4],[])
app.expand_page=lambda h:[]
app.fetch=fake_fetch
items,warnings,sources,stats=app.perform_search('Арматура А500 Ø18 мм','certificate','','')
ok(len(items)==1,('verified items',items,warnings))
ok(items[0]['material'].startswith('Арматура'),items)
ok(len(sources)==2,('sources',sources))
ok(h3['url'] not in [x['url'] for x in sources], 'HTTP 404 must be hidden')
ok(h2['url'] in [x['url'] for x in sources], 'HTTP 403 uncertain should be visible')
ok(h4['url'] in [x['url'] for x in sources], 'Timeout uncertain should be visible')
ok(stats['pdf_candidates']==4,stats)
ok('удалённая ссылка' in '\n'.join(warnings),warnings)

# Temporary PDF copy is stored on disk (not permanently in RAM).
import os
fileid=app.file_id_and_save(pdf)
filepath=app.file_cache[fileid][0]
ok(os.path.isfile(filepath),filepath)
ok(Path(filepath).read_bytes()==pdf,'temp PDF differs from input')

# The screenshot case: 12 search hits but zero downloaded files must NOT hide all sources.
app.search_hits=lambda q:([dict(h4,url=f'https://example.org/slow{i}.pdf') for i in range(12)],[])
app.fetch=lambda url,**kw: (_ for _ in ()).throw(requests.ConnectTimeout('simulated'))
items,warnings,sources,stats=app.perform_search('Арматура А500 Ø18 мм','all','','')
ok(not items,items)
ok(len(sources)==12,('all source candidates',len(sources)))
ok(stats['unverified_sources']==12,stats)
ok(any('сервер не подтвердил реквизиты' in w for w in warnings),warnings)

# Sources and confirmed documents are separate; only four record fields go to Excel.
ok(app.link_problem(failure(410))=='gone','410 exclude')
ok(app.link_problem(failure(429))=='restricted','429 unknown access')

# html failures are discarded; successful empty page can still be listed
app.expand_page=original_expand
app.search_hits=original_hits
app.fetch=original_fetch

# Two distinct search queries, broad should omit d18.
qs=app.search_queries('Арматура А500 Ø18 мм','certificate','','')
ok(len(qs)==2,qs)
ok(qs[0]!=qs[1],qs)
ok('18' not in qs[1],qs)

# Search both variants; do not abort after the first result.
calls=[]
class FakeDDGS:
 def __init__(self,timeout):pass
 def __enter__(self):return self
 def __exit__(self,*args):pass
 def text(self,q,**kw):
  calls.append(q)
  if '18' in q:
   return [{'href':h1['url'],'title':'Арматура','body':'Сертификат'}]
  return [{'href':h2['url'],'title':'Арматура','body':'Сертификат'}]
sys.modules['ddgs']=types.SimpleNamespace(DDGS=FakeDDGS)
import os
os.environ.pop('BRAVE_SEARCH_API_KEY',None)
hits,warns=app.search_hits(qs)
ok(len(hits)==2,(hits,warns))
ok(len(calls)==2,calls)

# Template must explicitly separate confirmed records and uncertain URLs.
html=(ROOT/'templates/index.html').read_text(encoding='utf-8')
ok('source-list' in html and 'drawSources()' in html,'missing candidate UI')
ok('Номер документа' in html and 'Производитель' in html, 'four-field UI')
print(f'PASS: {checks} offline checks; actual network / Render remains unverified')
