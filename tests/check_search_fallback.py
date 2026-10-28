"""Offline checks for search outage handling; no real Render/web calls."""
import os
import runpy
import sys
import types
from pathlib import Path
root=Path(__file__).resolve().parents[1]
os.environ.pop('BRAVE_SEARCH_API_KEY',None)
runpy.run_path(str(root/'tests'/'check_v8_offline.py'))
import app
queries=app.search_queries('Прожектор светодиодный 30Вт','certificate','','')
assert 1<=len(queries)<=2,queries
assert any('сертификат' in q for q in queries)

class FakeDDGS:
    backend_calls=[]
    def __init__(self,timeout):assert timeout<=7
    def __enter__(self):return self
    def __exit__(self,*a):pass
    def text(self,query,**kwargs):
        self.backend_calls.append(kwargs['backend'])
        if kwargs['backend']=='bing':raise TimeoutError('simulated')
        return [{'href':'https://example.org/cert.pdf','title':'Прожектор','body':'Сертификат'}]
sys.modules['ddgs']=types.SimpleNamespace(DDGS=FakeDDGS)
# Undo the mock from the prior regression test.
import importlib
app=importlib.reload(app)
hits,warnings=app.search_hits(queries)
assert hits and hits[0]['url']=='https://example.org/cert.pdf',hits
assert FakeDDGS.backend_calls==['bing','brave'],FakeDDGS.backend_calls
assert not warnings,warnings

class AllTimeout(FakeDDGS):
    def text(self,query,**kwargs):raise TimeoutError('simulated')
sys.modules['ddgs']=types.SimpleNamespace(DDGS=AllTimeout)
hits,warnings=app.search_hits(queries)
assert not hits
assert len(warnings)==1 and 'Поисковые источники не ответили' in warnings[0],warnings
html=(root/'templates/index.html').read_text()
assert 'yandex-link' in html and 'google-link' in html
assert "data.search_status==='unavailable'" in html
assert 'Поисковая система не ответила. Результат поиска неизвестен' in html
print('PASS: 8 offline search tests; real provider availability not tested')
