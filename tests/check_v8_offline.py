"""Local offline regression with a TEST DOUBLE, not a real RapidOCR quality test."""
import io,sys,re,types
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import fitz
from PIL import Image,ImageDraw,ImageFont
import local_ocr, extractor
FONT='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
lines=[
 'СЕРТИФИКАТ СООТВЕТСТВИЯ',
 'Регистрационный номер: 12345-26',
 'Срок действия с 20.05.2026 по 19.05.2028',
 'Продукция: Арматура стальная горячекатаная класса А500С, диаметром 18 мм',
 'Изготовитель: ООО «МеталлПром»',
]
img=Image.new('RGB',(1800,1300),'white'); d=ImageDraw.Draw(img);font=ImageFont.truetype(FONT,28)
for i,line in enumerate(lines):d.text((30,50+i*67),line,fill='black',font=font)
b=io.BytesIO();img.save(b,format='PNG')
doc=fitz.open();page=doc.new_page(width=595,height=842);page.insert_image(page.rect,stream=b.getvalue());pdf=doc.tobytes();doc.close()
assert len(extractor.inspect_bytes(pdf)['text'])<5
mock=types.SimpleNamespace(txts=lines)
class FakeOCR:
 def __call__(self,image):
  assert isinstance(image,Image.Image)
  assert image.width>=500
  return mock
local_ocr._engine=FakeOCR()
info=local_ocr.inspect_document(pdf,'Другая арматура','scan.pdf')
assert info['type']=='certificate',info
assert info['number']=='12345-26',info
assert info['date']=='с 20.05.2026г. по 19.05.2028г.',info
assert info['manufacturer']=='ООО «МеталлПром»',info
assert 'А500С' in info['material'],info
assert 'Другая' not in info['material'],info
assert info['ocr_used'],info
assert info['pages']==1
plain=fitz.open();p=plain.new_page();p.insert_text((30,60),'hello world');other=plain.tobytes();plain.close()
try:
 local_ocr.inspect_document(b'not a pdf')
 raise AssertionError('invalid PDF accepted')
except ValueError:pass
no_number='\n'.join(l for l in lines if 'Регистрационный' not in l)
assert local_ocr._parsed(no_number,'','scan.pdf')['number']=='б/н'
assert local_ocr._parsed('КАТАЛОГ ПРОДУКЦИИ\n'+'\n'.join(lines),'','catalog.pdf')['type']==''
print('PASS: 11 offline regression checks, OCR engine simulated; no claim of real RapidOCR inference')
# Browser API regression with Flask test doubles (Flask is not installed in this build environment).
flask=types.ModuleType('flask')
class FakeApp:
 def __init__(self,*a,**k):self.json=types.SimpleNamespace(ensure_ascii=True);self.config={}
 def before_request(self,f):return f
 def after_request(self,f):return f
 def get(self,path):return lambda f:f
 def post(self,path):return lambda f:f
 def run(self,**kwargs):pass
flask.Flask=FakeApp
flask.jsonify=lambda x:x
flask.render_template=lambda *args,**kw:'HTML'
flask.Response=lambda *args,**kw:('Response',args,kw)
flask.send_file=lambda f,**kw:('FILE',f.getvalue(),kw)
flask.request=types.SimpleNamespace(files={},form={},get_json=lambda **kw:{})
sys.modules['flask']=flask
import app
app.search_hits=lambda qs:([{'url':'https://public.example/cert.pdf','title':'Сертификат','snippet':''},
 {'url':'https://public.example/catalog.pdf','title':'Каталог','snippet':''},
 {'url':'https://public.example/broken.pdf','title':'Broken','snippet':''}],[])
app.fetch=lambda url,**kw: ((pdf,'application/pdf',url) if url.endswith('cert.pdf') else
                                  (pdf,'application/pdf',url) if url.endswith('catalog.pdf') else
                                  (_ for _ in ()).throw(ValueError('404')))
res,warns=app.perform_search('Арматура стальная А500С диаметром 18 мм','certificate','','')
assert len(res)==1,(res,warns)
assert res[0]['number']=='12345-26'
assert res[0]['manufacturer']=='ООО «МеталлПром»'
assert app.api_file(res[0]['id'])[1]==pdf
from openpyxl import load_workbook
flask.request.get_json=lambda **kw:{'results':res}
wb=load_workbook(io.BytesIO(app.api_export()[1]),read_only=True)
assert [c.value for c in wb.active[1]]==['Наименование материала','Номер документа','Дата документа','Производитель']
assert wb.active['D2'].value=='ООО «МеталлПром»'
class Upload:
 filename='scan.pdf'
 stream=io.BytesIO(pdf)
flask.request.files={'file':Upload()}
flask.request.form={'material':'Арматура А500С'}
record=app.api_upload()['result']
assert record['number']=='12345-26'
assert record['date']=='с 20.05.2026г. по 19.05.2028г.'
assert 'Gemini' not in repr(app.health())
print('PASS: 10 additional API/search/Excel checks using OCR test double (real model not run)')
