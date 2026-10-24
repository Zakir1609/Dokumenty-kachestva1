"""Offline smoke tests; uses lightweight import stubs if Flask/ddgs aren't installed.
Runs genuine local Tesseract + PyMuPDF on generated Russian image-only PDF.
Network search results are mocked so link filtering is deterministic.
"""
from __future__ import annotations
import io, sys, os, types, importlib, tempfile
from pathlib import Path
from unittest.mock import patch
import fitz
from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(HERE))
try:
    import flask  # noqa
except ImportError:
    dummy = types.ModuleType('flask')
    class DummyFlask:
        def __init__(self, name):
            self.json = types.SimpleNamespace(ensure_ascii=False)
            self.config = {}
            self.secret_key = 'test-secret'
        def before_request(self, fun): return fun
        def after_request(self, fun): return fun
        def route(self, *args, **kwargs): return lambda fun: fun
    dummy.Flask = DummyFlask
    dummy.jsonify = lambda obj: obj
    dummy.render_template = lambda *a, **kw: None
    dummy.request = types.SimpleNamespace()
    dummy.send_file = lambda f, **kw: f
    dummy.Response = lambda *a, **kw: None
    sys.modules['flask'] = dummy
try:
    import ddgs  # noqa
except ImportError:
    dummy = types.ModuleType('ddgs')
    class DummyDDGS:
        def __init__(self,*a,**kw): pass
        def __enter__(self): return self
        def __exit__(self,*a): return False
        def text(self,*a,**kw): return []
    dummy.DDGS = DummyDDGS
    sys.modules['ddgs'] = dummy

import app

TEXT = '''СЕРТИФИКАТ СООТВЕТСТВИЯ
Регистрационный номер: RU.TEST.181.357.36864
Дата выдачи: 24.03.2026
Действителен до: 24.03.2029
Изготовитель: ООО «СтальТест»
Наименование продукции: Арматура А500С диаметром 18 мм
ГОСТ 34028-2016
'''


def make_text_pdf():
    pdf=fitz.open()
    page=pdf.new_page(width=595,height=842)
    font_path='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
    page.insert_font(fontname='DV',fontfile=font_path)
    for i,line in enumerate(TEXT.splitlines()):
        page.insert_text((45,80+i*35),line,fontname='DV',fontsize=14)
    data=pdf.tobytes()
    pdf.close()
    return data


def make_scanned_pdf():
    image=Image.new('RGB',(1500,2100),'white')
    draw=ImageDraw.Draw(image)
    font=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',42)
    lines = TEXT.splitlines()
    for i,line in enumerate(lines):
        draw.text((65,100+125*i),line,font=font,fill='black')
    bio=io.BytesIO()
    image.save(bio,format='PNG',optimize=True)
    pdf=fitz.open()
    page=pdf.new_page(width=595,height=842)
    page.insert_image(page.rect,stream=bio.getvalue())
    data=pdf.tobytes()
    pdf.close()
    return data


def check(name, condition):
    if not condition: raise AssertionError(name)
    print('PASS',name)


def main():
    check('Tesseract rus ready and self-test', app.tesseract_status()['ready'])
    txt=make_text_pdf()
    a=app.parse_pdf_bytes(txt, 'Арматура А500С Ø18 мм', 'https://example.com/text.pdf')
    check('genuine text PDF opened',a['pdf_verified'] and a['pdf_read'] and a['pdf_pages']==1)
    check('text PDF manufacturer', 'СтальТест' in a['manufacturer'])
    check('text PDF issue date',a['date']=='24.03.2026')
    check('text PDF expiration',a['valid_until']=='24.03.2029')
    check('text PDF standard',a['gost_in_document']=='ГОСТ 34028-2016')
    check('text PDF document type',a['detected_type']=='certificate')
    scan=make_scanned_pdf()
    b=app.parse_pdf_bytes(scan, 'Арматура А500С Ø18 мм', 'https://example.com/scan.pdf')
    print('OCR status:',b['ocr_status'])
    print('OCR excerpt:',repr(b['text'][:350]))
    check('genuine image-only scanned PDF',b['pdf_scanned'] and b['pdf_verified'])
    check('scan OCR recognized Russian text',b['ocr_used'] and 'Изготовитель' in b['text'])
    check('scan OCR product class', '500' in b['text'])
    check('scan OCR manufacturer extracted', bool(b['manufacturer']))
    check('scan OCR date extracted', b['date']=='24.03.2026')
    check('scan OCR GOST extracted',b['gost_in_document']=='ГОСТ 34028-2016')
    try:
        app.parse_pdf_bytes(b'<html>Not a pdf</html>', 'test')
    except ValueError:
        check('non PDF rejected',True)
    else: raise AssertionError('non PDF not rejected')
    good = 'https://example.com/ok.pdf'
    dead = 'https://example.com/dead.pdf'
    def fake_download(url, material, enable_ocr):
        if url == dead:
            raise ValueError('HTTP 404')
        return app.parse_pdf_bytes(txt, material, url, enable_ocr=False)
    class FakeSearch:
        def __init__(self,**kw): pass
        def __enter__(self): return self
        def __exit__(self,*a): return False
        def text(self,*a,**kw):
            return [{'href':dead,'title':'Сертификат А500С','body':'арматура'},
                    {'href':good,'title':'Сертификат соответствия А500С','body':'арматура'}]
    with patch.object(app,'DDGS',FakeSearch),patch.object(app,'parse_pdf',side_effect=fake_download),\
         patch.object(app,'expand_page',return_value=[]), patch.dict(os.environ, {'BRAVE_SEARCH_API_KEY':''}):
        results, errors, raw_count=app.perform_search('Арматура А500С Ø18 мм','certificate','','',20)
    check('valid pdf is shown',len(results)==1 and results[0]['url']==good)
    check('broken link discarded',all(x['url']!=dead for x in results))
    check('validated search material information',results[0]['manufacturer'].startswith('ООО'))
    # Exercise the actual API route functions with mock request objects. A full
    # Flask HTTP test runs in deployment using tests/test_api.py.
    from openpyxl import load_workbook
    app.request.files = {'file': types.SimpleNamespace(filename='scan.pdf', stream=io.BytesIO(scan))}
    app.request.form = {'material': 'Арматура А500С Ø18 мм'}
    uploaded=app.upload_pdf_api()
    check('upload endpoint processed scan', bool(uploaded['metadata']['manufacturer']))
    check('upload endpoint verified PDF', uploaded['metadata']['pdf_verified'])
    check('upload endpoint status', uploaded['metadata']['status'] in ('Совпадение признаков', 'Требует проверки'))
    token = app.signer.dumps(good)
    app.request.get_json = lambda silent=True: {'url':good,'inspect_token':token,'material':'Арматура А500С Ø18 мм'}
    with patch.object(app,'parse_pdf',return_value=app.parse_pdf_bytes(scan, 'Арматура А500С Ø18 мм', good)):
        inspected=app.inspect_api()
    check('repeat-inspect endpoint OCR', bool(inspected['metadata']['manufacturer']))
    app.request.get_json=lambda silent=True: {'results':[uploaded['metadata']]}
    result_file=app.export()
    book=load_workbook(io.BytesIO(result_file.getvalue()))
    check('Excel export saved manufacturer', 'СтальТест' in book.active['M2'].value)
    app.request.args={'url':good,'token':token}
    with patch.object(app,'safe_request', return_value=(txt,'application/pdf',good)):
        downloaded=app.download_checked_pdf()
    check('PDF proxy returns actual PDF', downloaded.getvalue()[:5]==b'%PDF-')
    print('COMPLETE',len(results),'working PDF in mocked search. Warnings:',errors)

if __name__=='__main__':
    main()
