"""Integration tests using Flask's HTTP test client; no external network calls."""
import io
import os
import unittest
from unittest.mock import patch
from openpyxl import load_workbook
import app


class HttpApiTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()

    def test_health(self):
        res = self.client.get('/health')
        self.assertEqual(res.status_code, 200)
        self.assertTrue(res.json['ok'])

    def test_upload_no_material_required(self):
        fake = dict(manufacturer='ООО «Испытания»', material_in_document='Арматура А500 Ø18 мм',
                    gost_in_document='ГОСТ 34028-2016', date='24.03.2026', number='RU.TEST.01',
                    valid_until='24.03.2029', detected_type='certificate', pdf_read=True,
                    pdf_scanned=False, ocr_used=False, ocr_status='', pdf_pages=1,
                    pdf_bytes=500, pdf_filename='', pdf_title='', text='Арматура А500 Ø18 мм')
        with patch.object(app, 'parse_pdf_bytes', return_value=fake), patch.dict(os.environ, {'SITE_PASSWORD': ''}):
            res = self.client.post('/api/upload', data={'file': (io.BytesIO(b'%PDF-\nfake'), 'test.pdf')},
                                   content_type='multipart/form-data')
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.json['metadata']['manufacturer'], 'ООО «Испытания»')
        self.assertEqual(res.json['metadata']['status'], 'Без сопоставления')
        self.assertEqual(res.json['metadata']['pdf_filename'], 'test.pdf')

    def test_repeat_read_keeps_signed_url(self):
        orig='https://example.com/original.pdf'
        token=app.signer.dumps(orig)
        fake=dict(detected_type='certificate', pdf_read=False, text='', resolved_url='https://cdn.example.com/new.pdf')
        with patch.object(app,'parse_pdf',return_value=fake), patch.dict(os.environ, {'SITE_PASSWORD': ''}):
            res=self.client.post('/api/inspect', json={'url':orig,'inspect_token':token,'material':'Арматура А500 Ø18 мм'})
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.json['metadata']['url'],orig)
        self.assertEqual(res.json['metadata']['status'],'PDF не проверен')

    def test_excel_export(self):
        with patch.dict(os.environ, {'SITE_PASSWORD': ''}):
            res=self.client.post('/api/export',json={'results':[{'requested_material':'Арматура А500 Ø18 мм',
                                 'document_number':'RU.TEST.01', 'manufacturer':'Проверочный завод',
                                 'read_error':'Удалённый сервер недоступен'}]})
        self.assertEqual(res.status_code, 200)
        workbook=load_workbook(io.BytesIO(res.data))
        sheet=workbook.active
        self.assertEqual(sheet['F2'].value,'RU.TEST.01')
        self.assertEqual(sheet['M2'].value,'Проверочный завод')
        self.assertEqual(sheet['L2'].value,'Удалённый сервер недоступен')

    def test_no_fabricated_match_on_failed_download(self):
        self.assertEqual(app.assess_match('Арматура А500 Ø18 мм', '',
                         'Сертификат Арматура А400 Ø16 мм',False)[0],'PDF не проверен')

if __name__=='__main__':
    unittest.main()
