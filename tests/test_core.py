"""Offline unit tests. Run after installing requirements: python -m unittest discover -s tests"""
import unittest
from unittest.mock import patch

import app


class CoreTests(unittest.TestCase):
    def test_class_and_diameter(self):
        self.assertEqual(app.material_clues('Арматура А500 Ø18 мм'), ('A500', 18))
        self.assertEqual(app.material_clues('Арматура А500С d18мм'), ('A500C', 18))
        self.assertIsNone(app.material_clues('Бетон В25')[1])

    def test_document_types(self):
        self.assertEqual(app.guess_type('Сертификат соответствия'), 'certificate')
        self.assertEqual(app.guess_type('Декларация о соответствии'), 'declaration')
        self.assertEqual(app.guess_type('Паспорт качества'), 'passport')

    def test_document_attributes(self):
        text = ('СЕРТИФИКАТ СООТВЕТСТВИЯ\n'
                'Регистрационный номер: RU.НИЦ.001.003.00052\n'
                'Дата выдачи: 24.03.2026\n'
                'Изготовитель: ПАО «Северсталь»\n'
                'Наименование продукции: Арматура А500 диаметром 18 мм\n'
                'ГОСТ 34028-2016')
        self.assertEqual(app.extract_date(text), '24.03.2026')
        self.assertEqual(app.extract_manufacturer(text), 'ПАО «Северсталь»')
        self.assertEqual(app.extract_number(text), 'RU.НИЦ.001.003.00052')
        self.assertEqual(app.extract_product(text), 'Арматура А500 диаметром 18 мм')
        self.assertEqual(app.extract_standard(text), 'ГОСТ 34028-2016')

    def test_conservative_match(self):
        m = 'Арматура А500 Ø18 мм'
        self.assertEqual(app.assess_match(m, 'Арматура А500 диаметром 18 мм', '', True)[0],
                         'Совпадение признаков')
        self.assertEqual(app.assess_match(m, 'Арматура А400 диаметром 18 мм', '', True)[0],
                         'Возможное расхождение')
        self.assertEqual(app.assess_match(m, 'Арматура А500С диаметром 18 мм', '', True)[0],
                         'Возможное расхождение')
        self.assertEqual(app.assess_match(m, '', 'Арматура А500 18 мм', False)[0],
                         'Требует проверки')

    def test_query_composition(self):
        queries = app.create_queries('Арматура А500 Ø18 мм', 'certificate', '', '')
        self.assertTrue(any('filetype:pdf' in q for q in queries))
        self.assertTrue(all('сертификат соответствия' in q for q in queries))

    def test_exclude_local_url(self):
        self.assertFalse(app.safe_http_url('file:///tmp/test.pdf'))
        self.assertFalse(app.safe_http_url('http://localhost:80/test'))
        self.assertFalse(app.safe_http_url('http://127.0.0.1/test'))
        self.assertTrue(app.safe_http_url('https://example.com/test.pdf'))


if __name__ == '__main__':
    unittest.main()
