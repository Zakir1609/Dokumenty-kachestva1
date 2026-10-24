"""Four-field extraction of machine-readable PDF text (no OCR or Tesseract).
Scanned images are read by free local RapidOCR; text fields are parsed here.
"""
from __future__ import annotations

import io
import math
import os
import re
import shutil
from datetime import datetime
from functools import lru_cache
from typing import Optional

import fitz

DOC_TYPES = {
    'certificate': 'Сертификат соответствия',
    'declaration': 'Декларация о соответствии',
    'passport': 'Паспорт качества',
    'test': 'Протокол испытаний',
    'refusal': 'Отказное письмо',
}

TYPE_PATS = {
    'certificate': re.compile(r'сертификат\s+(?:о\s+)?соответстви[яеию]', re.I),
    'declaration': re.compile(r'декларац(?:ия|ии|ию)\s+о\s+соответствии', re.I),
    'passport': re.compile(r'(?:паспорт\s+(?:качеств[ао]|изделия|продукци[ию])|сертификат\s+качеств[ао])', re.I),
    'test': re.compile(r'протокол\s+испытани[йя]', re.I),
    'refusal': re.compile(r'(?:отказн(?:ое|ого)\s+письм[оа]|информационное\s+письмо)', re.I),
}

DATE_RE = r'(?<!\d)([0-3]?\d[.\-/][01]?\d[.\-/](?:20\d\d|19\d\d))(?!\d)'
DATE_WORD_RE = r'(?<!\d)([0-3]?\d\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+20\d\d)(?!\d)'
ORG_RE = re.compile(r'(?:(?:ООО|АО|ПАО|ОАО|ЗАО|ИП|ФГУП|ГУП)\s*[«"\']?[^\n]{4,100}?[»"\']?(?=\s*(?:,|;|\.|$)))', re.I)
LABEL_MANUFACTURER = re.compile(r'(?:предприятие\s*[-–]?\s*изготовитель|завод\s*[-–]?\s*изготовитель|наименование\s+(?:и\s+адрес\s+)?изготовителя|изготовитель(?:\s+продукции)?|производитель(?:\s+продукции)?)\s*[:—–-]?\s*(.*)$', re.I)
LABEL_PRODUCT = re.compile(r'^(?:наименование\s+(?:и\s+обозначение\s+)?(?:продукции|материала|изделия)|продукция(?:\s*\([^)]*\))?|объект\s+сертификации|наименование\s+продукции\s*,?\s*на\s+которую\s+выдан\s+сертификат|продукция\s+серийного\s+выпуска|материал)\s*[:—–-]?\s*(.*)$', re.I)
BREAK_LABEL = re.compile(r'^(?:адрес|место\s+нахождения|тел(?:ефон)?|заявитель|регистрационн|испытательн|орган\s+по|дата|сертификат|срок\s+действия|контакт|директор|руководитель|почтовый|соответствует|основание)\b', re.I)
IGNORE_CATALOG = re.compile(r'(?:каталог(?:\s+продукции|\s+изделий|\s+20\d\d)?|\bkatalog\w*|\bcatalog\w*|\bprice[-_]?list\b|прайс[- ]?лист|рекламный\s+буклет)', re.I)


def norm_space(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def pretty_date(value: str) -> str:
    if not value:
        return ''
    value = norm_space(value).lower().replace('/', '.').replace('-', '.')
    try:
        if re.fullmatch(r'\d{1,2}\.\d{1,2}\.\d{4}', value):
            dt = datetime.strptime(value, '%d.%m.%Y')
            return dt.strftime('%d.%m.%Y') + 'г.'
        months = {'января':1,'февраля':2,'марта':3,'апреля':4,'мая':5,'июня':6,
                  'июля':7,'августа':8,'сентября':9,'октября':10,'ноября':11,'декабря':12}
        m = re.fullmatch(r'(\d{1,2})\s+([а-я]+)\s+(\d{4})', value)
        if m and m.group(2) in months:
            return datetime(int(m.group(3)), months[m.group(2)], int(m.group(1))).strftime('%d.%m.%Y')+'г.'
    except ValueError:
        pass
    return ''


def find_dates(text: str) -> list[tuple[int, str]]:
    dates = []
    for regex in (DATE_RE, DATE_WORD_RE):
        for m in re.finditer(regex, text, re.I):
            parsed = pretty_date(m.group(1))
            if parsed:
                dates.append((m.start(), parsed))
    return sorted(dates)


def document_date(text: str, kind: str) -> str:
    lines = [norm_space(ln) for ln in text.splitlines()]
    # Period is only legal when text explicitly associates both boundaries.
    range_exprs = [
        rf'(?:срок\s+действия\s*[:—–-]?\s*|действител(?:ен|ьна|ьно)\s*[:—–-]?\s*)?с\s*{DATE_RE}\s*(?:г\.?\s*)?по\s*{DATE_RE}',
        rf'(?:срок\s+действия|действует|действителен)\s*[:—–-]?\s*{DATE_RE}\s*[-–—]\s*{DATE_RE}',
    ]
    for ln in lines:
        for pat in range_exprs:
            m = re.search(pat, ln, re.I)
            if m:
                a, b = pretty_date(m.group(1)), pretty_date(m.group(2))
                if a and b and a != b and datetime.strptime(a[:10], '%d.%m.%Y') < datetime.strptime(b[:10], '%d.%m.%Y'):
                    return f'с {a} по {b}'
    # Some certificate layouts split "с" and "по" over successive lines.
    flat = norm_space(text[:18000]); dates = find_dates(flat)
    for m in re.finditer(r'(?:срок\s+действия|действителен\s+с|действует\s+с|действителен)', flat, re.I):
        tail = flat[m.start():m.start()+180]
        pair = re.search(rf'\bс\s*{DATE_RE}\s*(?:г\.?\s*)?\bпо\s*{DATE_RE}', tail, re.I)
        if pair:
            a, b = pretty_date(pair.group(1)), pretty_date(pair.group(2))
            if a and b and datetime.strptime(a[:10], '%d.%m.%Y') <= datetime.strptime(b[:10], '%d.%m.%Y'):
                return f'с {a} по {b}'
    labels = re.compile(r'(?:дата\s+(?:выдачи|регистрации|составления|оформления|изготовления|документа)|выдан\s+|зарегистрирован\s+|оформлен\s+|\bот\s+)', re.I)
    # Only consider 1st-page header as first-choice; avoid "от" within normatives.
    scored = []
    for i, line in enumerate(lines[:100]):
        ds = find_dates(line)
        if not ds and i+1 < len(lines) and re.search(r'дата\s+(?:выдачи|регистрации|документа|составления)', line, re.I):
            ds = find_dates(lines[i+1])
        for _, date in ds:
            head = line[:max(0,line.find(date[:10]))]
            score = 0
            if re.search(r'дата\s+(?:выдачи|регистрации|составления|оформления|документа)', line, re.I): score = 90
            elif re.search(r'\b(?:выдан|зарегистрирован|оформлен)\b', line, re.I): score = 80
            elif re.search(r'\bот\s*'+re.escape(date[:10]), line, re.I): score = 50
            elif i < 15: score = 15
            if re.search(r'ГОСТ|ТУ\s*\d|ФЕДЕРАЛЬНЫЙ\s+ЗАКОН', line, re.I): score -= 55
            if re.search(r'(?:действителен|действует|окончани[ея]|по\s*\d)',line,re.I): score -= 40
            scored.append((score, -i, date))
    if scored:
        best = max(scored)
        if best[0] >= 15:
            return best[2]
    # Some certificate pdf layout unlabelled date follows its document ID.
    return ''


def document_number(text: str, kind: str) -> str:
    lines = [norm_space(t) for t in text.splitlines() if t.strip()]
    joined = '\n'.join(lines[:90])
    pats = [
        r'\b(?:регистрационный\s+(?:номер|№)|рег\.\s*№)\s*[:№Nno\s]*([^\n]{4,100})',
        r'(?:сертификат\s+(?:о\s+)?соответствия|декларация\s+о\s+соответствии|паспорт(?:\s+качества)?|протокол\s+испытаний|отказное\s+письмо)\s*(?:№|N[o°º]?|номер)\s*[:\s]*([^\n]{3,100})',
        r'(?:^|\n)\s*(?:№|N[o°º]?)\s*([A-ZА-ЯЁ0-9][^\n]{3,70})',
        r'\b(РОСС\s+RU\s*[СCАAДD]?[\s.\-–]*[А-ЯA-Z0-9./\-]{9,})',
        r'\b((?:ЕАЭС\s*)?(?:RU|РФ|ТС)\s*[СCДD]?[\s.\-–]*[А-ЯA-Z0-9./\-]{9,})',
    ]
    for pat in pats:
        for m in re.finditer(pat, joined, re.I | re.M):
            val = norm_space(m.group(1)).strip('№:;,. ')
            # Common layouts put the product description after ID on one line.
            val = re.split(r'\s{2,}|\s+(?:дата|от\s+\d{2}\.|выдан|изготовитель|продукция|действителен)\b',val,1,flags=re.I)[0]
            val = val.replace(' ', '') if re.match(r'^(?:РОСС|RU|ЕАЭС)',val,re.I) else val
            if (3 <= len(val) <= 75 and re.search(r'\d', val) and
                not re.match(r'(?:ГОСТ|ТУ|ОКПД|ТР\s*ТС)\b',val,re.I) and
                not re.match(r'^(?:20|19)\d\d',val)):
                return val
    return 'б/н'


def clean_value(value: str, *, maxlen=210) -> str:
    v=norm_space(value).strip(' :-–—;,')
    v=re.split(r'\s{2,}|\s+(?=(?:адрес|заявитель|изготовитель|регистрационн|орган\s+по|телефон|срок\s+действия|дата\s+выдачи)\s*[:—])',v,1,flags=re.I)[0]
    return v[:maxlen].strip(' :-–—;,')


def labelled_value(lines, regex, maxlen=210):
    for idx, ln in enumerate(lines):
        m = regex.search(ln)
        if not m:continue
        val=clean_value(m.group(1),maxlen=maxlen)
        if re.match(r'^(?:наименование|и\s+адрес|адрес|юридический\s+адрес)\b',val,re.I): val=''
        if len(val) >= 4 and not BREAK_LABEL.match(val):
            return val
        for follow in lines[idx+1:idx+4]:
            if BREAK_LABEL.match(follow):break
            vv=clean_value(follow,maxlen=maxlen)
            if len(vv)>=4:
                return vv
    return ''


def manufacturer(text: str) -> str:
    lines = [norm_space(ln) for ln in text.splitlines() if ln.strip()]
    item=labelled_value(lines[:150], LABEL_MANUFACTURER, maxlen=200)
    if item:
        return item
    # A small number of certificate layouts print "Manufacturer" instead.
    eng=re.compile(r'^(?:manufacturer|produced\s+by)\s*[:—-]?\s*(.*)$',re.I)
    item=labelled_value(lines[:100],eng,maxlen=200)
    if item:return item
    return ''


def material_name(text: str, material_hint: str = '') -> str:
    lines = [norm_space(ln) for ln in text.splitlines() if ln.strip()]
    item = labelled_value(lines[:220], LABEL_PRODUCT, maxlen=250)
    if item and not re.match(r'^(?:серийного\s+выпуска|зарегистрирован|соответствует|сертификат|наименование\s+и|нормативн|дата)\b',item,re.I):
        return item
    # Generic document name like "ПАСПОРТ КАЧЕСТВА\n Арматура ..."
    for i,ln in enumerate(lines[:45]):
        if TYPE_PATS['passport'].search(ln) and i+1<len(lines):
            next_line=lines[i+1]
            if len(next_line)>=8 and not re.match(r'^(?:№|от\s+\d|дата|изготовитель|партия|завод)',next_line,re.I):
                return clean_value(next_line,maxlen=230)
    # Fallback only to a real line on an actual PDF matching a distinctive query.
    def signature(x):
        out=[]
        for word in re.findall(r'[а-яa-z0-9]+',x.lower()):
            if len(word)>=5 and word not in ('короткая','короткий','размеров','партии','качества'):
                out.append(word[:4]);
            elif re.fullmatch(r'[аa]\s?500[сc]?',word):out.append('a500')
        return set(out)
    wants=signature(material_hint)
    if wants and len(wants)>=2:
        potential=[]
        for ln in lines:
            tokens=signature(ln)
            hits=len(tokens&wants)
            if hits>=min(2,len(wants)) and len(ln)>=8 and len(ln)<=250:
                bad=bool(re.search(r'(?i)(?<!\w)(?:www|http|email|заказ|доставка|цена|таблица\s+\d|оглавление)',ln))
                potential.append((hits,-len(ln),bad,ln))
        if potential:
            potential.sort(key=lambda x:(x[0],not x[2],x[1]),reverse=True)
            return clean_value(potential[0][3],maxlen=230)
    return ''


def doc_type(text: str, pdf_filename='') -> str:
    # Large catalogs may mention certificates on a back page. Detect real title.
    first = text[:4500]
    if IGNORE_CATALOG.search(pdf_filename):
        return ''
    first_lines = '\n'.join(text.splitlines()[:32])
    if IGNORE_CATALOG.search(first_lines[:900]) and not re.search(r'(?:^|\n)\s*(?:сертификат\s+соответствия|паспорт\s+качества|декларация\s+о\s+соответствии)\s*(?:№|$)',first_lines,re.I):
        return ''
    candidates=[]
    for kind,pat in TYPE_PATS.items():
        m=pat.search(first)
        if m: candidates.append((m.start(), kind))
    return min(candidates)[1] if candidates else ''


def inspect_bytes(data: bytes, material_hint: str='', filename: str='', *, use_ocr: bool=False) -> dict:
    """Read embedded PDF text without attempting OCR. The use_ocr flag is a legacy no-op."""
    if len(data)>25*1024*1024:
        raise ValueError('PDF больше 25 МБ')
    if not data[:1024].lstrip().startswith(b'%PDF-'):
        raise ValueError('Файл не является PDF')
    try:
        doc=fitz.open(stream=data,filetype='pdf')
    except Exception as exc:
        raise ValueError('PDF повреждён или не может быть открыт') from exc
    try:
        if doc.is_encrypted and not doc.authenticate(''):
            raise ValueError('PDF защищён паролем')
        pages=len(doc)
        if pages==0 or pages>200:
            raise ValueError('Число страниц PDF должно быть от 1 до 200')
        text='\n'.join(doc[i].get_text('text',sort=True)[:14000] for i in range(pages))
        quality=(len(text.strip())>40 and len(re.findall(r'[А-Яа-яЁёA-Za-z]',text))>=25)
        kind=doc_type(text,filename)
        return {
            'material':material_name(text,material_hint),
            'number':document_number(text,kind),
            'date':document_date(text,kind) or '—',
            'manufacturer':manufacturer(text) or '—',
            'type':kind,
            'pages':pages,
            'has_text':quality,
            'ocr_used':False,
            'ocr_note':'Текстовый слой PDF' if quality else 'Текстовый слой не найден',
            'filename':filename,
            'text':text[:140000],
        }
    finally:
        doc.close()
