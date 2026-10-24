"""Облачное веб-приложение для поиска документов качества.
Проверка реквизитов PDF не означает юридическую проверку документа или партии.
"""
from __future__ import annotations

import io
import os
import secrets
import hmac
import socket
import time
from collections import defaultdict, deque, OrderedDict
from functools import lru_cache
from threading import Lock, BoundedSemaphore
import shutil
import ipaddress
import logging
import math
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import PurePosixPath
from urllib.parse import unquote
from urllib.parse import urljoin, urlsplit, urlunsplit

import fitz
import requests
from bs4 import BeautifulSoup
from ddgs import DDGS
from flask import Flask, jsonify, render_template, request, send_file, Response
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired

app = Flask(__name__)
app.json.ensure_ascii = False
app.config['MAX_CONTENT_LENGTH'] = 27 * 1024 * 1024
app.config['JSON_SORT_KEYS'] = False
app.secret_key = os.environ.get('APP_SECRET_KEY') or secrets.token_hex(32)
signer = URLSafeTimedSerializer(app.secret_key, salt='pdf-inspect-url-v1')
request_jobs = BoundedSemaphore(2)
traffic = defaultdict(deque)
traffic_lock = Lock()
search_cache = OrderedDict()
cache_lock = Lock()
SEARCH_CACHE_TTL = 180
SEARCH_CACHE_SIZE = 24
logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

DOC_TYPES = {
    'all': 'Все виды документов',
    'certificate': 'Сертификат соответствия',
    'declaration': 'Декларация о соответствии',
    'passport': 'Паспорт качества',
    'test': 'Протокол испытаний',
    'refusal': 'Отказное письмо',
}
DOC_TERMS = {
    'certificate': ['сертификат соответствия', 'сертификаты соответствия'],
    'declaration': ['декларация о соответствии', 'декларации соответствия'],
    'passport': ['паспорт качества', 'сертификат качества', 'паспорт на продукцию'],
    'test': ['протокол испытаний', 'испытания продукции'],
    'refusal': ['отказное письмо', 'информационное письмо'],
}
DOC_PATTERNS = {
    'certificate': r'сертификат(?:ы|а|ов)?\s+(?:о\s+)?соответстви[яюеи]',
    'declaration': r'декларац(?:ия|ии|ию)\s+(?:о\s+)?соответстви[яюеи]',
    'passport': r'паспорт(?:ы|а)?\s+(?:качеств[ао]|продукци[ию])|сертификат(?:ы|а)?\s+качеств[ао]',
    'test': r'протокол(?:ы|а)?\s+испытани[йя]',
    'refusal': r'отказн(?:ое|ого)\s+письм[оа]',
}
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/pdf,text/html,application/xhtml+xml,*/*;q=0.8',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.6',
}
MAX_PDF_BYTES = 25 * 1024 * 1024
PDF_PREVIEW_PAGES = 6
MAX_OCR_PAGES = 3
OCR_PAGE_TIMEOUT = int(os.environ.get("OCR_PAGE_TIMEOUT", "75"))
OCR_DPI = int(os.environ.get("OCR_DPI", "170"))
MAX_SEARCH_QUERIES = 4
MAX_SEARCH_RESULTS = 36
MAX_PDFS_TO_READ = 16
MAX_HTML_PAGES_TO_EXPAND = 8


@app.before_request
def security_and_limits():
    # Optional browser-only password, configured in host's secret environment.
    password = os.environ.get('SITE_PASSWORD', '').strip()
    if password:
        auth = request.authorization
        if not auth or not hmac.compare_digest(auth.password or '', password):
            return Response('Для доступа к сайту введите пароль.', 401, {
                'WWW-Authenticate': 'Basic realm="Dokumenty kachestva", charset="UTF-8"',
                'Cache-Control': 'no-store',
            })

    policy = {'/api/search': 9, '/api/inspect': 15,
              '/api/upload': 15, '/api/export': 30, '/api/file': 20}
    if request.path not in policy or (request.method != 'POST' and not (request.path == '/api/file' and request.method == 'GET')):
        return None
    key = (request.remote_addr or 'unknown', request.path)
    now = time.monotonic()
    with traffic_lock:
        queue = traffic[key]
        while queue and now - queue[0] > 900:
            queue.popleft()
        if len(queue) >= policy[request.path]:
            return jsonify({'error': 'Превышен лимит запросов. Повторите попытку позднее.'}), 429
        queue.append(now)
        if len(traffic) > 3000:
            for old_key in list(traffic)[:300]:
                if not traffic[old_key] or now - traffic[old_key][-1] > 900:
                    del traffic[old_key]
    return None

@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['X-Frame-Options'] = 'DENY'
    if request.path.startswith('/api/'):
        response.headers['Cache-Control'] = 'no-store'
    return response


@app.route('/health')
def health():
    return jsonify({'ok': True, 'service': 'quality-docs-online-v5', 'ocr': tesseract_status()})


def assert_public_dns(url: str):
    """Best-effort anti-SSRF resolution check before outbound fetch."""
    parsed = urlsplit(url)
    if not safe_http_url(url) or parsed.username or parsed.password:
        raise ValueError('Непубличная ссылка запрещена')
    if parsed.port not in (None, 80, 443):
        raise ValueError('Недопустимый порт ресурса')
    hosts = socket.getaddrinfo(parsed.hostname, parsed.port or (443 if parsed.scheme == 'https' else 80), type=socket.SOCK_STREAM)
    if not hosts or any(not ipaddress.ip_address(item[4][0]).is_global for item in hosts):
        raise ValueError('Адрес ресурса не является общедоступным')


def run_limited_job(callback):
    if not request_jobs.acquire(blocking=False):
        return None
    try:
        return callback()
    finally:
        request_jobs.release()


def tidy(text: str, max_chars: int = 350) -> str:
    return re.sub(r'\s+', ' ', str(text or '')).strip()[:max_chars]


def safe_http_url(url: str) -> bool:
    """Restrict obvious local URLs; redirects are checked separately."""
    try:
        parsed = urlsplit(url)
        if parsed.scheme not in ('http', 'https') or not parsed.hostname:
            return False
        name = parsed.hostname.lower().rstrip('.')
        if name in {'localhost', 'localhost.localdomain'} or name.endswith(('.local', '.internal')):
            return False
        try:
            ip = ipaddress.ip_address(name)
            return ip.is_global
        except ValueError:
            return True
    except ValueError:
        return False


def error_for_download(exc: Exception) -> str:
    """Explain remote file failures without exposing Python stack traces."""
    if isinstance(exc, (requests.exceptions.ConnectTimeout, requests.exceptions.ReadTimeout, TimeoutError, socket.timeout)):
        return 'Удалённый сайт не ответил вовремя. Откройте PDF в браузере и загрузите его в форму на сайте.'
    if isinstance(exc, requests.exceptions.SSLError):
        return 'Ошибка защищённого соединения с сайтом документа (TLS). Можно загрузить PDF с компьютера.'
    if isinstance(exc, requests.exceptions.HTTPError):
        code = exc.response.status_code if exc.response is not None else None
        if code in (401, 403):
            return f'Сайт документа ограничил доступ (HTTP {code}). Попробуйте открыть PDF в браузере и загрузить его.'
        if code == 404:
            return 'Файл по ссылке не найден (HTTP 404).'
        return f'Сайт документа вернул HTTP {code or "ошибку"}. Проверьте ссылку.'
    if isinstance(exc, requests.exceptions.ConnectionError):
        return 'Сайт документа разорвал соединение или недоступен с сервера приложения. Откройте PDF в браузере и загрузите файл.'
    if isinstance(exc, (socket.gaierror, OSError)):
        return 'Не удалось установить соединение с сайтом документа.'
    return tidy(str(exc), 190) or 'Не удалось получить файл.'


def safe_request(url: str, *, max_bytes: int = MAX_PDF_BYTES, timeout: int = 10):
    """Bounded download with checked redirects. Network failure is not invalid certificate."""
    current = url
    for _ in range(5):
        assert_public_dns(current)
        response = None
        try:
            response = requests.get(current, timeout=(4, timeout), headers=HEADERS,
                                    stream=True, allow_redirects=False)
            if response.status_code in (301, 302, 303, 307, 308):
                target = urljoin(current, response.headers.get('Location', ''))
                current = target
                continue
            response.raise_for_status()
            content = bytearray()
            size_header = response.headers.get('Content-Length', '')
            if size_header.isdigit() and int(size_header) > max_bytes:
                raise ValueError(f'Файл превышает лимит {max_bytes // 1024 // 1024} МБ')
            for chunk in response.iter_content(64 * 1024):
                if not chunk:
                    continue
                content.extend(chunk)
                if len(content) > max_bytes:
                    raise ValueError(f'Файл превышает лимит {max_bytes // 1024 // 1024} МБ')
            return (bytes(content), response.headers.get('Content-Type', '').lower(), response.url)
        except (requests.exceptions.RequestException, socket.gaierror, socket.timeout, OSError) as exc:
            raise ValueError(error_for_download(exc)) from exc
        finally:
            if response is not None:
                response.close()
    raise ValueError('Слишком много перенаправлений')


def normalize_material(value: str) -> str:
    # Preserve original class designation; normalise typography only.
    s = re.sub(r'\s+', ' ', value or '').strip()
    return s.replace('Ø', 'ø').replace('⌀', 'ø')


def material_clues(material: str):
    normalized = material.upper().replace('А', 'A').replace('С', 'C')
    match = re.search(r'(?<![A-Z0-9])A\s*(400|500|550|600)\s*(C)?(?![A-Z0-9])', normalized)
    class_code = 'A' + match.group(1) + ('C' if match.group(2) else '') if match else ''
    diam = re.search(r'(?:[Øø⌀]|\b[dд]\s*=?\s*|диаметр(?:ом)?\s*|(?<!\d))\s*(\d{1,2})\s*(?:мм|mm)\b', material, re.I)
    diameter = None
    if diam:
        candidate = int(diam.group(1))
        if 3 <= candidate <= 80:
            diameter = candidate
    return class_code, diameter


def create_queries(material: str, dtype: str, manufacturer: str, gost: str) -> list[str]:
    term = normalize_material(material)
    extra = ' '.join(v for v in (manufacturer.strip(), gost.strip()) if v)
    groups = list(DOC_TERMS) if dtype == 'all' else [dtype]
    queries = []
    for group in groups:
        for keyword in DOC_TERMS[group][:1]:
            queries.append(f'{term} {keyword} {extra} filetype:pdf'.strip())
            if dtype != 'all':
                queries.append(f'{term} {keyword} {extra}'.strip())
    if dtype != 'all':
        queries.append(f'{term} {DOC_TERMS[dtype][0]} скачать pdf'.strip())
    else:
        queries.append(f'{term} документы качества pdf {extra}'.strip())
    return list(dict.fromkeys(queries))[:MAX_SEARCH_QUERIES]


def guess_type(text: str) -> str:
    source = str(text or '').lower()
    matches = [(m.start(), key) for key, regex in DOC_PATTERNS.items()
               if (m := re.search(regex, source, flags=re.I))]
    if matches:
        return sorted(matches)[0][1]
    return 'unknown'


def extract_manufacturer(text: str) -> str:
    lines = [re.sub(r'\s+', ' ', x).strip() for x in text.splitlines()]
    labels = r'(?:предприятие[- ]изготовитель|завод[- ]изготовитель|организация[- ]изготовитель|изготовитель|производитель)'
    for i, line in enumerate(lines):
        match = re.search(rf'(?<!\w){labels}\s*[:—–-]?\s*(.*)$', line, re.I)
        if not match:
            continue
        value = match.group(1).strip(' :;.,')
        if re.match(r'^(?:наименование|юридическое лицо|адрес|место нахождения)\b', value, re.I):
            value = ''
        if len(value) < 4:
            for nextline in lines[i+1:i+4]:
                if re.match(r'^(?:адрес|телефон|местонахождение|заявитель|продукция|сертификат|дата)\b', nextline, re.I):
                    break
                if len(nextline) >= 4:
                    value = nextline.strip(' :;.,')
                    break
        if 4 <= len(value) <= 170 and not re.match(r'^(?:адрес|наименование|местонахождени[ея])\b', value, re.I):
            return tidy(value, 170)
    return ''


def extract_date(text: str) -> str:
    """Find the issue/registration date. Never substitute PDF creation date."""
    patterns = [
        r'(?:дата\s+(?:выдачи|регистрации|оформления|составления|изготовления|выпуска|документа)|выдан(?:а|о)?|зарегистрирован[ао]?|выдано\s+от|дата\s+начала\s+действия)\s*[:№—–-]?\s*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'(?:дата\s+(?:выдачи|регистрации|составления)|выдан(?:а|о)?)\s*[:—–-]?\s*(\d{1,2}\s+[А-Яа-я]+\s+20\d{2}\s*(?:г\.)?)',
        r'(?:срок\s+действия\s*[:—–-]?\s*)?с\s*(\d{1,2}[./]\d{1,2}[./]\d{4})\s*(?:г\.?\s*)?по\s*\d{1,2}[./]\d{1,2}[./]\d{4}',
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if not m:
            continue
        value = tidy(m.group(1), 50)
        if re.fullmatch(r'\d{1,2}[./]\d{1,2}[./]\d{4}', value):
            try:
                datetime.strptime(value.replace('/', '.'), '%d.%m.%Y')
                return value.replace('/', '.')
            except ValueError:
                continue
        return value
    return ''


def extract_number(text: str) -> str:
    patterns = [
        r'(?:паспорт(?:\s+качества)?\s*(?:№|N[oº°]?)|протокол\s+испытаний\s*(?:№|N[oº°]?))\s*[:№\s]*([A-ZА-ЯЁ0-9][A-ZА-ЯЁ0-9.\-/]{2,70})',
        r'(?:регистрационн(?:ый|ого)\s+(?:номер|№)|рег\.\s*№|сертификат\s+соответствия\s*(?:№|N[oº°]?)|декларация\s+о\s+соответствии\s*(?:№|N[oº°]?))\s*[:№\s]*([А-ЯA-ZЁ0-9][А-ЯA-ZЁ0-9.\-/ ]{6,70})',
        r'\b((?:ЕАЭС\s*(?:N|№)\s*)?RU\s*[СCДDМM][\-\s.]?[A-ZА-Я0-9.\-/]{8,})',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            return tidy(m.group(1).strip(' .,:;'), 75)
    return ''



def extract_product(text: str) -> str:
    lines = [re.sub(r'\s+', ' ', x).strip() for x in text.splitlines()]
    for i, line in enumerate(lines):
        match = re.search(r'(?:наименование (?:и обозначение )?продукции|наименование изделия|сертифицируемая продукция|объект сертификации|^продукция|^изделие)\s*[:—–-]?\s*(.*)', line, re.I)
        if not match:
            continue
        value = match.group(1).strip(' :;.,')
        if len(value) < 4 and i + 1 < len(lines):
            value = lines[i + 1].strip(' :;.,')
        if 4 <= len(value) <= 200 and not re.match(r'^(изготовитель|адрес|производитель)\b', value, re.I):
            return tidy(value, 200)
    return ''


def extract_standard(text: str) -> str:
    m = re.search(r'\bГОСТ\s*(?:Р\s*)?\d{3,6}[—–-]\d{2,4}\b', text, re.I)
    if m:
        return tidy(m.group(0), 100)
    m = re.search(r'\bТУ\s*(?:\d+[.\-]){2,}\d+(?:[—–-]\d{2,4})?\b', text, re.I)
    return tidy(m.group(0), 100) if m else ''


def check_supplier_and_standard(manufacturer: str, gost: str, metadata: dict):
    observations=[]
    conflict=False
    documented=metadata.get('manufacturer', '')
    standard=metadata.get('gost_in_document', '')
    if manufacturer.strip():
        if documented:
            if manufacturer.casefold() in documented.casefold():
                observations.append('Изготовитель совпадает')
            else:
                conflict=True
                observations.append('Изготовитель отличается от указанного в запросе')
        else:
            observations.append('Изготовитель не установлен')
    if gost.strip():
        if standard:
            norm = lambda x: re.sub(r'\s+', '', x).replace('–','-').replace('—','-').upper()
            if norm(gost) == norm(standard):
                observations.append('Стандарт совпадает')
            else:
                observations.append('Нужна проверка применимого ГОСТ/ТУ')
        else:
            observations.append('Стандарт не извлечён из PDF')
    return conflict, observations



def assess_match(material: str, doc_text: str, title: str, pdf_read: bool):
    """Only assess PDF text; search results are NOT evidence of material conformity."""
    if not material.strip():
        return 'Без сопоставления', 'Не задан материал для проверки; реквизиты PDF можно просмотреть.'
    if not pdf_read or not doc_text.strip():
        return 'PDF не проверен', 'Текст документа не получен. Совпадение по заголовку поисковой выдачи не подтверждает материал.'
    requested_class, diameter = material_clues(material)
    hay = doc_text.upper().replace('А', 'A').replace('С', 'C')
    observations = []
    conflict = False
    score = 0
    if requested_class:
        present = re.findall(r'(?<![A-Z0-9])A\s*(?:400|500|550|600)\s*C?(?![A-Z0-9])', hay)
        present = {re.sub(r'\s+', '', v) for v in present}
        if requested_class in present:
            score += 2
            observations.append('Класс найден в тексте PDF')
        elif present:
            conflict = True
            observations.append('Обнаружен другой класс: ' + ', '.join(sorted(present)[:3]))
        else:
            observations.append('Класс в прочитанном тексте не найден')
    elif material.strip() and tidy(material, 30).lower() in doc_text.lower():
        score += 1
        observations.append('Название найдено в PDF')
    if diameter:
        exact_pattern = rf'(?:[øØ⌀]|(?:\bдиаметр\b|\bd\b|\bф\b)\s*[:=]?\s*)\s*{diameter}(?!\d)|(?<!\d){diameter}\s*(?:мм|mm)\b'
        exact = re.search(exact_pattern, hay, re.I)
        ranges = re.findall(r'(?<!\d)(\d{1,2})\s*[-–—]\s*(\d{1,2})\s*(?:мм|MM)\b', hay, re.I)
        in_range = any(int(a) <= diameter <= int(b) for a, b in ranges)
        if exact:
            score += 2
            observations.append(f'Диаметр {diameter} мм найден в PDF')
        elif in_range:
            score += 1
            observations.append(f'Диаметр {diameter} мм входит в диапазон, указанный в PDF')
        else:
            observations.append(f'Диаметр {diameter} мм в прочитанном тексте не подтверждён')
    if conflict:
        return 'Возможное расхождение', '; '.join(observations)
    if score >= (3 if requested_class and diameter else 2):
        return 'Совпадение признаков', '; '.join(observations)
    return 'Требует проверки', '; '.join(observations)



@lru_cache(maxsize=1)
def tesseract_status() -> dict:
    """Report actual OCR availability; test the executable rather than only its presence."""
    try:
        import pytesseract
        path = shutil.which('tesseract')
        if not path and os.name == 'nt':
            p = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            if os.path.isfile(p):
                path = p
        if not path:
            return {'ready': False, 'message': 'Не установлен Tesseract OCR на сервере'}
        pytesseract.pytesseract.tesseract_cmd = path
        langs = pytesseract.get_languages(config='')
        if 'rus' not in langs:
            return {'ready': False, 'message': 'Отсутствует пакет русского языка rus', 'languages': langs}
        # Verify that the binary can actually read a generated sample, not only
        # that the language pack is listed by Tesseract.
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new('L', (500, 120), 'white')
        drawer = ImageDraw.Draw(img)
        drawer.text((14, 24), 'TEST 123', font=ImageFont.load_default(32), fill='black')
        check = pytesseract.image_to_string(img, lang='rus+eng', config='--psm 6', timeout=15)
        if 'TEST' not in check.upper() and '123' not in check:
            return {'ready': False, 'message': 'OCR запущен, но не распознал контрольное изображение', 'languages': langs}
        return {'ready': True, 'message': 'Русский OCR готов · самопроверка пройдена', 'languages': langs}
    except Exception as exc:
        log.exception('OCR startup error')
        return {'ready': False, 'message': f'Ошибка запуска OCR: {type(exc).__name__}: {tidy(str(exc), 180)}'}


def ocr_error_message(exc: Exception) -> str:
    import pytesseract
    if isinstance(exc, RuntimeError) and 'timeout' in str(exc).lower():
        return f'Превышено время OCR ({OCR_PAGE_TIMEOUT} с на страницу). Скан слишком сложный или сервер перегружен.'
    if isinstance(exc, pytesseract.TesseractError):
        return 'Ошибка Tesseract: ' + tidy(str(exc), 220)
    return type(exc).__name__ + ': ' + tidy(str(exc), 200)


def ocr_pdf(doc, pages: int = MAX_OCR_PAGES) -> tuple[str, list[str]]:
    """OCR raster pages at readable DPI. Fallback to layout mode for sparse scans."""
    if not tesseract_status()['ready']:
        return '', [tesseract_status()['message']]
    import pytesseract
    from PIL import Image, ImageOps
    found = []
    errors = []
    for i in range(min(pages, len(doc))):
        try:
            page = doc.load_page(i)
            zoom = max(1.4, min(2.5, OCR_DPI / 72.0))
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csGRAY, alpha=False)
            # Guard against huge image memory usage on small hosting instances.
            if pix.width * pix.height > 9_000_000:
                factor = math.sqrt(9_000_000 / (pix.width * pix.height))
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom * factor, zoom * factor), colorspace=fitz.csGRAY, alpha=False)
            with Image.open(io.BytesIO(pix.tobytes('png'))) as loaded:
                img = loaded.convert('L')
                img = ImageOps.autocontrast(img, cutoff=1)
                recognized = pytesseract.image_to_string(img, lang='rus+eng', config='--oem 1 --psm 3', timeout=OCR_PAGE_TIMEOUT)
                if len(recognized.strip()) < 35:
                    try:
                        other = pytesseract.image_to_string(img, lang='rus+eng', config='--oem 1 --psm 6', timeout=OCR_PAGE_TIMEOUT)
                        if len(other.strip()) > len(recognized.strip()):
                            recognized = other
                    except Exception as exc:
                        log.warning('OCR fallback page %d: %s', i + 1, ocr_error_message(exc))
                found.append(recognized)
        except Exception as exc:
            message = f'Стр. {i + 1}: {ocr_error_message(exc)}'
            errors.append(message)
            log.warning('OCR page failed: %s', message)
    return '\n'.join(found), errors


def extract_valid_until(text: str) -> str:
    patterns = [
        r'(?:действител(?:ен|ьна|ьно)\s+до|действует\s+до|срок\s+действия\s+до|дата\s+окончания\s+действия|годен\s+до)\s*[:—–-]?\s*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'(?:срок\s+действия|действител(?:ен|ьна|ьно))\s*[:—–-]?\s*с\s*\d{1,2}[./]\d{1,2}[./]\d{4}\s*по\s*(\d{1,2}[./]\d{1,2}[./]\d{4})',
        r'(?:\bпо\b)\s*(\d{1,2}[./]\d{1,2}[./]\d{4})\s*(?:г\.?\b|$)',
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            value = m.group(1).replace('/', '.')
            try:
                datetime.strptime(value, '%d.%m.%Y')
                return value
            except ValueError:
                pass
    return ''


def extract_pdf_filename(url: str) -> str:
    try:
        candidate = unquote(PurePosixPath(urlsplit(url).path).name)
        return tidy(candidate, 170) if candidate else ''
    except Exception:
        return ''


def extract_properties(text: str) -> dict:
    return {
        'manufacturer': extract_manufacturer(text),
        'material_in_document': extract_product(text),
        'gost_in_document': extract_standard(text),
        'date': extract_date(text),
        'valid_until': extract_valid_until(text),
        'number': extract_number(text),
        'detected_type': guess_type(text[:8000]),
    }

def parse_pdf_bytes(data: bytes, material: str, url: str = '', *, enable_ocr: bool = True) -> dict:
    if not data[:1024].lstrip().startswith(b'%PDF-'):
        raise ValueError('Полученные данные не являются PDF-файлом')
    doc = fitz.open(stream=data, filetype='pdf')
    try:
        if len(doc) > 200:
            raise ValueError('PDF содержит более 200 страниц')
        if doc.is_encrypted and not doc.authenticate(''):
            raise ValueError('PDF защищён паролем')
        page_count = len(doc)
        parts = [doc.load_page(i).get_text('text', sort=True) for i in range(min(PDF_PREVIEW_PAGES, page_count))]
        extracted = '\n'.join(parts)[:130_000]
        document_text = extracted
        ocr_errors = []
        ocr_status = ''
        scanned = len(extracted.strip()) < 90
        # Text layers may be incomplete: scan pages if text doesn't contain any
        # verifiable document fields AND the text layer is very short.
        if (scanned or (len(extracted.strip()) < 350 and not any(v for k, v in extract_properties(extracted).items() if k != 'detected_type'))) and enable_ocr:
            status = tesseract_status()
            if status['ready']:
                recognized, ocr_errors = ocr_pdf(doc)
                if recognized.strip():
                    document_text = (extracted + '\n' + recognized if extracted else recognized)[:130_000]
                    ocr_status = 'OCR выполнен' + (f' с предупреждениями: {"; ".join(ocr_errors)[:150]}' if ocr_errors else '')
                elif ocr_errors:
                    ocr_status = 'Ошибка OCR: ' + '; '.join(ocr_errors)[:270]
                else:
                    ocr_status = 'OCR завершён, но читаемый текст не обнаружен'
            else:
                ocr_status = status['message']
        elif scanned:
            ocr_status = 'PDF является сканом: для чтения необходимо OCR'
        meta = doc.metadata or {}
    finally:
        doc.close()
    # A short but meaningful document can contain valid fields: don't mark it unread.
    fields = extract_properties(document_text)
    verified_text = len(document_text.strip()) >= 35
    fields.update({
        'pdf_read': verified_text,
        'pdf_scanned': scanned,
        'ocr_used': bool(ocr_status.startswith('OCR выполнен')),
        'ocr_status': ocr_status,
        'pdf_pages': page_count,
        'pdf_bytes': len(data),
        'pdf_title': tidy(meta.get('title', ''), 180),
        'pdf_filename': extract_pdf_filename(url),
        'text': document_text,
        'resolved_url': url,
        'pdf_verified': True,
    })
    return fields


def parse_pdf(url: str, material: str, *, enable_ocr: bool = True):
    data, content_type, resolved_url = safe_request(url)
    if not data[:1024].lstrip().startswith(b'%PDF-'):
        raise ValueError('По ссылке находится веб-страница или другой тип файла, а не PDF')
    return parse_pdf_bytes(data, material, resolved_url, enable_ocr=enable_ocr)


def expand_page(url: str, desired_type: str, material: str) -> list[dict]:
    """Pick document links from accessible HTML pages, not unrelated site navigation."""
    try:
        data, kind, landing = safe_request(url, max_bytes=1_400_000, timeout=7)
        if 'pdf' in kind or data[:5] == b'%PDF-':
            return [{'url': landing, 'title': 'PDF-документ', 'snippet': material,
                     'source_page': url, 'source_verified': True, 'priority': 30, 'force_pdf': True}]
        if 'html' not in kind and not data.lstrip()[:1] == b'<':
            return []
        soup = BeautifulSoup(data, 'html.parser')
        link_items = []
        count = 0
        for a in soup.select('a[href]'):
            raw_href = a.get('href', '')
            link = urljoin(landing, raw_href)
            if not safe_http_url(link) or not re.search(r'\.pdf(?:[?#]|$)', link, re.I):
                continue
            anchor = tidy(a.get_text(' ', strip=True), 180)
            parent = tidy(a.parent.get_text(' ', strip=True), 400) if a.parent else anchor
            if desired_type != 'all' and guess_type(f'{anchor} {parent} {link}') not in (desired_type, 'unknown'):
                continue
            count += 1
            score = sum(x.lower() in (anchor + ' ' + parent + ' ' + link).lower() for x in ('сертификат', 'паспорт', 'декларация', 'испытани', 'арматур'))
            score += 2 if re.search(r'500|34028', parent, re.I) and re.search(r'500|34028', material, re.I) else 0
            link_items.append({'url': link, 'title': anchor or 'PDF-документ', 'snippet': parent,
                               'source_page': landing, 'source_verified': True, 'priority': score})
            if count >= 70:
                break
        return sorted(link_items, key=lambda x: x['priority'], reverse=True)[:8]
    except Exception as exc:
        log.debug('Cannot expand %s: %s', url, exc)
        return []


def perform_search(material: str, dtype: str, manufacturer: str, gost: str, limit: int):
    queries = create_queries(material, dtype, manufacturer, gost)
    raw = []
    errors = []
    api_key = os.environ.get('BRAVE_SEARCH_API_KEY', '').strip()
    for query in queries:
        try:
            if api_key:
                r = requests.get('https://api.search.brave.com/res/v1/web/search',
                                 params={'q': query, 'country': 'RU', 'search_lang': 'ru', 'count': 12},
                                 headers={'Accept': 'application/json', 'X-Subscription-Token': api_key},
                                 timeout=12)
                r.raise_for_status()
                found = [{'href': x.get('url', ''), 'title': x.get('title', ''),
                          'body': x.get('description', '')}
                         for x in r.json().get('web', {}).get('results', [])]
            else:
                with DDGS(timeout=9) as search:
                    found = search.text(query, region='ru-ru', safesearch='moderate', max_results=12)
            for hit in (found or []):
                url = hit.get('href') or hit.get('url') or ''
                if not safe_http_url(url):
                    continue
                raw.append({'url': url,
                            'title': tidy(hit.get('title', ''), 240),
                            'snippet': tidy(hit.get('body', ''), 450),
                            'source_page': url, 'priority': 0})
        except Exception as exc:
            errors.append(f'Поиск «{query[:85]}»: {type(exc).__name__}: {str(exc)[:150]}')

    dedup = {}
    for obj in raw:
        key = obj['url'].split('#')[0].rstrip('/')
        if key not in dedup:
            dedup[key] = obj
    candidates = list(dedup.values())[:MAX_SEARCH_RESULTS]
    file_candidates = [v for v in candidates if re.search(r'\.pdf(?:[?#]|$)', v['url'], re.I)]
    landing_candidates = [v for v in candidates if v not in file_candidates]

    with ThreadPoolExecutor(max_workers=2) as executor:
        futures = [executor.submit(expand_page, v['url'], dtype, material)
                   for v in landing_candidates[:MAX_HTML_PAGES_TO_EXPAND]]
        for fut in as_completed(futures):
            try:
                file_candidates.extend(fut.result())
            except Exception as exc:
                log.debug('Expanding candidate failed: %s', exc)

    # Only PDFs proven accessible from this server go into public results.
    # A search hit is not a working document, even when its URL ends in ".pdf".
    combined = file_candidates + landing_candidates
    unique = {}
    for x in combined:
        key = x['url'].split('#')[0].rstrip('/')
        if key not in unique or x.get('priority', 0) > unique[key].get('priority', 0):
            unique[key] = x

    def score_candidate(c):
        info = (c['title'] + ' ' + c['snippet'] + ' ' + c['url']).lower()
        is_pdf_hint = bool(c.get('force_pdf') or re.search(r'\.pdf(?:[?#]|$)', c['url'], re.I))
        score = (25 if is_pdf_hint else 0) + c.get('priority', 0)
        if dtype != 'all' and guess_type(info) == dtype:
            score += 15
        if manufacturer and manufacturer.lower() in info:
            score += 8
        if gost and gost.lower() in info:
            score += 8
        class_code, dia = material_clues(material)
        if class_code and class_code in info.upper().replace('А', 'A').replace('С', 'C'):
            score += 6
        if dia and re.search(rf'(?<!\d){dia}\s*(?:мм|mm|[Øø])?', info, re.I):
            score += 3
        return score

    ranked = sorted(unique.values(), key=score_candidate, reverse=True)
    # Probe PDF-like links first; also accept a small number of direct download URLs.
    direct = [c for c in ranked if (c.get('force_pdf') or re.search(r'\.pdf(?:[?#]|$)', c['url'], re.I))]
    other = [c for c in ranked if c not in direct]
    to_verify = (direct + other[:5])[:MAX_PDFS_TO_READ]
    validated = {}
    discarded = 0
    with ThreadPoolExecutor(max_workers=4) as executor:
        future_map = {
            executor.submit(parse_pdf, c['url'], material, enable_ocr=False): c
            for i, c in enumerate(to_verify)
        }
        for f in as_completed(future_map):
            candidate = future_map[f]
            try:
                info = f.result()
                # PyMuPDF must have opened the real PDF, not just HTTP 200.
                if not info.get('pdf_verified') or not info.get('pdf_pages'):
                    discarded += 1
                    continue
                validated[candidate['url']] = info
            except Exception as exc:
                discarded += 1
                log.info('PDF discarded (%s): %s', urlsplit(candidate['url']).hostname, tidy(str(exc), 150))
    # After validating links, OCR the first three actually accessible scans.
    # No OCR time is spent on dead URLs, HTML responses or damaged PDFs.
    scan_candidates = [c for c in to_verify if c['url'] in validated and
                       validated[c['url']].get('pdf_scanned', False)][:3]
    if scan_candidates:
        with ThreadPoolExecutor(max_workers=2) as executor:
            ocr_futures = {executor.submit(parse_pdf, c['url'], material, enable_ocr=True): c['url']
                           for c in scan_candidates}
            for f in as_completed(ocr_futures):
                try:
                    recognized = f.result()
                    if recognized.get('pdf_verified'):
                        validated[ocr_futures[f]] = recognized
                except Exception as exc:
                    # The original PDF is still known to exist; retain it and
                    # allow re-running OCR via the inspect button.
                    validated[ocr_futures[f]]['ocr_status'] = 'Повторная загрузка для OCR не удалась: ' + tidy(str(exc), 150)
    if discarded:
        errors.append(f'Исключено недоступных или повреждённых PDF: {discarded}.')
    if not to_verify:
        errors.append('Поисковая система не вернула проверяемых PDF-ссылок.')

    result = []
    for hit in to_verify:
        url = hit['url']
        info = validated.get(url)
        if not info:
            continue
        detected_type = info.get('detected_type', 'unknown')
        if detected_type == 'unknown':
            detected_type = guess_type(hit['title'] + ' ' + hit['snippet'])
        if dtype != 'all' and detected_type not in (dtype, 'unknown'):
            continue
        status, explanation = assess_match(material, info.get('text', ''),
                                           hit['title'], info.get('pdf_read', False))
        supplement_conflict, supplement_notes = check_supplier_and_standard(manufacturer, gost, info)
        if supplement_conflict:
            status = 'Возможное расхождение'
        if supplement_notes:
            explanation = '; '.join(filter(None, [explanation, *supplement_notes]))
        number = info.get('number', '')
        kind = DOC_TYPES.get(detected_type, 'Вид документа не установлен')
        header = (f'{kind} № {number}' if number else
                  hit['title'] or info.get('pdf_filename') or 'PDF-документ')
        resolved = info.get('resolved_url', url)
        result.append({
            'title': tidy(header, 220),
            'requested_material': material,
            'material_in_document': info.get('material_in_document', ''),
            'gost_in_document': info.get('gost_in_document', ''),
            'document_type': kind,
            'document_type_verified': info.get('pdf_read', False) and info.get('detected_type') != 'unknown',
            'document_number': number,
            'document_date': info.get('date', ''),
            'valid_until': info.get('valid_until', ''),
            'pdf_pages': info.get('pdf_pages', ''),
            'pdf_bytes': info.get('pdf_bytes', ''),
            'pdf_filename': info.get('pdf_filename', '') or extract_pdf_filename(resolved),
            'pdf_title': info.get('pdf_title', ''),
            'ocr_used': info.get('ocr_used', False),
            'ocr_status': info.get('ocr_status', ''),
            'manufacturer': info.get('manufacturer', ''),
            'url': url,
            'redirect_url': resolved if resolved != url else '',
            'inspect_token': signer.dumps(url),
            'source_page': hit.get('source_page', '') if hit.get('source_verified') else '',
            'is_pdf': True,
            'pdf_verified': True,
            'pdf_read': info.get('pdf_read', False),
            'pdf_scanned': info.get('pdf_scanned', False),
            'read_error': '',
            'status': status,
            'explanation': explanation,
            'source_note': ('Распознанный текст скана (OCR)' if info.get('ocr_used') else 'Текст файла PDF') if info.get('pdf_read') else 'PDF открылся, текст не распознан',
            'snippet': hit['snippet'],
        })
        if len(result) >= limit:
            break
    return result, errors, len(raw)


def cached_search(material, dtype, manufacturer, gost, limit):
    key = (material, dtype, manufacturer, gost, limit)
    now = time.monotonic()
    with cache_lock:
        hit = search_cache.get(key)
        if hit and (now - hit[0]) < (60 if any(r.get('read_error') for r in hit[1][0]) else SEARCH_CACHE_TTL):
            search_cache.move_to_end(key)
            return hit[1]
    result = perform_search(material, dtype, manufacturer, gost, limit)
    with cache_lock:
        search_cache[key] = (time.monotonic(), result)
        search_cache.move_to_end(key)
        while len(search_cache) > SEARCH_CACHE_SIZE:
            search_cache.popitem(last=False)
    return result


@app.route('/')
def home():
    return render_template('index.html', types=DOC_TYPES)


@app.route('/api/search', methods=['POST'])
def search_api():
    data = request.get_json(silent=True) or {}
    material = normalize_material(str(data.get('material', '')))[:170]
    dtype = str(data.get('document_type', 'certificate'))
    manufacturer = tidy(data.get('manufacturer', ''), 120)
    gost = tidy(data.get('gost', ''), 120)
    try:
        limit = max(5, min(int(data.get('limit', 20)), 40))
    except (TypeError, ValueError):
        limit = 20
    if len(material) < 3:
        return jsonify({'error': 'Введите наименование материала (не менее 3 символов).'}), 400
    if dtype not in DOC_TYPES:
        return jsonify({'error': 'Некорректный тип документа.'}), 400
    try:
        result_data = run_limited_job(lambda: cached_search(material, dtype, manufacturer, gost, limit))
        if result_data is None:
            return jsonify({'error': 'Поиск занят. Повторите запрос через некоторое время.'}), 429
        results, errors, raw_count = result_data
    except Exception as exc:
        log.exception('Search failed')
        return jsonify({'error': f'Поисковый сервис недоступен: {type(exc).__name__}: {str(exc)[:180]}'}), 503
    return jsonify({'results': results, 'warnings': errors[:5], 'raw_count': raw_count,
                    'search_queries': create_queries(material, dtype, manufacturer, gost),
                    'notice': 'Совпадение признаков не означает подтверждения подлинности, срока действия или принадлежности конкретной партии.'})


@app.route('/api/file')
def download_checked_pdf():
    url = request.args.get('url', '')[:1800]
    token = request.args.get('token', '')[:4000]
    try:
        if not hmac.compare_digest(url, signer.loads(token, max_age=3600)):
            raise BadSignature('URL mismatch')
    except (BadSignature, SignatureExpired):
        return jsonify({'error': 'Ссылка на скачивание устарела. Повторите поиск.'}), 403
    if not safe_http_url(url):
        return jsonify({'error': 'Недопустимая ссылка.'}), 400
    try:
        content, _, resolved = safe_request(url, timeout=15)
        if not content[:1024].lstrip().startswith(b'%PDF-'):
            return jsonify({'error': 'Источник перестал возвращать PDF.'}), 422
        with fitz.open(stream=content, filetype='pdf') as doc:
            if len(doc) == 0:
                raise ValueError('Пустой PDF')
        filename = extract_pdf_filename(resolved)
        if not filename.lower().endswith('.pdf'):
            filename = 'dokument.pdf'
        # Don't use arbitrary upstream filename as a response header without sanitization.
        filename = re.sub(r'[^a-zA-Z0-9_.-]', '_', filename)[:120]
        return send_file(io.BytesIO(content), as_attachment=True, download_name=filename, mimetype='application/pdf')
    except Exception as exc:
        return jsonify({'error': 'Не удалось повторно получить файл: ' + tidy(str(exc), 170)}), 422


@app.route('/api/ocr-status')
def ocr_status_api():
    return jsonify(tesseract_status())


@app.route('/api/inspect', methods=['POST'])
def inspect_api():
    data = request.get_json(silent=True) or {}
    url = str(data.get('url', ''))[:1800]
    material = normalize_material(str(data.get('material', '')))[:170]
    token = str(data.get('inspect_token', ''))[:4000]
    try:
        signed_url = signer.loads(token, max_age=3600)
        if not hmac.compare_digest(url, signed_url):
            raise BadSignature('URL mismatch')
    except (BadSignature, SignatureExpired):
        return jsonify({'error': 'Ссылка на проверку устарела. Повторите поиск.'}), 403
    if not safe_http_url(url):
        return jsonify({'error': 'Недопустимая ссылка.'}), 400
    try:
        info = run_limited_job(lambda: parse_pdf(url, material))
        if info is None:
            return jsonify({'error': 'Обработка PDF занята. Повторите попытку.'}), 429
    except Exception as exc:
        return jsonify({'error': f'Не удалось прочитать PDF: {str(exc)[:230]}'}), 422
    status, explanation = assess_match(material, info['text'], '', info['pdf_read'])
    meta = response_metadata(info, status, explanation)
    meta['url'] = url  # Preserve signed search URL across any redirects.
    return jsonify({'metadata': meta})


def response_metadata(info: dict, status: str, explanation: str) -> dict:
    return {
        'material_in_document': info.get('material_in_document', ''),
        'gost_in_document': info.get('gost_in_document', ''),
        'document_type': DOC_TYPES.get(info.get('detected_type', ''), 'Не установлен'),
        'document_type_verified': info.get('pdf_read', False) and info.get('detected_type', 'unknown') != 'unknown',
        'document_number': info.get('number', ''),
        'document_date': info.get('date', ''),
        'valid_until': info.get('valid_until', ''),
        'manufacturer': info.get('manufacturer', ''),
        'pdf_filename': info.get('pdf_filename', ''),
        'pdf_title': info.get('pdf_title', ''),
        'pdf_pages': info.get('pdf_pages', ''),
        'pdf_bytes': info.get('pdf_bytes', ''),
        'pdf_read': info.get('pdf_read', False),
        'pdf_verified': info.get('pdf_verified', False),
        'pdf_scanned': info.get('pdf_scanned', False),
        'ocr_used': info.get('ocr_used', False),
        'ocr_status': info.get('ocr_status', ''),
        'read_error': '',
        'status': status,
        'explanation': explanation,
        'source_note': 'Распознанный текст PDF (OCR)' if info.get('ocr_used') else 'Текст PDF' if info.get('pdf_read') else 'PDF скачан, но текст пока не распознан',
        # URL is intentionally not replaced: signed source URL must remain stable.
    }


@app.route('/api/upload', methods=['POST'])
def upload_pdf_api():
    file = request.files.get('file')
    if not file or not file.filename:
        return jsonify({'error': 'Выберите PDF-файл.'}), 400
    material = normalize_material(request.form.get('material', ''))[:170]
    data = file.stream.read(MAX_PDF_BYTES + 1)
    if len(data) > MAX_PDF_BYTES:
        return jsonify({'error': f'PDF превышает лимит {MAX_PDF_BYTES // 1024 // 1024} МБ.'}), 413
    try:
        info = run_limited_job(lambda: parse_pdf_bytes(data, material))
        if info is None:
            return jsonify({'error': 'Обработка PDF занята. Повторите попытку.'}), 429
    except Exception as exc:
        return jsonify({'error': f'Не удалось прочитать документ: {str(exc)[:220]}'}), 422
    status, explanation = assess_match(material, info.get('text', ''), '', info['pdf_read'])
    meta = response_metadata(info, status, explanation)
    meta.update({'title': tidy(file.filename, 170), 'source_note': meta['source_note'] + ' · загружен пользователем',
                 'is_pdf': True, 'from_upload': True, 'requested_material': material,
                 'pdf_filename': tidy(file.filename, 170)})
    return jsonify({'metadata': meta})


def excel_safe(value) -> str:
    """Prevent formulas from executing when opening search results in Excel."""
    value = str(value or '')[:1500]
    return "'" + value if value.lstrip().startswith(('=', '+', '-', '@')) else value


@app.route('/api/export', methods=['POST'])
def export():
    data = request.get_json(silent=True) or {}
    rows = data.get('results', [])
    if not isinstance(rows, list) or len(rows) > 100:
        return jsonify({'error': 'Неверные данные для выгрузки'}), 400
    wb = Workbook()
    ws = wb.active
    ws.title = 'Результаты поиска'
    cols = [
        ('requested_material', 'Запрошенный материал'),
        ('material_in_document', 'Наименование материала в PDF'),
        ('gost_in_document', 'ГОСТ / ТУ в PDF'),
        ('document_type', 'Вид документа'),
        ('title', 'Наименование найденного документа'),
        ('document_number', 'Номер документа'),
        ('document_date', 'Дата документа'),
        ('valid_until', 'Действителен до'),
        ('pdf_filename', 'Имя PDF-файла'),
        ('pdf_pages', 'Количество страниц PDF'),
        ('ocr_status', 'Распознавание скана'),
        ('read_error', 'Ошибка чтения PDF'),
        ('manufacturer', 'Изготовитель'),
        ('status', 'Результат сопоставления'),
        ('explanation', 'Признаки сопоставления'),
        ('source_note', 'Источник реквизитов'),
        ('url', 'Ссылка на файл / страницу'),
        ('source_page', 'Страница-источник'),
    ]
    ws.append([name for _, name in cols])
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(2, len(rows) + 1)}"
    for row in rows:
        if not isinstance(row, dict):
            continue
        ws.append([excel_safe(row.get(key, '')) for key, _ in cols])
    for cell in ws[1]:
        cell.fill = PatternFill('solid', fgColor='173455')
        cell.font = Font(name='Calibri', color='FFFFFF', bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    for i, w in enumerate([29, 48, 25, 25, 56, 28, 18, 22, 35, 15, 28, 32, 32, 26, 64, 25, 75, 75], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical='top', wrap_text=True)
        ws.row_dimensions[row[0].row].height = 36
    ws.row_dimensions[1].height = 34
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name='poisk_dokumentov_materialov.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    # For maintenance only; public hosting uses gunicorn.
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '10000')), debug=False, threaded=True)
