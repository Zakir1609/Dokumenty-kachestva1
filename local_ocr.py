"""Free local Cyrillic PDF OCR (RapidOCR + PP-OCRv5). No remote AI service.

OCR is only applied when selectable PDF text does not carry the desired fields.
The model is explicit *Cyrillic*; RapidOCR's default Chinese/English model is
unsuitable for Russian certificates. Models are warmed during Docker build.
"""
from __future__ import annotations

import importlib.util
import re
from threading import Lock

import fitz

from extractor import inspect_bytes, doc_type, document_number, document_date, material_name, manufacturer

_engine = None
_init_error = ''
_lock = Lock()

class RecognitionError(RuntimeError):
    pass


def recognition_status() -> dict:
    installed = bool(importlib.util.find_spec('rapidocr') and importlib.util.find_spec('onnxruntime'))
    return {
        'configured': installed,
        'ready': _engine is not None,
        'provider': 'RapidOCR + PP-OCRv5 Cyrillic (локально)',
        'reason': ('Модель загружена и готова' if _engine is not None else
                   'Модель будет инициализирована при первом скане' if installed else
                   'Пакеты rapidocr и onnxruntime отсутствуют'),
    }


def _make_engine():
    try:
        from rapidocr import RapidOCR
        return RapidOCR(params={
            'Det.engine_type': 'onnxruntime',
            'Det.model_type': 'mobile',
            'Det.ocr_version': 'PP-OCRv4',
            'Rec.engine_type': 'onnxruntime',
            'Rec.lang_type': 'cyrillic',
            'Rec.model_type': 'mobile',
            'Rec.ocr_version': 'PP-OCRv5',
            'Global.use_cls': False,
            'Global.use_preprocess_img': False,
        })
    except Exception as exc:
        raise RecognitionError('Не удалось загрузить модель RapidOCR для русского языка: ' +
                               type(exc).__name__ + ': ' + str(exc)[:160]) from exc


def _engine_instance():
    global _engine
    if _engine is None:
        _engine = _make_engine()
    return _engine


def _ocr_lines(page) -> list[str]:
    from PIL import Image
    # A4 at 137 dpi, memory-limited for Render Free (512 MB)
    factor = 1.9
    if max(page.rect.width, page.rect.height) * factor > 2600:
        factor = 2600 / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=fitz.Matrix(factor,factor), alpha=False, colorspace=fitz.csRGB)
    img = Image.frombytes('RGB', (pix.width,pix.height), pix.samples)
    del pix
    with _lock:  # onnxruntime inference is serialised across requests
        engine = _engine_instance()
        try:
            result = engine(img)
            lines = getattr(result, 'txts', None)
            return [str(s).strip() for s in (lines if lines is not None else []) if str(s).strip()]
        except Exception as exc:
            raise RecognitionError('Ошибка распознавания RapidOCR: '+type(exc).__name__+': '+str(exc)[:150]) from exc


def text_from_scans(data:bytes, max_pages:int=3) -> str:
    try:
        with fitz.open(stream=data,filetype='pdf') as pdf:
            if pdf.is_encrypted and not pdf.authenticate(''):
                raise RecognitionError('PDF защищён паролем')
            # First pages usually include material, registration and manufacturer.
            indexes=list(range(min(len(pdf),max_pages)))
            lines=[]
            for i in indexes:
                lines.extend(_ocr_lines(pdf[i]))
            return '\n'.join(lines)
    except RecognitionError:raise
    except Exception as exc:
        raise RecognitionError('Не удалось открыть PDF для OCR: '+type(exc).__name__) from exc


def _parsed(text: str, hint: str, filename: str) -> dict:
    kind=doc_type(text,filename)
    return {
        'material': material_name(text,hint),
        'number': document_number(text,kind),
        'date': document_date(text,kind) or '—',
        'manufacturer': manufacturer(text) or '—',
        'type':kind,
    }


def inspect_document(data:bytes, hint:str='', filename:str='', *, use_ocr:bool=True)->dict:
    """API-compatible with old site; use_ocr=True means local OCR if needed."""
    local=inspect_bytes(data,hint,filename,use_ocr=False)
    local['extraction_source']='pdf_text'
    # A text-only PDF with document and key details avoids expensive recognition.
    if not use_ocr:
        return local
    sufficient=(local['has_text'] and local['type'] and local['material'] and
                local['date']!='—' and local['manufacturer']!='—')
    if sufficient:
        return local
    # PDF with proper selectable text but with absent fields should not be
    # forced through OCR: missing manufacturer/date can be genuinely absent.
    if local['has_text'] and len(local['text'])>280 and local['type']:
        return local
    try:
        recognised=text_from_scans(data)
    except RecognitionError:
        if local['has_text']:
            local['ocr_note']='OCR недоступен; показан текстовый слой PDF'
            return local
        raise
    if not recognised or len(re.findall(r'[А-Яа-яA-Za-z]',recognised))<12:
        if local['has_text']:
            return local
        local.update(ocr_used=True,ocr_note='Распознавание не обнаружило читаемого текста',has_text=False)
        return local
    # Do not copy manufacturer/material from the search query or URL.
    combined=(local['text']+'\n'+recognised) if local['has_text'] else recognised
    parsed=_parsed(combined,hint,filename)
    local.update(parsed)
    local.update(text=combined[:140000],has_text=True,ocr_used=True,
                 ocr_note='Распознано локальной кириллической моделью RapidOCR',
                 extraction_source='rapidocr')
    return local
