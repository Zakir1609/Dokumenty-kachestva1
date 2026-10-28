"""Regression for the Render build failure: parameters must be RapidOCR enums.

This is a **test double** for the RapidOCR API. It does not load real models.
"""
from enum import Enum
from pathlib import Path
import sys
import types

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))

class EngineType(Enum):
    ONNXRUNTIME = 'onnxruntime'

class ModelType(Enum):
    MOBILE = 'mobile'

class OCRVersion(Enum):
    PPOCRV4 = 'PP-OCRv4'
    PPOCRV5 = 'PP-OCRv5'

class LangRec(Enum):
    CYRILLIC = 'cyrillic'

class RapidOCR:
    def __init__(self, params):
        required={
           'Det.engine_type':EngineType,
           'Det.model_type':ModelType,
           'Det.ocr_version':OCRVersion,
           'Rec.engine_type':EngineType,
           'Rec.model_type':ModelType,
           'Rec.ocr_version':OCRVersion,
           'Rec.lang_type':LangRec,
        }
        for name,enum_type in required.items():
            value=params[name]
            assert isinstance(value,enum_type),f'{name} should be {enum_type}, got {value!r}'
        assert params['Det.engine_type'] is EngineType.ONNXRUNTIME
        assert params['Rec.engine_type'] is EngineType.ONNXRUNTIME
        assert params['Det.ocr_version'] is OCRVersion.PPOCRV4
        assert params['Rec.ocr_version'] is OCRVersion.PPOCRV5
        assert params['Rec.lang_type'] is LangRec.CYRILLIC
        assert params['Global.use_cls'] is False
        self.params=params

rapidocr=types.ModuleType('rapidocr')
for name in ('EngineType','ModelType','OCRVersion','LangRec','RapidOCR'):
    setattr(rapidocr,name,globals()[name])
sys.modules['rapidocr']=rapidocr
import local_ocr
engine=local_ocr._make_engine()
assert isinstance(engine,RapidOCR)
print('PASS: RapidOCR initializer passes Enums instead of strings (offline API double)')
