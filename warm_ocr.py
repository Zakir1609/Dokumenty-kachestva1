"""Warm Cyrillic model during Docker build, so PDF scans work after startup."""
from PIL import Image, ImageDraw, ImageFont
from local_ocr import _engine_instance
font='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
img=Image.new('RGB',(1000,350),'white')
draw=ImageDraw.Draw(img)
ft=ImageFont.truetype(font,40)
draw.text((30,35),'СЕРТИФИКАТ СООТВЕТСТВИЯ',font=ft,fill='black')
draw.text((30,120),'Изготовитель: ООО МеталлПром',font=ft,fill='black')
result=_engine_instance()(img)
lines=getattr(result,'txts',None)
if not lines:
    raise SystemExit('Не удалось распознать кириллический проверочный текст; проверьте модель и загрузку')
print('RapidOCR Cyrillic smoke test:',str(lines)[:250])
