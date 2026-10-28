FROM python:3.11-slim-bookworm
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PORT=10000 \
    OMP_NUM_THREADS=1 OPENBLAS_NUM_THREADS=1 MKL_NUM_THREADS=1 \
    ORT_NUM_THREADS=1
RUN apt-get update && apt-get install -y --no-install-recommends ca-certificates libgomp1 libglib2.0-0 fonts-dejavu-core && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py extractor.py local_ocr.py warm_ocr.py ./
COPY templates/ ./templates/
# The RapidOCR package downloads mobile ONNX models into its own models/
# directory. That directory is protected when the image runs as appuser.
# Warm models during IMAGE BUILD as root, so no runtime writes/downloads.
RUN python warm_ocr.py
# The runtime itself stays unprivileged; models are read-only to appuser.
RUN useradd --create-home --uid 10001 appuser \
    && chown -R appuser:appuser /app \
    && chmod -R a+rX /usr/local/lib/python3.11/site-packages/rapidocr/models
USER appuser
# Verify that both detection and Cyrillic recognition models can be read
# by the non-root user; this check does not load the inference engine again.
RUN python -c "from pathlib import Path; import os; p=Path('/usr/local/lib/python3.11/site-packages/rapidocr/models'); f=list(p.rglob('*.onnx')); assert len(f)>=2 and any('cyrillic' in x.name.lower() for x in f), 'RapidOCR models missing'; assert all(os.access(x, os.R_OK) for x in f), 'RapidOCR model unreadable'; print('OCR models included and readable:', len(f))"
EXPOSE 10000
CMD ["sh", "-c", "exec gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 2 --timeout 270 --access-logfile - --error-logfile - app:app"]
