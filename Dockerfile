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
RUN useradd --create-home --uid 10001 appuser && chown -R appuser:appuser /app
USER appuser
# Fetch Cyrillic model into the container DURING BUILD, never at runtime.
# Failure here is explicit deployment failure instead of empty OCR fields.
RUN python warm_ocr.py
EXPOSE 10000
CMD ["sh", "-c", "exec gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 2 --timeout 270 --access-logfile - --error-logfile - app:app"]
