FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    TZ=Europe/Berlin

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        curl \
        tzdata \
        libreoffice-calc \
        libreoffice-impress \
        fonts-dejavu \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

COPY watcher/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY watcher/watcher.py ./watcher.py

RUN mkdir -p /data/input /data/output /data/archive /data/error /data/mapping /data/work /data/logs /data/assets

CMD ["python", "/app/watcher.py"]
