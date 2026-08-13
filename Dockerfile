# ── Yori Prime Filter Bot ─────────────────────────────────────────────
# Runs in POLLING mode (no webhook). A tiny /health HTTP server listens
# on $PORT (Render injects this automatically) so Render's health check
# and UptimeRobot can keep the free service awake.
#
#   Build locally:   docker build -t yori-prime .
#   Run locally:     docker run --rm -e TELEGRAM_BOT_TOKEN=... -p 8080:8080 yori-prime
#   Deploy:          Render Blueprint (render.yaml) or "New Web Service" → Docker

FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Install dependencies first so the layer is cached across builds
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application
COPY . .

# Fallback port for local runs — Render overrides this with its own $PORT
ENV PORT=8080

EXPOSE 8080

# Lightweight container health check (optional; Render/UptimeRobot also ping /health)
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import os,urllib.request;urllib.request.urlopen('http://127.0.0.1:'+os.environ.get('PORT','8080')+'/health',timeout=3)" || exit 1

# Polling mode is the default (WEBHOOK_URL is left unset on Render)
CMD ["python", "main.py"]
