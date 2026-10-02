# ==============================================================================
# FileShift: Local-First File Converter
# Container image using Python 3.12-slim with Poppler & Ghostscript
# ==============================================================================

FROM python:3.12-slim

# Install system dependencies for document and PDF layout rendering
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    ghostscript \
    && rm -rf /var/lib/apt/lists/*

# Environment settings
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8080

# Create non-root system user for security
RUN useradd -m -u 1000 -s /bin/bash appuser

WORKDIR /app

# Copy application directories and server script
COPY converter/ ./converter/
COPY public/ ./public/
COPY server.py ./server.py
COPY README.md ./README.md

# Set ownership
RUN chown -R appuser:appuser /app

USER appuser

EXPOSE 8080

# Container healthcheck
HEALTHCHECK --interval=30s --timeout=5s --start-period=5s --retries=3 \
  CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/api/health')" || exit 1

ENTRYPOINT ["python3", "server.py", "8080"]
