# ──────────────────────────────────────────────
# Thomson EA Dashboard — Docker Image
# Base: Python 3.13 slim (Windows container)
# Note: MetaTrader5 library only works on Windows
#       so this must run on Windows Docker Desktop
#       or Windows VPS with Docker Engine (Windows containers)
# ──────────────────────────────────────────────
FROM python:3.13-windowsservercore-ltsc2022

WORKDIR /app

# Copy requirements first for layer caching
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy dashboard source
COPY app.py .

# Streamlit config
COPY .streamlit/config.toml .streamlit/config.toml

# Expose port
EXPOSE 8501

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD curl -f http://localhost:8501/_stcore/health || exit 1

# Run
ENTRYPOINT ["streamlit", "run", "app.py", \
            "--server.address=0.0.0.0", \
            "--server.port=8501", \
            "--server.headless=true"]
