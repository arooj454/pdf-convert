# Use Python 3.11 slim image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Set environment variables
ENV HOME=/tmp
ENV PYTHONUNBUFFERED=1
ENV DEBIAN_FRONTEND=noninteractive

# Install system dependencies including LibreOffice and all required libraries
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    libreoffice-java-common \
    default-jre-headless \
    fonts-liberation \
    fonts-dejavu-core \
    fonts-liberation2 \
    fonts-noto-core \
    fontconfig \
    libfontconfig1 \
    libxrender1 \
    libxext6 \
    libx11-6 \
    libxinerama1 \
    libgl1-mesa-glx \
    libcups2 \
    libdbus-glib-1-2 \
    ca-certificates \
    && fc-cache -f -v \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Create necessary directories with proper permissions
RUN mkdir -p /tmp/.config /tmp/libreoffice /tmp/cache && \
    chmod -R 777 /tmp

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app.py .

# Test LibreOffice installation
RUN soffice --headless --version || echo "LibreOffice version check failed"

# Expose port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8000/health')"

# Run the application with single worker
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
