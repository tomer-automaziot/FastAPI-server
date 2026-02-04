FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies including LibreOffice for XLS conversion
RUN apt-get update && apt-get install -y \
    libjpeg-dev \
    zlib1g-dev \
    libpng-dev \
    libreoffice-calc \
    libreoffice-common \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/* \
    && rm -rf /var/cache/apt/*

# Copy requirements first (for Docker cache)
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app.py .
COPY static/ ./static/

# Create temp directory for LibreOffice
RUN mkdir -p /tmp/libreoffice && chmod 777 /tmp/libreoffice

# Set LibreOffice home
ENV HOME=/tmp

# Expose port
EXPOSE 8080

# Run with gunicorn for production
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "2", "--timeout", "300", "app:app"]
