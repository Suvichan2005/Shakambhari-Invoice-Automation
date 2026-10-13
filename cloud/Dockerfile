# Use official Python runtime
FROM python:3.11-slim

# Install system dependencies for PDF generation.
# We use LibreOffice headless to convert generated XLSX to PDF for visual parity.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \
    shared-mime-info \
    libreoffice \
    libreoffice-calc \
    fonts-crosextra-carlito \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements first for caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create directories for temporary files
RUN mkdir -p /tmp/invoices /tmp/pdfs

# Set environment variables
ENV PORT=8080
ENV PYTHONUNBUFFERED=1

# Run the application
CMD exec gunicorn --bind :$PORT --workers 2 --timeout 300 app_cloud:app
