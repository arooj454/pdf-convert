# Use official Python slim image as base
FROM python:3.9-slim

# Install system dependencies for Tesseract OCR
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    libtesseract-dev \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements if you have any (Flask, FastAPI, etc.)
COPY requirements.txt .

# Install Python dependencies, including pytesseract
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY . .

# Expose port
EXPOSE 8000

# Start FastAPI or Flask app
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
