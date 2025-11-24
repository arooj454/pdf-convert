# Use official Python slim image as base
FROM python:3.9-slim

# Install LibreOffice and other dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-common \
    libreoffice-writer \
    libreoffice-calc \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements if you have any (Flask etc.)
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY . .

# Expose port for Flask app
EXPOSE 8000

# Start the Flask app
CMD ["python", "app.py"]
