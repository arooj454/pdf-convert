# Use Python 3.11 slim image
FROM python:3.11-slim

WORKDIR /app

ENV HOME=/tmp
ENV PYTHONUNBUFFERED=1
ENV DEBIAN_FRONTEND=noninteractive

# Install pandoc and wkhtmltopdf (lighter than LibreOffice)
RUN apt-get update && apt-get install -y \
    pandoc \
    wkhtmltopdf \
    fonts-liberation \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

RUN mkdir -p /tmp && chmod -R 777 /tmp

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install additional requirement for docx handling
RUN pip install --no-cache-dir pypandoc

COPY app.py .

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
