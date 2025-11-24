FROM python:3.11-slim

# Install ALL packages LibreOffice needs
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    default-jre \
    fonts-dejavu \
    fonts-liberation \
    libxinerama1 \
    libx11-xcb1 \
    libxrender1 \
    libxrandr2 \
    libglu1-mesa \
    mesa-utils \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
