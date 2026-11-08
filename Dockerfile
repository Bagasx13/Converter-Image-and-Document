# Gunakan image Python resmi
FROM python:3.11-slim

# Install LibreOffice & dependensi sistem lain yang dibutuhkan
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-impress \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements & install dependensi Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy seluruh kode project
COPY . .

# Railway memberikan variabel PORT, tapi kita bisa ekspos 8000 sebagai default
EXPOSE 8000

# Jalankan server
CMD ["python", "main.py"]
