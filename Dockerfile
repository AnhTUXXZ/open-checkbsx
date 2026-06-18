FROM python:3.10-slim

# Cài đặt các gói thư viện hệ thống cần thiết cho Tesseract OCR và OpenCV trên Linux
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-vie \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Cài đặt các thư viện Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ code vào Docker
COPY . .

# Khai báo biến môi trường nhận diện Render
ENV RENDER=true

# Khởi chạy Flask thông qua Gunicorn với Port mặc định của Render
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-10000} app:app"]