# Sử dụng Python 3.11 slim tối ưu kích thước container
FROM python:3.11-slim

# Ngăn Python ghi file .pyc và bật log realtime
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080

# Cài đặt các thư viện hệ thống cần thiết (xử lý tài liệu và mạng)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Cài đặt Python dependencies (tận dụng cache Docker layer)
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Sao chép toàn bộ mã nguồn vào container
COPY . .

# Đảm bảo các thư mục cần thiết tồn tại và có quyền ghi
RUN mkdir -p exports .streamlit

# Mở cổng 8080
EXPOSE 8080

# Chạy ứng dụng Streamlit lắng nghe biến môi trường $PORT từ Google Cloud Run
CMD ["sh", "-c", "streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0 --server.enableCORS=false --server.enableXsrfProtection=false --server.headless=true"]
