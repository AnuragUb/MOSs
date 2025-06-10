FROM python:3.9-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    ffmpeg \
    vlc \
    libvlc-dev \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements first to leverage Docker cache
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create uploads directory
RUN mkdir -p /tmp/uploads

# Set environment variables
ENV PORT=8080
ENV PYTHONUNBUFFERED=1
ENV MEMORY_LIMIT=256Mi
ENV CPU_LIMIT=1
ENV MAX_CONCURRENCY=80
ENV TIMEOUT=300

# Run the application with optimized settings for free tier
CMD exec gunicorn \
    --bind :$PORT \
    --workers 1 \
    --threads 8 \
    --timeout $TIMEOUT \
    --worker-class gthread \
    --worker-tmp-dir /dev/shm \
    --log-level info \
    app:app 