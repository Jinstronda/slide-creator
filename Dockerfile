FROM python:3.12-slim

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    libcairo2-dev \
    pkg-config \
    python3-dev \
    libxcb1-dev \
    libxcb-render0-dev \
    libxcb-shm0-dev \
    gcc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy and install dependencies first (better caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Set Python path
ENV PYTHONPATH=/app

# Expose port (documentation only)
EXPOSE 8000

# Start with shell form to properly expand PORT variable
CMD sh -c "uvicorn api.app:app --host 0.0.0.0 --port ${PORT:-8000}"
