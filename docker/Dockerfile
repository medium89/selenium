FROM python:3.11-slim

# Install Chromium + Chromedriver and dependencies
RUN apt-get update && \
    DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
      chromium \
      chromium-driver \
      fonts-liberation \
      ca-certificates \
      dumb-init \
    && rm -rf /var/lib/apt/lists/*

ENV CHROME_BIN=/usr/bin/chromium \
    CHROMEDRIVER=/usr/bin/chromedriver \
    HEADLESS=1 \
    USER_DATA_DIR=/profile \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

RUN mkdir -p /profile && chmod 777 /profile

COPY config/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copy project files
COPY . /app

# Default to office_manager; override CMD to run project_manager if needed
CMD ["dumb-init", "python", "app/OfficeManager.py"]
