# ============================================
# Dockerfile для Burlak Backend (FastAPI + Celery)
# ============================================

FROM python:3.12-slim

# Устанавливаем рабочую директорию
WORKDIR /app

# Устанавливаем системные зависимости
RUN apt-get update && apt-get install -y \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Копируем зависимости
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копируем весь код
COPY . .

# Создаем необходимые директории
RUN mkdir -p /app/data /app/storage /app/reports /app/output

# Переменные окружения
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    REDIS_URL=redis://redis:6379/0 \
    SQLITE_DB_PATH=/app/data/tasks.db \
    STORAGE_PATH=/app/storage

# Открываем порт для API
EXPOSE 8000

# Команда запуска (переопределяется для Celery)
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
