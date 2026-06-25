# ============================================
# Dockerfile для Burlak Backend (FastAPI + Celery)
# Используем uv sync --no-dev
# ============================================

FROM python:3.12-slim

# Устанавливаем рабочую директорию
WORKDIR /app

# Устанавливаем системные зависимости
RUN apt-get update && apt-get install -y \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Устанавливаем uv
RUN pip install --no-cache-dir uv

# Копируем файлы проекта
COPY pyproject.toml uv.lock ./

# Устанавливаем зависимости в системный Python (без виртуального окружения)
RUN uv pip install --no-cache --system -r pyproject.toml

# Копируем весь код
COPY . .

# Создаем директории
RUN mkdir -p /app/data /app/storage /app/reports /app/output

# Переменные окружения
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    REDIS_URL=redis://redis:6379/0 \
    SQLITE_DB_PATH=/app/data/tasks.db \
    STORAGE_PATH=/app/storage

# Открываем порт
EXPOSE 8000

# Команда запуска
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
