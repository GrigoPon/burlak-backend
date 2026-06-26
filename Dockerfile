# ============================================
# Dockerfile для Burlak Backend (FastAPI + Celery)
# ============================================

FROM python:3.12-slim

WORKDIR /app

RUN apt-get update && apt-get install -y curl && rm -rf /var/lib/apt/lists/*

# Копируем uv из официального образа
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /usr/local/bin/

# Копируем зависимости
COPY pyproject.toml uv.lock ./

# 1. СОЗДАЁМ ВИРТУАЛЬНОЕ ОКРУЖЕНИЕ
RUN uv venv /app/.venv --clear

# 2. УСТАНАВЛИВАЕМ ЗАВИСИМОСТИ ЧЕРЕЗ uv sync (БЕЗ PIP!)
RUN . /app/.venv/bin/activate && uv sync --frozen --no-dev

# Копируем код
COPY . .

# Создаём папку для данных
RUN mkdir -p /data

# Переменные окружения
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    REDIS_URL=redis://redis:6379/0 \
    DB_URL=/data/jobs.db \
    STORAGE_PATH=/data

EXPOSE 8000

# Запускаем через абсолютный путь к Python из .venv
CMD ["/app/.venv/bin/python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
