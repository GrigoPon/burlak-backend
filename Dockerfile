# ============================================
# Dockerfile для Burlak Backend (FastAPI + Celery)
# ============================================

FROM python:3.12-slim

WORKDIR /app

RUN apt-get update && apt-get install -y \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Копируем uv и uvx из официального образа (БЫСТРЕЕ!)
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /usr/local/bin/

# Копируем зависимости
COPY pyproject.toml uv.lock ./

# СОЗДАЁМ ВИРТУАЛЬНОЕ ОКРУЖЕНИЕ
RUN uv venv /app/.venv

# АКТИВИРУЕМ И УСТАНАВЛИВАЕМ ЗАВИСИМОСТИ
RUN . /app/.venv/bin/activate && \
    uv sync --frozen --no-dev --no-install-project

# Копируем код
COPY . .

# Создаём папку для данных
RUN mkdir -p /data

# Переменные окружения
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    UV_COMPILE_BYTECODE=1 \
    PATH="/app/.venv/bin:$PATH" \
    PYTHONPATH="/app:$PYTHONPATH" \
    REDIS_URL=redis://redis:6379/0 \
    DB_URL=/data/jobs.db \
    STORAGE_PATH=/data

EXPOSE 8000

# Запускаем uvicorn через полный путь (гарантированно)
CMD ["/app/.venv/bin/uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
