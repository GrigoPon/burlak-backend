import sqlite3
from pathlib import Path

import pytest
from sqlalchemy import create_engine, text

from app.core.config import get_settings
from app.db.database import Base

# Явно импортируем модели
from app.db.models import Cards, Jobs  # noqa: F401


@pytest.fixture(scope="function")
def temp_db_path(tmp_path: Path):
    """Создаёт тестовую БД, **гарантированно** создаёт все таблицы
    и переключает приложение (через settings.db_url) на эту БД,
    чтобы app.db.database.get_async_db / get_db реально использовали её.
    """
    db_file = tmp_path / "test.db"
    db_path = str(db_file)

    # Создаём подключение
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.commit()
    conn.close()

    # Создаём engine и таблицы
    engine = create_engine(f"sqlite:///{db_path}", echo=False)
    Base.metadata.create_all(engine)

    # Принудительная проверка и создание, если что-то пропущено
    with engine.connect() as con:
        con.execute(text("PRAGMA foreign_keys=ON"))
        tables = [
            row[0]
            for row in con.execute(
                text("SELECT name FROM sqlite_master WHERE type='table'")
            ).fetchall()
        ]

        if "jobs" not in tables:
            print("⚠️ Таблицы не найдены, создаём вручную...")
            Base.metadata.create_all(engine)  # повторная попытка
    engine.dispose()

    # КРИТИЧНО: settings — синглтон (get_settings закэширован), и именно
    # его читает app.db.database.get_async_db()/get_db() при каждом запросе.
    # Без этой подмены приложение продолжает стучаться в боевую/дефолтную БД,
    # где таблицы jobs не существует -> "no such table: jobs".
    settings = get_settings()
    original_db_url = settings.db_url
    settings.db_url = db_path

    yield db_path

    # Восстанавливаем настройки
    settings.db_url = original_db_url

    # Очистка файлов БД (основной + WAL/SHM)
    for suffix in ("", "-wal", "-shm"):
        p = Path(db_path + suffix)
        try:
            if p.exists():
                p.unlink()
        except Exception:
            pass


@pytest.fixture(scope="function")
def mock_storage_path(tmp_path: Path):
    settings = get_settings()
    original = settings.storage_path
    settings.storage_path = str(tmp_path)
    yield tmp_path
    settings.storage_path = original


@pytest.fixture
def api_client(temp_db_path: str):
    """TestClient явно зависит от temp_db_path: settings.db_url должен быть
    переключён на тестовую БД ДО того, как приложение начнёт обрабатывать запросы.
    """
    from fastapi.testclient import TestClient

    from app.main import app

    return TestClient(app)
