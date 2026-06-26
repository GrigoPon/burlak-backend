import sqlite3
from unittest.mock import patch

from fastapi.testclient import TestClient
from sqlalchemy import create_engine

from app.db.database import Base

# Примечание: api_client теперь берётся ИСКЛЮЧИТЕЛЬНО из conftest.py.
# Там он явно зависит от temp_db_path, что гарантирует: settings.db_url
# уже переключён на тестовую БД к моменту, когда TestClient начинает
# обрабатывать запросы. Локальная переопределённая фикстура без этой
# зависимости (как было раньше) — и есть причина "no such table: jobs".


def test_full_job_lifecycle(
    api_client: TestClient, temp_db_path: str, mock_storage_path
):
    """E2E тест создания задачи с явным созданием таблиц."""

    # === ЯВНОЕ СОЗДАНИЕ ТАБЛИЦ В ТОЙ ЖЕ БД, КОТОРУЮ ИСПОЛЬЗУЕТ ТЕСТ ===
    engine = create_engine(f"sqlite:///{temp_db_path}")
    Base.metadata.create_all(engine)

    # Проверка, что таблицы реально созданы
    conn = sqlite3.connect(temp_db_path)
    tables = [
        row[0]
        for row in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()
    ]
    conn.close()

    assert "jobs" in tables, f"Таблица 'jobs' не найдена! Найдены: {tables}"

    # ==================== ТЕСТ ====================

    # 1. Создание задачи
    resp = api_client.post("/api/v1/jobs")
    assert resp.status_code == 201, f"Ожидался 201, получен {resp.status_code}"
    job_id = resp.json()["id"]

    # 2. Загрузка BOM
    api_client.put(
        f"/api/v1/jobs/{job_id}/files/bom/chunks/0",
        content=b"bom content",
        headers={"X-Total-Chunks": "1"},
    )
    resp = api_client.post(f"/api/v1/jobs/{job_id}/files/bom/complete")
    assert resp.status_code == 200

    # 3. Загрузка Archive
    api_client.put(
        f"/api/v1/jobs/{job_id}/files/archive/chunks/0",
        content=b"archive content",
        headers={"X-Total-Chunks": "1"},
    )
    resp = api_client.post(f"/api/v1/jobs/{job_id}/files/archive/complete")
    assert resp.status_code == 200

    # 4. Запуск обработки
    with patch(
        "app.services.job_processing_service.JobProcessingService.dispatch_processing"
    ):
        resp = api_client.post(f"/api/v1/jobs/{job_id}/start")
        assert resp.status_code == 202

    print(f"✅ Тест пройден успешно! job_id = {job_id}")
