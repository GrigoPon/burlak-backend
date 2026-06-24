import os
import sqlite3
import threading
from collections.abc import Generator

import aiosqlite
import pytest
from sqlalchemy import create_engine

from app.core.config import get_settings
from app.db import (
    async_repository,
    models,  # noqa: F401
    sync_repository,
)
from app.db.database import Base


@pytest.fixture
def test_db() -> Generator[str, None, None]:
    db_path = "./test_run.db"
    if os.path.exists(db_path):
        os.remove(db_path)

    # Override settings db_url for tests
    settings = get_settings()
    original_db_url = settings.db_url
    settings.db_url = db_path

    # Re-initialize engine and tables on the test database file
    test_engine = create_engine(f"sqlite:///{db_path}")
    Base.metadata.create_all(test_engine)

    # Set WAL mode
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.close()

    yield db_path

    # Restore original settings and cleanup
    settings.db_url = original_db_url
    if os.path.exists(db_path):
        os.remove(db_path)
    for suffix in ["-wal", "-shm"]:
        path = db_path + suffix
        if os.path.exists(path):
            os.remove(path)


@pytest.mark.asyncio
async def test_async_flow(test_db: str) -> None:
    async with aiosqlite.connect(test_db) as db:
        # Create Job
        job_id = await async_repository.create_job(db)
        assert job_id > 0

        # Verify job state
        job = await async_repository.get_job(db, job_id)
        assert job is not None
        assert job["status"] == "awaiting_upload"
        assert job["total"] == 0
        assert job["processed"] == 0
        assert job["failed"] == 0

        # Update file upload status
        await async_repository.update_file_upload(
            db, job_id, "bom", "/data/bom.xlsx", True
        )
        await async_repository.update_file_upload(
            db, job_id, "archive", "/data/archive.zip", True
        )

        job = await async_repository.get_job(db, job_id)
        assert job is not None
        assert job["bom_uploaded"] is True
        assert job["archive_uploaded"] is True
        assert job["bom_path"] == "/data/bom.xlsx"
        assert job["archive_path"] == "/data/archive.zip"

        # Update mapping config
        mapping = {"keys": ["Part Number", "Description"], "mappings": {}}
        await async_repository.update_mapping_config(db, job_id, mapping)

        job = await async_repository.get_job(db, job_id)
        assert job is not None
        assert job["mapping_config"] == mapping

        # Create cards
        card_paths = [f"card_{i}.xlsx" for i in range(10)]
        await async_repository.create_cards(db, job_id, card_paths)

        job = await async_repository.get_job(db, job_id)
        assert job is not None
        assert job["total"] == 10


def test_sync_concurrency(test_db: str) -> None:
    # 1. Create a job first via a temporary connection
    conn = sqlite3.connect(test_db)
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        ("processing", 10, 0, 0, 1, 1, "2026-06-22", "2026-06-22"),
    )
    job_id = cursor.lastrowid
    assert job_id is not None

    # Create card placeholders
    for i in range(10):
        cursor.execute(
            """
            INSERT INTO cards (job_id, card_path, status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (job_id, f"card_{i}.xlsx", "pending", "2026-06-22", "2026-06-22"),
        )
    conn.commit()
    conn.close()

    updates = [
        ("card_0.xlsx", True, None),
        ("card_1.xlsx", True, None),
        ("card_2.xlsx", True, None),
        ("card_3.xlsx", False, "Missing column"),
        ("card_4.xlsx", True, None),
        ("card_5.xlsx", False, "Corrupted XLSX formatting"),
        ("card_6.xlsx", True, None),
        ("card_7.xlsx", True, None),
        ("card_8.xlsx", False, "ML Translation timeout"),
        ("card_9.xlsx", True, None),
    ]

    def thread_worker(index: int) -> None:
        card_path, success, err_msg = updates[index]
        sync_repository.increment_progress(
            job_id, card_path, success=success, error_message=err_msg
        )

    threads = []
    for i in range(len(updates)):
        t = threading.Thread(target=thread_worker, args=(i,))
        threads.append(t)
        t.start()

    for t in threads:
        t.join()

    # Verify results
    conn = sqlite3.connect(test_db)
    conn.row_factory = sqlite3.Row
    job = conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()
    assert job["processed"] == 7
    assert job["failed"] == 3
    assert job["total"] == 10

    # Verify cards table
    failed_cards = conn.execute(
        "SELECT card_path, error_message FROM cards WHERE job_id = ? AND status = 'failed'",
        (job_id,),
    ).fetchall()
    assert len(failed_cards) == 3
    failed_paths = {r["card_path"] for r in failed_cards}
    assert failed_paths == {"card_3.xlsx", "card_5.xlsx", "card_8.xlsx"}

    # Test idempotency (updating card_0.xlsx again as success should not change counters)
    res = sync_repository.increment_progress(job_id, "card_0.xlsx", success=True)
    assert res.processed == 7
    assert res.failed == 3

    # Test updating card_0.xlsx to failed (should decrement processed, increment failed)
    res = sync_repository.increment_progress(
        job_id, "card_0.xlsx", success=False, error_message="Switched to failed"
    )
    assert res.processed == 6
    assert res.failed == 4

    # Test updating it back to success (should restore counters)
    res = sync_repository.increment_progress(job_id, "card_0.xlsx", success=True)
    assert res.processed == 7
    assert res.failed == 3
    conn.close()
