import sqlite3
from collections.abc import AsyncGenerator, Generator
from pathlib import Path
from typing import Any
from unittest.mock import patch

import aiosqlite
import pytest
from fastapi.testclient import TestClient

from app.core.config import get_settings
from app.db.database import Base, get_async_db
from app.main import app


class MockRedisSuccess:
    async def ping(self) -> bool:
        return True

    async def __aenter__(self) -> "MockRedisSuccess":
        return self

    async def __aexit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        pass


class MockRedisFailure:
    async def ping(self) -> bool:
        raise Exception("Redis connection timed out")

    async def __aenter__(self) -> "MockRedisFailure":
        return self

    async def __aexit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        pass


@pytest.fixture
def temp_db_path(tmp_path: Path) -> Generator[str, None, None]:
    """Provides a temporary SQLite database path with initial schema."""
    db_file = tmp_path / "test_api.db"
    conn = sqlite3.connect(str(db_file))
    conn.execute("PRAGMA journal_mode=WAL")

    # Build schema
    from sqlalchemy import create_engine

    engine = create_engine(f"sqlite:///{db_file}")
    Base.metadata.create_all(engine)
    conn.close()

    yield str(db_file)

    if db_file.exists():
        db_file.unlink()


@pytest.fixture
def mock_storage_path(tmp_path: Path) -> Generator[Path, None, None]:
    """Fixture to override settings storage_path to a temp directory."""
    settings = get_settings()
    original_path = settings.storage_path
    settings.storage_path = str(tmp_path)
    yield tmp_path
    settings.storage_path = original_path


@pytest.fixture
def api_client(temp_db_path: str) -> Generator[TestClient, None, None]:
    """Provides a TestClient with overridden get_async_db dependency."""

    async def override_get_async_db() -> AsyncGenerator[aiosqlite.Connection, None]:
        async with aiosqlite.connect(temp_db_path) as db:
            db.row_factory = aiosqlite.Row
            yield db

    app.dependency_overrides[get_async_db] = override_get_async_db
    yield TestClient(app)
    app.dependency_overrides.clear()


# ----------------------------------------------------------------------
# Health Endpoint Tests
# ----------------------------------------------------------------------


def test_health_endpoint_healthy(
    api_client: TestClient,
    mock_storage_path: Path,
) -> None:
    """Test health endpoint returns 200 OK when all systems are healthy."""
    with patch("app.api.v1.health.aioredis.from_url", return_value=MockRedisSuccess()):
        response = api_client.get("/api/v1/health")
        assert response.status_code == 200
        data = response.json()
        assert data["status"] == "healthy"
        assert data["checks"]["database"] == "ok"
        assert data["checks"]["redis"] == "ok"
        assert data["checks"]["storage"] == "ok"


def test_health_endpoint_unhealthy(
    api_client: TestClient,
    mock_storage_path: Path,
) -> None:
    """Test health endpoint returns 503 Service Unavailable when a check fails."""
    with patch("app.api.v1.health.aioredis.from_url", return_value=MockRedisFailure()):
        response = api_client.get("/api/v1/health")
        assert response.status_code == 503
        data = response.json()
        assert data["status"] == "unhealthy"
        assert data["checks"]["redis"].startswith("failed:")


# ----------------------------------------------------------------------
# Jobs Endpoints Tests
# ----------------------------------------------------------------------


def test_create_job(api_client: TestClient) -> None:
    """Test creating a new job."""
    response = api_client.post("/api/v1/jobs")
    assert response.status_code == 201
    data = response.json()
    assert "id" in data
    assert data["status"] == "awaiting_upload"
    assert "created_at" in data


def test_get_job_status_success(api_client: TestClient, temp_db_path: str) -> None:
    """Test fetching job status of an existing job."""
    # Seed a job directly
    conn = sqlite3.connect(temp_db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            "awaiting_upload",
            0,
            0,
            0,
            0,
            0,
            "2026-06-22T00:00:00",
            "2026-06-22T00:00:00",
        ),
    )
    job_id = cursor.lastrowid
    conn.commit()
    conn.close()

    assert job_id is not None
    response = api_client.get(f"/api/v1/jobs/{job_id}")
    assert response.status_code == 200
    data = response.json()
    assert data["id"] == job_id
    assert data["status"] == "awaiting_upload"
    assert data["bom_uploaded"] is False


def test_get_job_status_not_found(api_client: TestClient) -> None:
    """Test fetching job status of a non-existent job."""
    response = api_client.get("/api/v1/jobs/99999")
    assert response.status_code == 404
    data = response.json()
    assert data["error"]["code"] == "JOB_NOT_FOUND"
    assert "not found" in data["error"]["message"]


def test_start_job_processing_success(
    api_client: TestClient, temp_db_path: str
) -> None:
    """Test starting job processing with valid state."""
    # Seed a job with uploaded files
    conn = sqlite3.connect(temp_db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, bom_path, archive_path, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            "awaiting_upload",
            0,
            0,
            0,
            1,
            1,
            "/data/bom.xlsx",
            "/data/archive.zip",
            "2026-06-22T00:00:00",
            "2026-06-22T00:00:00",
        ),
    )
    job_id = cursor.lastrowid
    conn.commit()
    conn.close()

    assert job_id is not None
    with patch(
        "app.services.job_processing_service.JobProcessingService.dispatch_processing"
    ) as mock_dispatch:
        response = api_client.post(f"/api/v1/jobs/{job_id}/start")
        assert response.status_code == 202
        data = response.json()
        assert data["job_id"] == job_id
        assert data["status"] == "processing"
        assert data["stage"] == "unpacking"
        mock_dispatch.assert_called_once_with(job_id)


def test_start_job_processing_validation_error(
    api_client: TestClient, temp_db_path: str
) -> None:
    """Test starting job processing with missing files triggers 409 error."""
    # Seed a job without uploaded files
    conn = sqlite3.connect(temp_db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            "awaiting_upload",
            0,
            0,
            0,
            0,
            0,
            "2026-06-22T00:00:00",
            "2026-06-22T00:00:00",
        ),
    )
    job_id = cursor.lastrowid
    conn.commit()
    conn.close()

    assert job_id is not None
    response = api_client.post(f"/api/v1/jobs/{job_id}/start")
    assert response.status_code == 409
    data = response.json()
    assert data["error"]["code"] == "INVALID_JOB_STATE"


# ----------------------------------------------------------------------
# Files Endpoints Tests
# ----------------------------------------------------------------------


def test_upload_chunk(
    api_client: TestClient,
    temp_db_path: str,
    mock_storage_path: Path,
) -> None:
    """Test uploading a file chunk."""
    # Seed a job
    conn = sqlite3.connect(temp_db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            "awaiting_upload",
            0,
            0,
            0,
            0,
            0,
            "2026-06-22T00:00:00",
            "2026-06-22T00:00:00",
        ),
    )
    job_id = cursor.lastrowid
    conn.commit()
    conn.close()

    assert job_id is not None
    response = api_client.put(
        f"/api/v1/jobs/{job_id}/files/bom/chunks/0",
        content=b"chunk content data",
        headers={"X-Total-Chunks": "1"},
    )
    assert response.status_code == 200
    data = response.json()
    assert data["received"] == 18
    assert data["chunk_index"] == 0
    assert data["total_chunks"] == 1


def test_complete_file_upload(
    api_client: TestClient,
    temp_db_path: str,
    mock_storage_path: Path,
) -> None:
    """Test finalizing chunk upload and assembling file."""
    # Seed a job
    conn = sqlite3.connect(temp_db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            "awaiting_upload",
            0,
            0,
            0,
            0,
            0,
            "2026-06-22T00:00:00",
            "2026-06-22T00:00:00",
        ),
    )
    job_id = cursor.lastrowid
    conn.commit()
    conn.close()

    assert job_id is not None

    # 1. Upload a chunk first
    api_client.put(
        f"/api/v1/jobs/{job_id}/files/bom/chunks/0",
        content=b"file content",
        headers={"X-Total-Chunks": "1"},
    )

    # 2. Complete upload
    response = api_client.post(f"/api/v1/jobs/{job_id}/files/bom/complete")
    assert response.status_code == 200
    data = response.json()
    assert data["role"] == "bom"
    assert data["file_size"] == 12
    assert "bom.xlsx" in data["file_path"]


# ----------------------------------------------------------------------
# Results Endpoints Tests
# ----------------------------------------------------------------------


def test_download_results_not_ready(api_client: TestClient, temp_db_path: str) -> None:
    """Test downloading results when job status is not finished."""
    # Seed a job in processing state
    conn = sqlite3.connect(temp_db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        ("processing", 0, 0, 0, 1, 1, "2026-06-22T00:00:00", "2026-06-22T00:00:00"),
    )
    job_id = cursor.lastrowid
    conn.commit()
    conn.close()

    assert job_id is not None

    # Try diff download
    response = api_client.get(f"/api/v1/jobs/{job_id}/results/diff")
    assert response.status_code == 409
    assert response.json()["error"]["code"] == "RESULTS_NOT_READY"

    # Try cards download
    response2 = api_client.get(f"/api/v1/jobs/{job_id}/results/cards")
    assert response2.status_code == 409
    assert response2.json()["error"]["code"] == "RESULTS_NOT_READY"


def test_download_results_success(
    api_client: TestClient,
    temp_db_path: str,
    mock_storage_path: Path,
) -> None:
    """Test downloading results successfully when files exist and job is done."""
    # Seed a job in done state
    conn = sqlite3.connect(temp_db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        ("done", 0, 0, 0, 1, 1, "2026-06-22T00:00:00", "2026-06-22T00:00:00"),
    )
    job_id = cursor.lastrowid
    conn.commit()
    conn.close()

    assert job_id is not None

    # Write dummy files to job directory
    job_dir = mock_storage_path / str(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)
    diff_file = job_dir / "diff.xlsx"
    diff_file.write_bytes(b"dummy excel report data")
    cards_file = job_dir / "translated_cards.zip"
    cards_file.write_bytes(b"dummy zip cards data")

    # Download diff
    response = api_client.get(f"/api/v1/jobs/{job_id}/results/diff")
    assert response.status_code == 200
    assert response.content == b"dummy excel report data"
    assert (
        response.headers["content-type"]
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Download cards
    response2 = api_client.get(f"/api/v1/jobs/{job_id}/results/cards")
    assert response2.status_code == 200
    assert response2.content == b"dummy zip cards data"
    assert response2.headers["content-type"] == "application/zip"
