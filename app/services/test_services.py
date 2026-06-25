import sqlite3
from collections.abc import Generator
from pathlib import Path
from unittest.mock import patch

import aiosqlite
import pytest

from app.core.config import get_settings
from app.core.exceptions import (
    ChunkCorruptedError,
    FileUploadError,
    JobCreationError,
    JobNotFoundError,
    JobStateError,
    ResultsNotReadyError,
)
from app.db import async_repository
from app.db.database import Base
from app.services.file_service import FileService
from app.services.job_creation_service import JobCreationService
from app.services.job_processing_service import JobProcessingService
from app.services.result_service import ResultService


@pytest.fixture
def temp_db_path(tmp_path: Path) -> Generator[str, None, None]:
    """Provides a temporary SQLite database path with initial schema."""
    db_file = tmp_path / "test_services.db"
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


# ----------------------------------------------------------------------
# JobCreationService Tests
# ----------------------------------------------------------------------


@pytest.mark.asyncio
async def test_job_creation_success(temp_db_path: str) -> None:
    """Test successful job creation."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        job = await JobCreationService.create(async_db)
        assert job["id"] > 0
        assert job["status"] == "awaiting_upload"


@pytest.mark.asyncio
async def test_job_creation_failure(temp_db_path: str) -> None:
    """Test job creation raises JobCreationError when get_job returns None."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        with patch("app.services.job_creation_service.get_job", return_value=None):
            with pytest.raises(JobCreationError):
                await JobCreationService.create(async_db)


# ----------------------------------------------------------------------
# FileService Tests
# ----------------------------------------------------------------------


@pytest.mark.asyncio
async def test_file_service_upload_chunk_validation(temp_db_path: str) -> None:
    """Test role and job validation in upload_chunk."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        service = FileService(async_db)

        # 1. Invalid role
        with pytest.raises(FileUploadError):
            await service.upload_chunk(1, "invalid_role", 0, b"data", 1)

        # 2. Job not found
        with pytest.raises(JobNotFoundError):
            await service.upload_chunk(999, "bom", 0, b"data", 1)


@pytest.mark.asyncio
async def test_file_service_upload_chunk_status_check(
    temp_db_path: str,
) -> None:
    """Test upload_chunk fails if job status is not 'awaiting_upload'."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        job_id = await async_repository.create_job(async_db)
        await async_repository.update_job_status(async_db, job_id, "processing", None)

        service = FileService(async_db)
        with pytest.raises(JobStateError):
            await service.upload_chunk(job_id, "bom", 0, b"data", 1)


@pytest.mark.asyncio
async def test_file_service_upload_chunk_success(
    temp_db_path: str,
    mock_storage_path: Path,
) -> None:
    """Test successful chunk upload with idempotency check."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        job_id = await async_repository.create_job(async_db)
        service = FileService(async_db)

        # Upload first time (writes chunk)
        res = await service.upload_chunk(job_id, "bom", 0, b"hello", 1)
        assert res.received == 5
        assert res.chunk_index == 0
        assert res.total_chunks == 1

        # Upload second time with same size (idempotent no-op)
        res2 = await service.upload_chunk(job_id, "bom", 0, b"hello", 1)
        assert res2.received == 5

        # Upload with different size (corrupted error)
        with pytest.raises(ChunkCorruptedError):
            await service.upload_chunk(job_id, "bom", 0, b"hello world", 1)


@pytest.mark.asyncio
async def test_file_service_complete_file_upload_errors(
    temp_db_path: str,
) -> None:
    """Test error handling in complete_file_upload."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        service = FileService(async_db)

        # 1. Invalid role
        with pytest.raises(FileUploadError):
            await service.complete_file_upload(1, "invalid_role")

        # 2. Job not found
        with pytest.raises(JobNotFoundError):
            await service.complete_file_upload(999, "bom")

        # 3. No chunks found
        job_id = await async_repository.create_job(async_db)
        with pytest.raises(FileUploadError):
            await service.complete_file_upload(job_id, "bom")


@pytest.mark.asyncio
async def test_file_service_complete_file_upload_success(
    temp_db_path: str,
    mock_storage_path: Path,
) -> None:
    """Test successful assembly of chunks and database updates."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        job_id = await async_repository.create_job(async_db)
        service = FileService(async_db)

        # Upload 2 chunks
        await service.upload_chunk(job_id, "bom", 0, b"first ", 2)
        await service.upload_chunk(job_id, "bom", 1, b"second", 2)

        # Complete upload
        res = await service.complete_file_upload(job_id, "bom")
        assert res.role == "bom"
        assert res.file_size == 12
        assert res.file_path.endswith("bom.xlsx")

        # Verify output content
        output_path = Path(res.file_path)
        assert output_path.exists()
        assert output_path.read_bytes() == b"first second"

        # Verify database updated
        job = await async_repository.get_job(async_db, job_id)
        assert job is not None
        assert job["bom_uploaded"] is True
        assert job["bom_path"] == str(output_path)


# ----------------------------------------------------------------------
# JobProcessingService Tests
# ----------------------------------------------------------------------


@pytest.mark.asyncio
async def test_job_processing_transition_errors(
    temp_db_path: str,
) -> None:
    """Test errors when transitioning to processing."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        # 1. Job not found
        with pytest.raises(JobNotFoundError):
            await JobProcessingService.transition_to_processing(async_db, 999)

        # 2. Files not uploaded
        job_id = await async_repository.create_job(async_db)
        with pytest.raises(JobStateError):
            await JobProcessingService.transition_to_processing(async_db, job_id)

        # 3. Invalid status (already processing)
        await async_repository.update_file_upload(
            async_db, job_id, "bom", "/data/bom.xlsx", True
        )
        await async_repository.update_file_upload(
            async_db, job_id, "archive", "/data/archive.zip", True
        )
        await async_repository.update_job_status(
            async_db, job_id, "processing", "unpacking"
        )
        with pytest.raises(JobStateError):
            await JobProcessingService.transition_to_processing(async_db, job_id)


@pytest.mark.asyncio
async def test_job_processing_transition_success(
    temp_db_path: str,
) -> None:
    """Test successful transition to processing state."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        job_id = await async_repository.create_job(async_db)
        await async_repository.update_file_upload(
            async_db, job_id, "bom", "/data/bom.xlsx", True
        )
        await async_repository.update_file_upload(
            async_db, job_id, "archive", "/data/archive.zip", True
        )

        state = await JobProcessingService.transition_to_processing(async_db, job_id)
        assert state["status"] == "processing"
        assert state["stage"] == "unpacking"


def test_job_processing_dispatch() -> None:
    """Test that dispatch_processing triggers successfully (no-op placeholder)."""
    JobProcessingService.dispatch_processing(1)


# ----------------------------------------------------------------------
# ResultService Tests
# ----------------------------------------------------------------------


@pytest.mark.asyncio
async def test_result_service_validate_job_ready(
    temp_db_path: str,
) -> None:
    """Test job readiness validation."""
    async with aiosqlite.connect(temp_db_path) as async_db:
        # 1. Job not found
        with pytest.raises(JobNotFoundError):
            await ResultService.validate_job_ready(async_db, 999)

        # 2. Not ready (status is processing)
        job_id = await async_repository.create_job(async_db)
        await async_repository.update_job_status(async_db, job_id, "processing", None)
        with pytest.raises(ResultsNotReadyError):
            await ResultService.validate_job_ready(async_db, job_id)

        # 3. Ready (status is done)
        await async_repository.update_job_status(async_db, job_id, "done", None)
        await ResultService.validate_job_ready(async_db, job_id)

        # 4. Ready (status is error)
        await async_repository.update_job_status(async_db, job_id, "error", None)
        await ResultService.validate_job_ready(async_db, job_id)


def test_result_service_get_result_path(mock_storage_path: Path) -> None:
    """Test get_result_path returns file path or raises if missing."""
    job_id = 777
    result_type = "diff"

    # 1. File missing
    with pytest.raises(ResultsNotReadyError):
        ResultService.get_result_path(job_id, result_type)

    # 2. File exists
    job_dir = mock_storage_path / str(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)
    diff_file = job_dir / "diff.xlsx"
    diff_file.write_text("dummy sheet content")

    path = ResultService.get_result_path(job_id, result_type)
    assert path == diff_file
