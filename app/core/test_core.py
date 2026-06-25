import shutil
from collections.abc import Generator
from pathlib import Path
from typing import Any

import pytest

from app.core import exceptions
from app.core.config import Settings, get_settings
from app.core.exceptions import BurlakError, ChunkCorruptedError, StorageError
from app.core.storage import (
    assemble_chunks,
    cleanup_chunks,
    cleanup_job,
    get_assembled_path,
    get_chunk_path,
    get_chunks_dir,
    get_job_dir,
    get_results_path,
    verify_chunk,
    write_chunk,
)


@pytest.fixture
def mock_storage_path(tmp_path: Path) -> Generator[Path, None, None]:
    """Fixture to override settings storage_path to a temp directory."""
    settings = get_settings()
    original_path = settings.storage_path
    settings.storage_path = str(tmp_path)
    yield tmp_path
    settings.storage_path = original_path


def test_settings_initialization() -> None:
    """Test that Settings can be instantiated and get_settings returns cached config."""
    settings = get_settings()
    assert isinstance(settings, Settings)
    assert settings.db_url is not None

    # Test lru_cache caching behavior
    settings_again = get_settings()
    assert settings is settings_again


def test_custom_exceptions() -> None:
    """Test the hierarchy and properties of Burlak exceptions."""
    assert issubclass(exceptions.JobNotFoundError, BurlakError)
    assert issubclass(exceptions.JobStateError, BurlakError)
    assert issubclass(exceptions.FileUploadError, BurlakError)
    assert issubclass(exceptions.ChunkCorruptedError, exceptions.FileUploadError)
    assert issubclass(exceptions.JobCreationError, BurlakError)
    assert issubclass(exceptions.ResultsNotReadyError, BurlakError)
    assert issubclass(exceptions.StorageError, BurlakError)

    # Check status codes
    assert exceptions.JobNotFoundError("error").status_code == 404
    assert exceptions.JobStateError("error").status_code == 409
    assert exceptions.ResultsNotReadyError("error").status_code == 409
    assert exceptions.ChunkCorruptedError("error").status_code == 422
    assert exceptions.StorageError("error").status_code == 500


def test_storage_paths(mock_storage_path: Path) -> None:
    """Test storage path construction utilities."""
    job_id = 42
    expected_job_dir = mock_storage_path / str(job_id)
    expected_chunks_dir = expected_job_dir / "chunks"

    assert get_job_dir(job_id) == expected_job_dir
    assert get_chunks_dir(job_id) == expected_chunks_dir
    assert get_chunk_path(job_id, "bom", 0) == expected_chunks_dir / "bom_0.part"
    assert get_assembled_path(job_id, "bom") == expected_job_dir / "bom.xlsx"
    assert get_assembled_path(job_id, "archive") == expected_job_dir / "archive.zip"


def test_verify_chunk_scenarios(mock_storage_path: Path) -> None:
    """Test verify_chunk for non-existent, correct, and corrupted chunks."""
    job_id = 101
    role = "bom"
    n = 0
    size = 100

    # Non-existent
    assert verify_chunk(job_id, role, n, size) is False

    # Create chunk
    write_chunk(job_id, role, n, b"a" * size)
    assert verify_chunk(job_id, role, n, size) is True

    # Corrupted size
    with pytest.raises(ChunkCorruptedError):
        verify_chunk(job_id, role, n, size + 50)


def test_write_chunk_error(mock_storage_path: Path) -> None:
    """Test write_chunk failure case under storage directory creation error."""
    # Force OSError by creating a file where the directory should be
    job_id = 202
    job_dir = get_job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)

    # Create a regular file in place of the chunks directory
    chunks_dir = get_chunks_dir(job_id)
    chunks_dir.write_text("not a directory")

    with pytest.raises(StorageError):
        write_chunk(job_id, "bom", 0, b"some data")


def test_assemble_chunks_flow(mock_storage_path: Path) -> None:
    """Test successful concatenation of chunks and missing chunk error."""
    job_id = 303
    role = "bom"
    chunk_data = [b"hello ", b"world", b"!"]
    total = len(chunk_data)

    for i, data in enumerate(chunk_data):
        write_chunk(job_id, role, i, data)

    assembled_path = assemble_chunks(job_id, role, total)
    assert assembled_path.exists()
    assert assembled_path.read_bytes() == b"hello world!"

    # Missing chunk assembly test
    cleanup_chunks(job_id, role, total)
    with pytest.raises(StorageError):
        assemble_chunks(job_id, role, total)


def test_assemble_chunks_mkdir_error(
    mock_storage_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test assemble_chunks handling of OSError on directory creation."""
    job_id = 304

    # Mock Path.mkdir to raise OSError
    def mock_mkdir(self: Path, *args: Any, **kwargs: Any) -> None:
        raise OSError("Permission denied")

    monkeypatch.setattr(Path, "mkdir", mock_mkdir)
    with pytest.raises(StorageError):
        assemble_chunks(job_id, "bom", 1)


def test_assemble_chunks_write_error(
    mock_storage_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test assemble_chunks handling of OSError on file write operations."""
    job_id = 305
    role = "bom"
    write_chunk(job_id, role, 0, b"data")

    # Mock open to raise OSError on the assembled file path
    original_open = open

    def mock_open(file: Any, mode: str = "r", *args: Any, **kwargs: Any) -> Any:
        if isinstance(file, Path) and file.name == "bom.xlsx" and "w" in mode:
            raise OSError("Write error")
        return original_open(file, mode, *args, **kwargs)

    monkeypatch.setattr("builtins.open", mock_open)
    with pytest.raises(StorageError):
        assemble_chunks(job_id, role, 1)


def test_cleanup_job_flow(mock_storage_path: Path) -> None:
    """Test cleaning up all files associated with a job."""
    job_id = 404
    write_chunk(job_id, "bom", 0, b"data")
    job_dir = get_job_dir(job_id)
    assert job_dir.exists()

    cleanup_job(job_id)
    assert not job_dir.exists()


def test_cleanup_job_error(
    mock_storage_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Test cleanup_job handling of OSError."""
    job_id = 405
    # Create the directory to ensure it is deleted
    job_dir = get_job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)

    def mock_rmtree(path: Any, *args: Any, **kwargs: Any) -> None:
        raise OSError("Cannot delete")

    monkeypatch.setattr(shutil, "rmtree", mock_rmtree)
    with pytest.raises(StorageError):
        cleanup_job(job_id)


def test_get_results_path() -> None:
    """Test result path builders for diff and cards types, and validation of type."""
    job_id = 505
    assert get_results_path(job_id, "diff").name == "diff.xlsx"
    assert get_results_path(job_id, "cards").name == "translated_cards.zip"

    with pytest.raises(ValueError):
        get_results_path(job_id, "invalid_type")
