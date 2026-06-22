import os
import shutil
from pathlib import Path

from app.core.config import get_settings
from app.core.exceptions import ChunkCorruptedError, StorageError


def _get_storage_root() -> Path:
    """Returns the root storage path from settings."""
    return Path(get_settings().storage_path)


def get_job_dir(job_id: int) -> Path:
    """Returns /data/{job_id}/ directory path."""
    return _get_storage_root() / str(job_id)


def get_chunks_dir(job_id: int) -> Path:
    """Returns /data/{job_id}/chunks/ directory path."""
    return get_job_dir(job_id) / "chunks"


def get_chunk_path(job_id: int, role: str, n: int) -> Path:
    """Returns path for chunk file: {job_id}/chunks/{role}_{n}.part."""
    return get_chunks_dir(job_id) / f"{role}_{n}.part"


def get_assembled_path(job_id: int, role: str) -> Path:
    """Returns path for assembled file: {job_id}/{role}.ext."""
    ext = "xlsx" if role == "bom" else "zip"
    return get_job_dir(job_id) / f"{role}.{ext}"


def verify_chunk(job_id: int, role: str, n: int, expected_size: int) -> bool:
    """Checks if chunk exists and has correct size (idempotency).

    Returns True if chunk exists with matching size.
    Raises ChunkCorruptedError if chunk exists but size differs.
    """
    path = get_chunk_path(job_id, role, n)
    if not path.exists():
        return False
    actual_size = path.stat().st_size
    if actual_size == expected_size:
        return True
    raise ChunkCorruptedError(
        f"Chunk {role}_{n} for job {job_id} already exists with size "
        f"{actual_size}, expected {expected_size}"
    )


def write_chunk(job_id: int, role: str, n: int, data: bytes) -> None:
    """Writes a chunk to disk. Creates directories if needed."""
    path = get_chunk_path(job_id, role, n)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
    except OSError as e:
        raise StorageError(f"Failed to write chunk {role}_{n} for job {job_id}: {e}") from e


def assemble_chunks(job_id: int, role: str, total_chunks: int) -> Path:
    """Concatenates chunks into final file using streaming copy.

    Validates that all chunks exist before assembly.
    Returns the path to the assembled file.
    """
    output_path = get_assembled_path(job_id, role)
    chunks_dir = get_chunks_dir(job_id)

    # Ensure output directory exists
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        raise StorageError(f"Failed to create output directory for job {job_id}: {e}") from e

    # Stream chunks into the final file (never load entire file in memory)
    try:
        with open(output_path, "wb") as outfile:
            for n in range(total_chunks):
                chunk_path = chunks_dir / f"{role}_{n}.part"
                if not chunk_path.exists():
                    raise StorageError(
                        f"Missing chunk {role}_{n} for job {job_id} during assembly"
                    )
                with open(chunk_path, "rb") as infile:
                    shutil.copyfileobj(infile, outfile)
    except OSError as e:
        raise StorageError(f"Failed to assemble chunks for job {job_id}: {e}") from e

    return output_path


def cleanup_chunks(job_id: int, role: str, total_chunks: int) -> None:
    """Removes chunk files after successful assembly."""
    chunks_dir = get_chunks_dir(job_id)
    for n in range(total_chunks):
        chunk_path = chunks_dir / f"{role}_{n}.part"
        try:
            if chunk_path.exists():
                chunk_path.unlink()
        except OSError:
            pass  # Best-effort cleanup


def cleanup_job(job_id: int) -> None:
    """Removes all job files (for cron cleanup)."""
    job_dir = get_job_dir(job_id)
    try:
        if job_dir.exists():
            shutil.rmtree(job_dir)
    except OSError as e:
        raise StorageError(f"Failed to cleanup job {job_id}: {e}") from e


def get_results_path(job_id: int, result_type: str) -> Path:
    """Returns path to result file.

    Args:
        job_id: The job ID.
        result_type: 'diff' for diff.xlsx, 'cards' for translated_cards.zip.
    """
    if result_type == "diff":
        return get_job_dir(job_id) / "diff.xlsx"
    elif result_type == "cards":
        return get_job_dir(job_id) / "translated_cards.zip"
    else:
        raise ValueError(f"Unknown result type: {result_type}")