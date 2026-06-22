from fastapi import APIRouter, Depends, Header, Request, status

import aiosqlite

from app.core.exceptions import FileUploadError, JobNotFoundError, JobStateError
from app.core.storage import (
    assemble_chunks,
    cleanup_chunks,
    get_assembled_path,
    verify_chunk,
    write_chunk,
)
from app.db.async_repository import get_job, update_file_upload
from app.db.database import get_async_db
from app.schemas.file import ChunkUploadResponse, FileCompleteResponse

router = APIRouter(tags=["files"])

VALID_ROLES = {"bom", "archive"}


@router.put("/jobs/{job_id}/files/{role}/chunks/{n}")
async def upload_chunk(
    job_id: int,
    role: str,
    n: int,
    request: Request,
    x_total_chunks: int = Header(..., alias="X-Total-Chunks"),
    db: aiosqlite.Connection = Depends(get_async_db),
) -> ChunkUploadResponse:
    """Upload a single file chunk.

    Path params:
        role: 'bom' or 'archive'
        n: 0-based chunk index

    Headers:
        X-Total-Chunks: total number of chunks for this file

    Idempotency: if chunk already exists with matching size, returns 200 OK.
    If chunk exists but size differs, raises 422 CHUNK_CORRUPTED.
    """
    # Validate role
    if role not in VALID_ROLES:
        raise FileUploadError(f"Invalid role '{role}'. Must be one of: {', '.join(VALID_ROLES)}")

    # Validate job exists and is in correct state
    job = await get_job(db, job_id)
    if job is None:
        raise JobNotFoundError(f"Job {job_id} not found")
    if job["status"] != "awaiting_upload":
        raise JobStateError(
            f"Job {job_id} is in status '{job['status']}', expected 'awaiting_upload'"
        )

    # Read raw bytes from request body
    body = await request.body()
    received_bytes = len(body)

    # Idempotency check
    if verify_chunk(job_id, role, n, received_bytes):
        return ChunkUploadResponse(
            received=received_bytes,
            chunk_index=n,
            total_chunks=x_total_chunks,
        )

    # Write chunk to disk
    write_chunk(job_id, role, n, body)

    return ChunkUploadResponse(
        received=received_bytes,
        chunk_index=n,
        total_chunks=x_total_chunks,
    )


@router.post("/jobs/{job_id}/files/{role}/complete")
async def complete_file_upload(
    job_id: int,
    role: str,
    db: aiosqlite.Connection = Depends(get_async_db),
) -> FileCompleteResponse:
    """Finalize chunk upload and assemble the file.

    Concatenates all chunks into the final file,
    updates the database, and cleans up chunk files.
    """
    # Validate role
    if role not in VALID_ROLES:
        raise FileUploadError(f"Invalid role '{role}'. Must be one of: {', '.join(VALID_ROLES)}")

    # Validate job exists
    job = await get_job(db, job_id)
    if job is None:
        raise JobNotFoundError(f"Job {job_id} not found")

    # Assemble chunks into final file
    from pathlib import Path

    from app.core.storage import get_chunks_dir

    chunks_dir = get_chunks_dir(job_id)
    chunk_files = sorted(
        chunks_dir.glob(f"{role}_*.part"),
        key=lambda p: int(p.stem.split("_")[-1]),
    )

    if not chunk_files:
        raise FileUploadError(f"No chunks found for role '{role}' in job {job_id}")

    total_chunks = len(chunk_files)
    output_path = assemble_chunks(job_id, role, total_chunks)
    file_size = output_path.stat().st_size

    # Update database
    await update_file_upload(db, job_id, role, str(output_path), True)

    # Clean up chunk files
    cleanup_chunks(job_id, role, total_chunks)

    return FileCompleteResponse(
        role=role,
        file_size=file_size,
        file_path=str(output_path),
    )