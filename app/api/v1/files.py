import aiosqlite
from fastapi import APIRouter, Depends, Header, Request

from app.db.database import get_async_db
from app.schemas.file import ChunkUploadResponse, FileCompleteResponse
from app.services.file_service import FileService

router = APIRouter(tags=["files"])


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
    body = await request.body()
    service = FileService(db)
    return await service.upload_chunk(
        job_id=job_id,
        role=role,
        chunk_index=n,
        data=body,
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
    service = FileService(db)
    return await service.complete_file_upload(
        job_id=job_id,
        role=role,
    )
