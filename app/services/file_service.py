from pathlib import Path

import aiosqlite

from app.core.exceptions import FileUploadError, JobNotFoundError, JobStateError
from app.core.storage import (
    assemble_chunks,
    cleanup_chunks,
    get_chunks_dir,
    verify_chunk,
    write_chunk,
)
from app.db.async_repository import get_job, update_file_upload
from app.schemas.file import ChunkUploadResponse, FileCompleteResponse

VALID_ROLES = {"bom", "archive"}


class FileService:
    """Encapsulates file upload business logic.

    Separates HTTP concerns from domain logic so that upload behaviour
    can be tested without a running FastAPI server.
    """

    def __init__(self, db: aiosqlite.Connection) -> None:
        self._db = db

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    async def upload_chunk(
        self,
        job_id: int,
        role: str,
        chunk_index: int,
        data: bytes,
        total_chunks: int,
    ) -> ChunkUploadResponse:
        """Upload a single file chunk with idempotency support.

        Raises:
            FileUploadError: Invalid role or chunk-level error.
            JobNotFoundError: Job does not exist.
            JobStateError: Job is not in 'awaiting_upload' state.
        """
        self._validate_role(role)
        await self._ensure_job_awaiting_upload(job_id)

        received_bytes = len(data)

        # Idempotency: if chunk already exists with matching size, skip write.
        if not verify_chunk(job_id, role, chunk_index, received_bytes):
            write_chunk(job_id, role, chunk_index, data)

        return ChunkUploadResponse(
            received=received_bytes,
            chunk_index=chunk_index,
            total_chunks=total_chunks,
        )

    async def complete_file_upload(
        self,
        job_id: int,
        role: str,
    ) -> FileCompleteResponse:
        """Finalise chunk upload and assemble the file.

        Raises:
            FileUploadError: Invalid role or no chunks found.
            JobNotFoundError: Job does not exist.
        """
        self._validate_role(role)

        job = await get_job(self._db, job_id)
        if job is None:
            raise JobNotFoundError(f"Job {job_id} not found")

        # Discover chunks on disk
        chunks_dir = get_chunks_dir(job_id)
        chunk_files: list[Path] = sorted(
            chunks_dir.glob(f"{role}_*.part"),
            key=lambda p: int(p.stem.split("_")[-1]),
        )

        if not chunk_files:
            raise FileUploadError(f"No chunks found for role '{role}' in job {job_id}")

        total_chunks = len(chunk_files)
        output_path = assemble_chunks(job_id, role, total_chunks)
        file_size = output_path.stat().st_size

        await update_file_upload(self._db, job_id, role, str(output_path), True)

        # Best-effort cleanup of chunk files after successful assembly
        cleanup_chunks(job_id, role, total_chunks)

        return FileCompleteResponse(
            role=role,
            file_size=file_size,
            file_path=str(output_path),
        )

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _validate_role(role: str) -> None:
        if role not in VALID_ROLES:
            raise FileUploadError(
                f"Invalid role '{role}'. Must be one of: {', '.join(sorted(VALID_ROLES))}"
            )

    async def _ensure_job_awaiting_upload(self, job_id: int) -> None:
        job = await get_job(self._db, job_id)
        if job is None:
            raise JobNotFoundError(f"Job {job_id} not found")
        if job["status"] != "awaiting_upload":
            raise JobStateError(
                f"Job {job_id} is in status '{job['status']}', "
                f"expected 'awaiting_upload'"
            )
