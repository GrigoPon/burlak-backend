from fastapi import APIRouter, Depends, status

import aiosqlite

from app.core.exceptions import JobNotFoundError, JobStateError
from app.db.async_repository import create_job, get_job, update_job_status
from app.db.database import get_async_db
from app.schemas.job import JobCreateResponse, JobStatusResponse

router = APIRouter(tags=["jobs"])


@router.post("/jobs", status_code=status.HTTP_201_CREATED)
async def create_new_job(db: aiosqlite.Connection = Depends(get_async_db)) -> JobCreateResponse:
    """Create a new job in awaiting_upload status.

    Returns the job ID, initial status, and creation timestamp.
    """
    job_id = await create_job(db)
    job = await get_job(db, job_id)
    if job is None:
        raise JobNotFoundError(f"Job {job_id} not found after creation")

    return JobCreateResponse(
        id=job["id"],
        status=job["status"],
        created_at=job["created_at"],
    )


@router.get("/jobs/{job_id}")
async def get_job_status(
    job_id: int,
    db: aiosqlite.Connection = Depends(get_async_db),
) -> JobStatusResponse:
    """Get job status and progress.

    Returns current status, stage, and processing counters.
    """
    job = await get_job(db, job_id)
    if job is None:
        raise JobNotFoundError(f"Job {job_id} not found")

    return JobStatusResponse(
        id=job["id"],
        status=job["status"],
        stage=job.get("stage"),
        total=job["total"],
        processed=job["processed"],
        failed=job["failed"],
        bom_uploaded=bool(job["bom_uploaded"]),
        archive_uploaded=bool(job["archive_uploaded"]),
        created_at=job["created_at"],
        updated_at=job["updated_at"],
    )


@router.post("/jobs/{job_id}/start", status_code=status.HTTP_202_ACCEPTED)
async def start_job_processing(
    job_id: int,
    db: aiosqlite.Connection = Depends(get_async_db),
) -> dict:
    """Start processing a job.

    Validates that:
    - Job exists and is in 'awaiting_upload' status
    - Both BOM and archive files have been uploaded

    Transitions status to 'processing' with stage 'unpacking'
    and triggers the Celery unpack task.
    """
    job = await get_job(db, job_id)
    if job is None:
        raise JobNotFoundError(f"Job {job_id} not found")

    if job["status"] != "awaiting_upload":
        raise JobStateError(
            f"Job {job_id} is in status '{job['status']}', expected 'awaiting_upload'"
        )

    if not job["bom_uploaded"]:
        raise JobStateError(f"Job {job_id}: BOM file not uploaded yet")

    if not job["archive_uploaded"]:
        raise JobStateError(f"Job {job_id}: Archive file not uploaded yet")

    # Update status to processing
    await update_job_status(db, job_id, "processing", "unpacking")

    # TODO: Trigger Celery task unpack.delay(job_id)
    # from app.worker.tasks.unpack import unpack
    # unpack.delay(job_id)

    return {
        "message": "Job processing started",
        "job_id": job_id,
        "status": "processing",
        "stage": "unpacking",
    }