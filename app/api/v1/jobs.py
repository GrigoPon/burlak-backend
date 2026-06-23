from fastapi import APIRouter, Depends, status

import aiosqlite

from app.services.job_creation_service import JobCreationService
from app.core.exceptions import JobNotFoundError, JobStateError
from app.services.job_processing_service import JobProcessingService
from app.db.async_repository import get_job
from app.db.database import get_async_db
from app.schemas.job import JobCreateResponse, JobStatusResponse

router = APIRouter(tags=["jobs"])


@router.post("/jobs", status_code=status.HTTP_201_CREATED)
async def create_new_job(db: aiosqlite.Connection = Depends(get_async_db)) -> JobCreateResponse:
    """Create a new job in awaiting_upload status.

    Returns the job ID, initial status, and creation timestamp.
    """
    job = await JobCreationService.create(db)

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
    """Start processing a job."""
    
    result = await JobProcessingService.validate_and_start_processing(db, job_id)
    return {
        "message": "Job processing started",
        "job_id": job_id,
        **result,
    }