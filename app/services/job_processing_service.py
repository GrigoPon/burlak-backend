from app.core.exceptions import JobNotFoundError, JobStateError 
from app.db.async_repository import get_job, update_job_status

import aiosqlite

class JobProcessingService:
    """Start processing a job.
    
    Validates that:
    - Job exists and is in 'awaiting_upload' status
    - Both BOM and archive files have been uploaded

    Transitions status to 'processing' with stage 'unpacking'
    and triggers the Celery unpack task.
    """
    @staticmethod
    async def validate_and_start_processing(
        db: aiosqlite.Connection,
        job_id: int,
    ):
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
        return {"status": "processing", "stage": "unpacking"}
        