import aiosqlite

from app.core.exceptions import JobNotFoundError, JobStateError
from app.db.async_repository import get_job, try_start_processing


class JobProcessingService:
    """Coordinates the job processing lifecycle.

    Responsibilities are split into two phases:
    1. transition_to_processing — atomically validate and persist the new state
    2. dispatch_processing      — trigger asynchronous processing task
    """

    @staticmethod
    async def transition_to_processing(
        db: aiosqlite.Connection,
        job_id: int,
    ) -> dict[str, str]:
        """Atomically validate preconditions and transition job status to 'processing'.

        If preconditions are not met, raises corresponding error (404 or 409).

        Returns:
            The new status and stage written to the database.
        """
        success = await try_start_processing(db, job_id)
        if not success:
            job = await get_job(db, job_id)
            if job is None:
                raise JobNotFoundError(f"Job {job_id} not found")
            if job["status"] != "awaiting_upload":
                raise JobStateError(
                    f"Job {job_id} is in status '{job['status']}', expected 'awaiting_upload'"
                )
            if not job["bom_uploaded"] or not job["archive_uploaded"]:
                raise JobStateError(
                    f"Job {job_id} cannot be started: BOM and archive must be fully uploaded"
                )
            raise JobStateError(f"Job {job_id} cannot be started")
        return {"status": "processing", "stage": "unpacking"}

    @staticmethod
    def dispatch_processing(job_id: int) -> None:
        """Trigger the Celery unpack task asynchronously.

        This is intentionally a sync method — it only enqueues work
        without awaiting the result.
        """
        # TODO: Trigger Celery task unpack.delay(job_id)
        # from app.worker.tasks.unpack import unpack
        # unpack.delay(job_id)
