import aiosqlite

from app.core.exceptions import JobNotFoundError, JobStateError
from app.db.async_repository import get_job, update_job_status


class JobProcessingService:
    """Coordinates the job processing lifecycle.

    Responsibilities are split into three distinct phases:
    1. validate_can_start  — check preconditions (SRP: validation only)
    2. transition_to_processing — persist the new state (SRP: mutation only)
    3. dispatch_processing  — trigger async work (SRP: side-effect only)
    """

    @staticmethod
    async def validate_can_start(
        db: aiosqlite.Connection,
        job_id: int,
    ) -> None:
        """Validate that a job is ready to start processing.

        Checks:
        - Job exists
        - Job is in 'awaiting_upload' status
        - Both BOM and archive files have been uploaded

        Raises:
            JobNotFoundError: Job does not exist.
            JobStateError: Job is not in the correct state or files are missing.
        """
        job = await get_job(db, job_id)
        if job is None:
            raise JobNotFoundError(f"Job {job_id} not found")

        if job["status"] != "awaiting_upload":
            raise JobStateError(
                f"Job {job_id} is in status '{job['status']}', "
                f"expected 'awaiting_upload'"
            )

        if not job["bom_uploaded"]:
            raise JobStateError(f"Job {job_id}: BOM file not uploaded yet")

        if not job["archive_uploaded"]:
            raise JobStateError(f"Job {job_id}: Archive file not uploaded yet")

    @staticmethod
    async def transition_to_processing(
        db: aiosqlite.Connection,
        job_id: int,
    ) -> dict:
        """Persist the status transition to 'processing' with stage 'unpacking'.

        This is a pure state-mutation operation with no validation or side effects.

        Returns:
            The actual status and stage written to the database.
        """
        await update_job_status(db, job_id, "processing", "unpacking")
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
