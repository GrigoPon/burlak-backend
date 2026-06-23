from pathlib import Path

import aiosqlite

from app.core.exceptions import JobNotFoundError, ResultsNotReadyError
from app.core.storage import get_results_path
from app.db.async_repository import get_job

class ResultService:
    """Service layer for result file operations.
    Encapsulates all business logic and validation
    so that API handlers stay thin.
    """

    VALID_STATUSES = ("done", "error")

    @staticmethod
    async def validate_and_get_path(
        db: aiosqlite.Connection,
        job_id: int,
        result_type: str,
    ) -> Path:
        job = await get_job(db, job_id)
        if job is None:
            raise JobNotFoundError(f"Job {job_id} not found")

        if job["status"] not in ResultService.VALID_STATUSES:
            raise ResultsNotReadyError(
                f"Job {job_id} is in status '{job['status']}', "
                f"expected 'done' or 'error'"
            )

        result_path = get_results_path(job_id, result_type)
        if not result_path.exists():
            raise ResultsNotReadyError(f"Diff file not yet available for job {job_id}")

        return result_path