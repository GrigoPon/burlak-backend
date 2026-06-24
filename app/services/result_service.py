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
    async def validate_job_ready(
        db: aiosqlite.Connection,
        job_id: int,
    ) -> None:
        """Validate that the job exists and is in a terminal state.

        Raises:
            JobNotFoundError: Job does not exist.
            ResultsNotReadyError: Job is not in 'done' or 'error' status.
        """
        job = await get_job(db, job_id)
        if job is None:
            raise JobNotFoundError(f"Job {job_id} not found")

        if job["status"] not in ResultService.VALID_STATUSES:
            raise ResultsNotReadyError(
                f"Job {job_id} is in status '{job['status']}', "
                f"expected 'done' or 'error'"
            )

    @staticmethod
    def get_result_path(job_id: int, result_type: str) -> Path:
        """Build the filesystem path for a result file.

        Args:
            job_id: The job ID.
            result_type: 'diff' for diff.xlsx, 'cards' for translated_cards.zip.

        Returns:
            Path to the result file.

        Raises:
            ResultsNotReadyError: Result file does not exist on disk.
        """
        result_path = get_results_path(job_id, result_type)
        if not result_path.exists():
            raise ResultsNotReadyError(
                f"Result file '{result_type}' not yet available for job {job_id}"
            )
        return result_path
