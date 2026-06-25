from typing import Any

import aiosqlite

from app.core.exceptions import JobCreationError
from app.db.async_repository import create_job, get_job


class JobCreationService:
    """Creates a new job and validates the result."""

    @staticmethod
    async def create(db: aiosqlite.Connection) -> dict[str, Any]:
        job_id = await create_job(db)
        job = await get_job(db, job_id)
        if job is None:
            raise JobCreationError(f"Failed to create job {job_id} in the database")
        return job
