import json
from datetime import datetime
from typing import Any

import aiosqlite


async def create_job(db: aiosqlite.Connection) -> int:
    """Creates a new job in the database and returns its ID."""
    now = datetime.utcnow().isoformat()
    async with db.execute(
        """
        INSERT INTO jobs (status, total, processed, failed, bom_uploaded, archive_uploaded, created_at, updated_at)
        VALUES ('awaiting_upload', 0, 0, 0, 0, 0, ?, ?)
        """,
        (now, now),
    ) as cursor:
        await db.commit()
        if cursor.lastrowid is None:
            raise ValueError("Failed to create job: lastrowid is None")
        return cursor.lastrowid


async def get_job(db: aiosqlite.Connection, job_id: int) -> dict[str, Any] | None:
    """Retrieves a job by its ID, parsing JSON fields and converting booleans."""
    db.row_factory = aiosqlite.Row
    async with db.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)) as cursor:
        row = await cursor.fetchone()
        if row:
            data = dict(row)
            if data.get("mapping_config"):
                data["mapping_config"] = json.loads(data["mapping_config"])
            # Convert boolean integer values to actual booleans
            data["bom_uploaded"] = bool(data["bom_uploaded"])
            data["archive_uploaded"] = bool(data["archive_uploaded"])
            return data
        return None


async def update_job_status(
    db: aiosqlite.Connection, job_id: int, status: str, stage: str | None = None
) -> None:
    """Updates the status and stage of a job."""
    now = datetime.utcnow().isoformat()
    await db.execute(
        """
        UPDATE jobs
        SET status = ?, stage = ?, updated_at = ?
        WHERE id = ?
        """,
        (status, stage, now, job_id),
    )
    await db.commit()


async def update_file_upload(
    db: aiosqlite.Connection, job_id: int, role: str, path: str, uploaded: bool
) -> None:
    """Updates file upload state and path for either 'bom' or 'archive' role."""
    now = datetime.utcnow().isoformat()
    uploaded_val = 1 if uploaded else 0
    if role == "bom":
        await db.execute(
            """
            UPDATE jobs
            SET bom_path = ?, bom_uploaded = ?, updated_at = ?
            WHERE id = ?
            """,
            (path, uploaded_val, now, job_id),
        )
    elif role == "archive":
        await db.execute(
            """
            UPDATE jobs
            SET archive_path = ?, archive_uploaded = ?, updated_at = ?
            WHERE id = ?
            """,
            (path, uploaded_val, now, job_id),
        )
    else:
        raise ValueError(f"Invalid file role: {role}")
    await db.commit()


async def create_cards(
    db: aiosqlite.Connection, job_id: int, card_paths: list[str]
) -> None:
    """Creates card records and sets the total card count on the job."""
    if not card_paths:
        return
    now = datetime.utcnow().isoformat()
    cards_data = [(job_id, path, "pending", now, now) for path in card_paths]
    await db.executemany(
        """
        INSERT INTO cards (job_id, card_path, status, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        cards_data,
    )
    await db.execute(
        """
        UPDATE jobs
        SET total = ?, updated_at = ?
        WHERE id = ?
        """,
        (len(card_paths), now, job_id),
    )
    await db.commit()


async def get_failed_cards(
    db: aiosqlite.Connection, job_id: int
) -> list[dict[str, Any]]:
    """Retrieves all failed cards for a specific job."""
    db.row_factory = aiosqlite.Row
    async with db.execute(
        """
        SELECT card_path, error_message
        FROM cards
        WHERE job_id = ? AND status = 'failed'
        """,
        (job_id,),
    ) as cursor:
        rows = await cursor.fetchall()
        return [dict(row) for row in rows]


async def update_mapping_config(
    db: aiosqlite.Connection, job_id: int, mapping_config: dict[str, Any]
) -> None:
    """Updates the mapping config field for a job."""
    now = datetime.utcnow().isoformat()
    config_json = json.dumps(mapping_config)
    await db.execute(
        """
        UPDATE jobs
        SET mapping_config = ?, updated_at = ?
        WHERE id = ?
        """,
        (config_json, now, job_id),
    )
    await db.commit()
