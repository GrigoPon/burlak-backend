import sqlite3
from dataclasses import dataclass
from datetime import datetime

from app.core.config import get_settings


@dataclass
class ProgressResult:
    is_complete: bool
    processed: int
    failed: int
    total: int


def increment_progress(
    job_id: int, card_path: str, *, success: bool, error_message: str | None = None
) -> ProgressResult:
    """Atomically and idempotently updates card status and increments progress counters in jobs.

    Uses BEGIN IMMEDIATE transaction on raw sqlite3 connection to prevent WAL deadlocks.
    """
    db_path = get_settings().db_url
    now = datetime.utcnow().isoformat()

    conn = sqlite3.connect(db_path, timeout=30.0)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("BEGIN IMMEDIATE")

        # 1. Check current status of the card to ensure idempotency
        cursor = conn.execute(
            "SELECT status FROM cards WHERE job_id = ? AND card_path = ?",
            (job_id, card_path),
        )
        card = cursor.fetchone()
        if not card:
            raise ValueError(f"Card {card_path} not found for job {job_id}")

        current_status = card["status"]
        new_status = "success" if success else "failed"

        # 2. Determine counter adjustments
        processed_delta = 0
        failed_delta = 0

        if current_status == "pending":
            if success:
                processed_delta = 1
            else:
                failed_delta = 1
        elif current_status == "success" and not success:
            processed_delta = -1
            failed_delta = 1
        elif current_status == "failed" and success:
            processed_delta = 1
            failed_delta = -1

        # 3. Update the card record
        conn.execute(
            """
            UPDATE cards
            SET status = ?, error_message = ?, updated_at = ?
            WHERE job_id = ? AND card_path = ?
            """,
            (new_status, error_message, now, job_id, card_path),
        )

        # 4. Update the job counters if there are adjustments
        if processed_delta != 0 or failed_delta != 0:
            conn.execute(
                """
                UPDATE jobs
                SET processed = processed + ?, failed = failed + ?, updated_at = ?
                WHERE id = ?
                """,
                (processed_delta, failed_delta, now, job_id),
            )

        # 5. Fetch current job state
        cursor = conn.execute(
            "SELECT processed, failed, total FROM jobs WHERE id = ?", (job_id,)
        )
        job = cursor.fetchone()
        if not job:
            raise ValueError(f"Job {job_id} not found")

        processed = job["processed"]
        failed = job["failed"]
        total = job["total"]

        is_complete = (processed + failed == total)

        conn.commit()
        return ProgressResult(
            is_complete=is_complete,
            processed=processed,
            failed=failed,
            total=total,
        )
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
