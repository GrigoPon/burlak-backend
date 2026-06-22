from fastapi import APIRouter, Depends
from fastapi.responses import StreamingResponse

import aiosqlite

from app.core.exceptions import JobNotFoundError, ResultsNotReadyError
from app.core.storage import get_results_path
from app.db.async_repository import get_job
from app.db.database import get_async_db

router = APIRouter(tags=["results"])


def _file_streamer(file_path: str, chunk_size: int = 65536):
    """Generator that reads a file in chunks for streaming response."""
    with open(file_path, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            yield chunk


@router.get("/jobs/{job_id}/results/diff")
async def download_diff(
    job_id: int,
    db: aiosqlite.Connection = Depends(get_async_db),
):
    """Download diff.xlsx — the comparison report.

    Job must be in 'done' or 'error' status.
    Returns a streaming binary response.
    """
    job = await get_job(db, job_id)
    if job is None:
        raise JobNotFoundError(f"Job {job_id} not found")

    if job["status"] not in ("done", "error"):
        raise ResultsNotReadyError(
            f"Job {job_id} is in status '{job['status']}', "
            f"expected 'done' or 'error'"
        )

    diff_path = get_results_path(job_id, "diff")
    if not diff_path.exists():
        raise ResultsNotReadyError(f"Diff file not yet available for job {job_id}")

    return StreamingResponse(
        _file_streamer(str(diff_path)),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="diff_{job_id}.xlsx"',
            "Content-Length": str(diff_path.stat().st_size),
        },
    )


@router.get("/jobs/{job_id}/results/cards")
async def download_cards(
    job_id: int,
    db: aiosqlite.Connection = Depends(get_async_db),
):
    """Download translated_cards.zip — all translated operation cards.

    Job must be in 'done' or 'error' status.
    Returns a streaming binary response.
    """
    job = await get_job(db, job_id)
    if job is None:
        raise JobNotFoundError(f"Job {job_id} not found")

    if job["status"] not in ("done", "error"):
        raise ResultsNotReadyError(
            f"Job {job_id} is in status '{job['status']}', "
            f"expected 'done' or 'error'"
        )

    cards_path = get_results_path(job_id, "cards")
    if not cards_path.exists():
        raise ResultsNotReadyError(f"Translated cards file not yet available for job {job_id}")

    return StreamingResponse(
        _file_streamer(str(cards_path)),
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="translated_cards_{job_id}.zip"',
            "Content-Length": str(cards_path.stat().st_size),
        },
    )