from fastapi import APIRouter, Depends
from fastapi.responses import StreamingResponse

import aiosqlite

from app.services.result_service import ResultService
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
    await ResultService.validate_job_ready(db, job_id)
    diff_path = ResultService.get_result_path(job_id, "diff")

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
    await ResultService.validate_job_ready(db, job_id)
    cards_path = ResultService.get_result_path(job_id, "cards")

    return StreamingResponse(
        _file_streamer(str(cards_path)),
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="translated_cards_{job_id}.zip"',
            "Content-Length": str(cards_path.stat().st_size),
        },
    )