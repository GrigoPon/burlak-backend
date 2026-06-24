import aiosqlite
from fastapi import APIRouter, Depends
from fastapi.responses import FileResponse

from app.db.database import get_async_db
from app.services.result_service import ResultService

router = APIRouter(tags=["results"])


@router.get("/jobs/{job_id}/results/diff")
async def download_diff(
    job_id: int,
    db: aiosqlite.Connection = Depends(get_async_db),
) -> FileResponse:
    """Download diff.xlsx — the comparison report.

    Job must be in 'done' or 'error' status.
    Returns a streaming binary response.
    """
    await ResultService.validate_job_ready(db, job_id)
    diff_path = ResultService.get_result_path(job_id, "diff")

    return FileResponse(
        path=diff_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"diff_{job_id}.xlsx",
    )


@router.get("/jobs/{job_id}/results/cards")
async def download_cards(
    job_id: int,
    db: aiosqlite.Connection = Depends(get_async_db),
) -> FileResponse:
    """Download translated_cards.zip — all translated operation cards.

    Job must be in 'done' or 'error' status.
    Returns a streaming binary response.
    """
    await ResultService.validate_job_ready(db, job_id)
    cards_path = ResultService.get_result_path(job_id, "cards")

    return FileResponse(
        path=cards_path,
        media_type="application/zip",
        filename=f"translated_cards_{job_id}.zip",
    )
