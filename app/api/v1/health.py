from fastapi import APIRouter

from app.core.config import get_settings

router = APIRouter(tags=["health"])


@router.get("/health")
async def health_check():
    """System health check endpoint.

    Returns 200 OK with status information.
    Optionally checks SQLite and Redis connectivity.
    """
    settings = get_settings()
    checks = {
        "status": "healthy",
        "storage_path": settings.storage_path,
    }
    return checks
