from fastapi import APIRouter

from . import files, health, jobs, results

router = APIRouter(prefix="/api/v1")

router.include_router(jobs.router)
router.include_router(files.router)
router.include_router(results.router)
router.include_router(health.router)
