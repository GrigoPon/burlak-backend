from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.responses import JSONResponse

from app.api.v1.router import router as v1_router
from app.core.exceptions import BurlakError
from app.schemas.job import ErrorDetail, ErrorResponse


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan: startup and shutdown events."""
    # Startup: nothing to initialize yet
    yield
    # Shutdown: cleanup if needed


app = FastAPI(
    title="BOM Verification System API",
    description="Backend API for BOM verification and card translation",
    version="0.1.0",
    lifespan=lifespan,
)


@app.exception_handler(BurlakError)
async def burlak_error_handler(request, exc: BurlakError):
    """Global exception handler for all BurlakError subclasses.

    Returns uniform error response format:
    { "error": { "code": "...", "message": "...", "detail": null } }
    """
    return JSONResponse(
        status_code=exc.status_code,
        content=ErrorResponse(
            error=ErrorDetail(
                code=exc.code,
                message=str(exc),
                detail=None,
            )
        ).model_dump(),
    )


# Register all v1 routers
app.include_router(v1_router)