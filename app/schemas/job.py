from datetime import datetime
from typing import Any

from pydantic import BaseModel


class JobCreateResponse(BaseModel):
    """Response for POST /api/v1/jobs"""

    id: int
    status: str
    created_at: datetime


class JobStatusResponse(BaseModel):
    """Response for GET /api/v1/jobs/{job_id}"""

    id: int
    status: str  # awaiting_upload | processing | done | error
    stage: (
        str | None
    )  # unpacking | analyzing_mapping | processing_cards | aggregating | packaging
    total: int
    processed: int
    failed: int
    bom_uploaded: bool
    archive_uploaded: bool
    created_at: datetime
    updated_at: datetime


class ErrorResponse(BaseModel):
    """Uniform error response format"""

    error: "ErrorDetail"


class ErrorDetail(BaseModel):
    code: str
    message: str
    detail: Any = None
