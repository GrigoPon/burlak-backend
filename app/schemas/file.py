from pydantic import BaseModel


class ChunkUploadResponse(BaseModel):
    """Response for PUT /api/v1/jobs/{job_id}/files/{role}/chunks/{n}"""

    received: int  # bytes received
    chunk_index: int
    total_chunks: int


class FileCompleteResponse(BaseModel):
    """Response for POST /api/v1/jobs/{job_id}/files/{role}/complete"""

    role: str  # bom | archive
    file_size: int
    file_path: str