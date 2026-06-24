class BurlakError(Exception):
    """Base exception for all application errors."""

    code: str = "INTERNAL_ERROR"
    status_code: int = 500


class JobNotFoundError(BurlakError):
    code = "JOB_NOT_FOUND"
    status_code = 404


class JobStateError(BurlakError):
    code = "INVALID_JOB_STATE"
    status_code = 409


class FileUploadError(BurlakError):
    code = "FILE_UPLOAD_ERROR"
    status_code = 422


class ChunkCorruptedError(FileUploadError):
    code = "CHUNK_CORRUPTED"
    status_code = 422


class JobCreationError(BurlakError):
    code = "JOB_CREATION_ERROR"
    status_code = 500


class ResultsNotReadyError(BurlakError):
    code = "RESULTS_NOT_READY"
    status_code = 409


class StorageError(BurlakError):
    code = "STORAGE_ERROR"
    status_code = 500