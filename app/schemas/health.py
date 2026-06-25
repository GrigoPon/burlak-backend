from pydantic import BaseModel


class HealthChecks(BaseModel):
    database: str
    redis: str
    storage: str


class HealthResponse(BaseModel):
    status: str
    checks: HealthChecks
