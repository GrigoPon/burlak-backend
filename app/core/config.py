from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    db_url: str = "./dev.db"
    redis_url: str = "redis://redis:6379/0"
    ml_service_url: str = "http://ml-service:8000"
    storage_path: str = "/data"
    chunk_size_bytes: int = 20971520

    model_config = SettingsConfigDict(env_file=".env", extra="ignore")


@lru_cache()
def get_settings() -> Settings:
    return Settings()
