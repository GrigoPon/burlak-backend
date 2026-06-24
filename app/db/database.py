import aiosqlite
from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

from app.core.config import get_settings

settings = get_settings()

db_url = settings.db_url
if not db_url.startswith("sqlite://"):
    db_url = f"sqlite:///{db_url}"

engine = create_engine(db_url)
SessionLocal = sessionmaker(bind=engine)

Base = declarative_base()


def get_db():
    """Synchronous database session (for Celery workers)."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


async def get_async_db() -> aiosqlite.Connection:
    """Asynchronous database connection (for FastAPI endpoints).

    Provides an aiosqlite connection to the SQLite database.
    The connection is closed when the request is finished.
    """
    db_path = settings.db_url
    if db_path.startswith("sqlite:///"):
        db_path = db_path[len("sqlite:///"):]
    db = await aiosqlite.connect(db_path)
    db.row_factory = aiosqlite.Row
    try:
        yield db
    finally:
        await db.close()
