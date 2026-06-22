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
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
