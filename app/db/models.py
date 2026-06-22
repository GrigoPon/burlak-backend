from datetime import datetime
from sqlalchemy import Column, Integer, String, Boolean, DateTime, ForeignKey, JSON
from sqlalchemy.orm import relationship

from .database import Base


class Jobs(Base):
    __tablename__ = "jobs"

    id = Column(Integer, primary_key=True, autoincrement=True)
    status = Column(String, nullable=False, default="awaiting_upload")
    stage = Column(String, nullable=True)

    total = Column(Integer, nullable=False, default=0)
    processed = Column(Integer, nullable=False, default=0)
    failed = Column(Integer, nullable=False, default=0)

    mapping_config = Column(JSON, nullable=True)

    bom_path = Column(String, nullable=True)
    archive_path = Column(String, nullable=True)
    bom_uploaded = Column(Boolean, nullable=False, default=False)
    archive_uploaded = Column(Boolean, nullable=False, default=False)

    created_at = Column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at = Column(DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    cards = relationship("Cards", back_populates="job", cascade="all, delete-orphan")


class Cards(Base):
    __tablename__ = "cards"

    id = Column(Integer, primary_key=True, autoincrement=True)
    job_id = Column(Integer, ForeignKey("jobs.id", ondelete="CASCADE"), nullable=False)
    card_path = Column(String, nullable=False)
    status = Column(String, nullable=False, default="pending")  # pending, success, failed
    error_message = Column(String, nullable=True)

    created_at = Column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at = Column(DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    job = relationship("Jobs", back_populates="cards")
