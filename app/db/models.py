from typing import Any
from sqlalchemy.orm import Mapped, mapped_column, relationship
from datetime import datetime, timezone
from sqlalchemy import ForeignKey, JSON

from .database import Base


class Jobs(Base):
    __tablename__ = "jobs"

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    status: Mapped[str] = mapped_column(nullable=False, default="awaiting_upload")
    stage: Mapped[str] = mapped_column(nullable=True)

    total: Mapped[int] = mapped_column(nullable=False, default=0)
    processed: Mapped[int] = mapped_column(nullable=False, default=0)
    failed: Mapped[int] = mapped_column(nullable=False, default=0)

    mapping_config: Mapped[dict[str, Any] | None] = mapped_column(JSON, nullable=True)

    bom_path: Mapped[str | None] = mapped_column(nullable=True)
    archive_path: Mapped[str | None] = mapped_column(nullable=True)
    bom_uploaded: Mapped[bool] = mapped_column(nullable=False, default=False)
    archive_uploaded: Mapped[bool] = mapped_column(nullable=False, default=False)

    created_at: Mapped[datetime] = mapped_column(
        nullable=False, default=lambda: datetime.now(timezone.utc)
    )
    updated_at: Mapped[datetime] = mapped_column(
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )

    cards: Mapped[list["Cards"]] = relationship(
        back_populates="job", cascade="all, delete-orphan"
    )


class Cards(Base):
    __tablename__ = "cards"

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    job_id: Mapped[int] = mapped_column(
        ForeignKey("jobs.id", ondelete="CASCADE"), nullable=False
    )
    card_path: Mapped[str] = mapped_column(nullable=False)
    status: Mapped[str] = mapped_column(
        nullable=False, default="pending"
    )  # pending, success, failed
    error_message: Mapped[str | None] = mapped_column(nullable=True)

    created_at: Mapped[datetime] = mapped_column(
        nullable=False, default=lambda: datetime.now(timezone.utc)
    )
    updated_at: Mapped[datetime] = mapped_column(
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )

    job: Mapped["Jobs"] = relationship(back_populates="cards")
