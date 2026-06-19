import uuid
from datetime import datetime
from sqlalchemy import String, DateTime, Text, Integer, ForeignKey, func
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.database import Base


class TrainingPlan(Base):
    __tablename__ = "training_plans"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    member_id: Mapped[str] = mapped_column(String(36), ForeignKey("members.id"), unique=True)
    version: Mapped[str] = mapped_column(String(10), default="V1")

    # AI-generated plan content (markdown)
    raw_content: Mapped[str] = mapped_column(Text, default="")
    coach_edited_content: Mapped[str] = mapped_column(Text, default="")

    # Plan metadata
    plan_type: Mapped[str] = mapped_column(String(20), default="")
    split_type: Mapped[str] = mapped_column(String(50), default="")
    weekly_frequency: Mapped[int] = mapped_column(Integer, default=3)
    session_duration: Mapped[int] = mapped_column(Integer, default=60)

    status: Mapped[str] = mapped_column(String(20), default="generating")  # generating / ready / delivered
    coach_notes: Mapped[str] = mapped_column(Text, default="")
    delivered_at: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now(), onupdate=func.now())

    member: Mapped["Member"] = relationship(back_populates="plan")
