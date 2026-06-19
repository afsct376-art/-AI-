import uuid
from datetime import datetime
from sqlalchemy import String, DateTime, Integer, Float, Text, ForeignKey, func
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.database import Base


class Member(Base):
    __tablename__ = "members"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    user_id: Mapped[str] = mapped_column(String(36), ForeignKey("users.id"))
    name: Mapped[str] = mapped_column(String(32))
    member_type: Mapped[str] = mapped_column(String(10), default="")  # TYPE_A ~ TYPE_G
    status: Mapped[str] = mapped_column(String(20), default="pending")  # pending / questionnaire_done / generating / completed

    # 10-dimension feature vector (stored as JSON string for flexibility)
    dim_experience: Mapped[int] = mapped_column(Integer, default=0)
    dim_frequency: Mapped[int] = mapped_column(Integer, default=0)
    dim_equipment: Mapped[int] = mapped_column(Integer, default=0)
    dim_injury: Mapped[int] = mapped_column(Integer, default=0)
    dim_lifestyle: Mapped[int] = mapped_column(Integer, default=0)
    dim_psychology: Mapped[int] = mapped_column(Integer, default=0)
    dim_training_history: Mapped[int] = mapped_column(Integer, default=0)
    dim_special_needs: Mapped[int] = mapped_column(Integer, default=0)
    dim_female_specific: Mapped[int] = mapped_column(Integer, default=0)
    dim_motivation: Mapped[int] = mapped_column(Integer, default=0)

    coach_notes: Mapped[str] = mapped_column(Text, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now(), onupdate=func.now())

    user: Mapped["User"] = relationship(back_populates="members")
    questionnaire: Mapped["QuestionnaireResponse"] = relationship(back_populates="member", uselist=False, cascade="all, delete-orphan")
    plan: Mapped["TrainingPlan"] = relationship(back_populates="member", uselist=False, cascade="all, delete-orphan")
