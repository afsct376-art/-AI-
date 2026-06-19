import uuid
from datetime import datetime
from sqlalchemy import String, DateTime, Integer, Float, ForeignKey, func
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.database import Base


class Order(Base):
    __tablename__ = "orders"

    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    user_id: Mapped[str] = mapped_column(String(36), ForeignKey("users.id"))
    member_id: Mapped[str] = mapped_column(String(36), ForeignKey("members.id"), nullable=True)
    order_no: Mapped[str] = mapped_column(String(64), unique=True)
    total_fee: Mapped[int] = mapped_column(Integer, default=990)  # 单位:分
    status: Mapped[str] = mapped_column(String(20), default="pending")  # pending / paid / refunded

    # WeChat Pay fields
    prepay_id: Mapped[str] = mapped_column(String(128), default="")
    transaction_id: Mapped[str] = mapped_column(String(128), default="")
    paid_at: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(DateTime, server_default=func.now(), onupdate=func.now())

    user: Mapped["User"] = relationship(back_populates="orders")
