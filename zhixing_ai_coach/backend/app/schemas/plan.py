from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class PlanOut(BaseModel):
    id: str
    member_id: str
    version: str
    plan_type: str
    split_type: str
    weekly_frequency: int
    session_duration: int
    status: str
    raw_content: str
    coach_edited_content: str
    coach_notes: str
    delivered_at: Optional[datetime]
    created_at: datetime

    class Config:
        from_attributes = True


class PlanRegenerate(BaseModel):
    member_id: str
    coach_notes: Optional[str] = ""
