from pydantic import BaseModel
from typing import Optional, Any
from datetime import datetime


class MemberCreate(BaseModel):
    name: str


class MemberOut(BaseModel):
    id: str
    name: str
    member_type: str
    status: str
    dim_experience: int
    dim_frequency: int
    dim_equipment: int
    dim_injury: int
    dim_lifestyle: int
    dim_psychology: int
    dim_training_history: int
    dim_special_needs: int
    dim_female_specific: int
    dim_motivation: int
    coach_notes: str
    created_at: datetime

    class Config:
        from_attributes = True


class MemberListItem(MemberOut):
    questionnaire_submitted: bool = False
    plan_ready: bool = False
    has_plan: bool = False


class MemberListOut(BaseModel):
    members: list[MemberListItem]
    total: int
