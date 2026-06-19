from pydantic import BaseModel
from typing import Optional, Any
from datetime import datetime


class QuestionnaireSubmit(BaseModel):
    member_id: str
    raw_data: dict[str, Any]


class QuestionnaireOut(BaseModel):
    id: str
    member_id: str
    version: str
    raw_data: dict[str, Any]
    created_at: datetime

    class Config:
        from_attributes = True
