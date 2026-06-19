from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class UserLogin(BaseModel):
    code: str
    nickname: Optional[str] = ""
    avatar_url: Optional[str] = ""


class UserOut(BaseModel):
    id: str
    nickname: str
    avatar_url: str
    phone: str
    is_coach: bool
    created_at: datetime

    class Config:
        from_attributes = True


class TokenOut(BaseModel):
    access_token: str
    token_type: str = "bearer"
    user: UserOut


class UserUpdate(BaseModel):
    nickname: Optional[str] = None
    avatar_url: Optional[str] = None
    phone: Optional[str] = None
