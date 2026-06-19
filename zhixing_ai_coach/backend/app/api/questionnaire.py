import json
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.database import get_db
from app.models.user import User
from app.models.member import Member
from app.models.questionnaire import QuestionnaireResponse
from app.schemas.questionnaire import QuestionnaireSubmit, QuestionnaireOut
from app.utils.auth import get_current_user

router = APIRouter(prefix="/api/questionnaire", tags=["问卷"])


@router.post("/submit", response_model=QuestionnaireOut)
async def submit_questionnaire(
    data: QuestionnaireSubmit,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Member).where(Member.id == data.member_id, Member.user_id == current_user.id)
    )
    member = result.scalar_one_or_none()
    if member is None:
        raise HTTPException(status_code=404, detail="会员不存在")

    existing = await db.execute(
        select(QuestionnaireResponse).where(QuestionnaireResponse.member_id == data.member_id)
    )
    if existing.scalar_one_or_none():
        raise HTTPException(status_code=400, detail="问卷已提交")

    q = QuestionnaireResponse(
        member_id=data.member_id,
        raw_data=json.dumps(data.raw_data, ensure_ascii=False),
    )
    db.add(q)
    member.status = "questionnaire_done"
    await db.commit()
    await db.refresh(q)

    return QuestionnaireOut(
        id=q.id,
        member_id=q.member_id,
        version=q.version,
        raw_data=json.loads(q.raw_data),
        created_at=q.created_at,
    )


@router.get("/{member_id}", response_model=QuestionnaireOut)
async def get_questionnaire(
    member_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Member).where(Member.id == member_id, Member.user_id == current_user.id)
    )
    member = result.scalar_one_or_none()
    if member is None:
        raise HTTPException(status_code=404, detail="会员不存在")
    if member.questionnaire is None:
        raise HTTPException(status_code=404, detail="问卷尚未提交")

    q = member.questionnaire
    return QuestionnaireOut(
        id=q.id,
        member_id=q.member_id,
        version=q.version,
        raw_data=json.loads(q.raw_data),
        created_at=q.created_at,
    )
