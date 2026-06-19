from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.database import get_db
from app.models.user import User
from app.models.member import Member
from app.schemas.member import MemberOut, MemberCreate, MemberListItem, MemberListOut
from app.utils.auth import get_current_user

router = APIRouter(prefix="/api/users", tags=["用户"])


@router.post("/members", response_model=MemberOut)
async def create_member(
    data: MemberCreate,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    member = Member(user_id=current_user.id, name=data.name)
    db.add(member)
    await db.commit()
    await db.refresh(member)
    return MemberOut.model_validate(member)


@router.get("/members", response_model=MemberListOut)
async def list_my_members(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    items = []
    for m in current_user.members:
        item = MemberListItem(
            id=m.id,
            name=m.name,
            member_type=m.member_type,
            status=m.status,
            dim_experience=m.dim_experience,
            dim_frequency=m.dim_frequency,
            dim_equipment=m.dim_equipment,
            dim_injury=m.dim_injury,
            dim_lifestyle=m.dim_lifestyle,
            dim_psychology=m.dim_psychology,
            dim_training_history=m.dim_training_history,
            dim_special_needs=m.dim_special_needs,
            dim_female_specific=m.dim_female_specific,
            dim_motivation=m.dim_motivation,
            coach_notes=m.coach_notes,
            created_at=m.created_at,
            questionnaire_submitted=m.questionnaire is not None,
            plan_ready=m.plan is not None and m.plan.status == "ready",
            has_plan=m.plan is not None,
        )
        items.append(item)

    return MemberListOut(members=sorted(items, key=lambda x: x.created_at, reverse=True), total=len(items))
