import json
from fastapi import APIRouter, Depends, HTTPException, BackgroundTasks
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.database import get_db
from app.models.user import User
from app.models.member import Member
from app.models.plan import TrainingPlan
from app.schemas.plan import PlanOut, PlanRegenerate
from app.services.plan_generator import create_plan_for_member, build_member_profile
from app.services.decision_tree import classify
from app.utils.auth import get_current_user

router = APIRouter(prefix="/api/plans", tags=["方案"])


async def _generate_plan_task(member_id: str, db: AsyncSession):
    result = await db.execute(select(Member).where(Member.id == member_id))
    member = result.scalar_one_or_none()
    if member is None or member.questionnaire is None:
        return

    q_data = json.loads(member.questionnaire.raw_data)
    profile = build_member_profile(q_data)
    classification = classify(profile)

    member_info = {
        "name": member.name,
        "frequency": profile.frequency,
        "goals": q_data.get("Q9", ""),
        "goal_details": q_data.get("Q10", ""),
    }

    plan_content = await create_plan_for_member(member_info, q_data)

    plan = member.plan
    if plan is None:
        plan = TrainingPlan(
            member_id=member.id,
            raw_content=plan_content,
            plan_type=classification["member_type"],
            split_type=classification["split"],
            weekly_frequency=profile.frequency,
            status="ready",
        )
        db.add(plan)
    else:
        plan.raw_content = plan_content
        plan.plan_type = classification["member_type"]
        plan.split_type = classification["split"]
        plan.status = "ready"

    member.member_type = classification["member_type"]
    member.dim_experience = profile.experience
    member.dim_frequency = profile.frequency
    member.dim_equipment = profile.equipment
    member.dim_injury = profile.injury
    member.dim_lifestyle = profile.lifestyle
    member.dim_psychology = profile.psychology
    member.dim_training_history = profile.training_history
    member.dim_special_needs = profile.special_needs
    member.dim_female_specific = profile.female_specific
    member.dim_motivation = profile.motivation
    member.status = "completed"

    await db.commit()


@router.get("/{member_id}", response_model=PlanOut)
async def get_plan(
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
    if member.plan is None:
        raise HTTPException(status_code=404, detail="方案尚未生成")

    p = member.plan
    return PlanOut(
        id=p.id,
        member_id=p.member_id,
        version=p.version,
        plan_type=p.plan_type,
        split_type=p.split_type,
        weekly_frequency=p.weekly_frequency,
        session_duration=p.session_duration,
        status=p.status,
        raw_content=p.raw_content,
        coach_edited_content=p.coach_edited_content,
        coach_notes=p.coach_notes,
        delivered_at=p.delivered_at,
        created_at=p.created_at,
    )


@router.post("/generate", response_model=PlanOut)
async def generate_plan(
    data: PlanRegenerate,
    background_tasks: BackgroundTasks,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Member).where(Member.id == data.member_id, Member.user_id == current_user.id)
    )
    member = result.scalar_one_or_none()
    if member is None:
        raise HTTPException(status_code=404, detail="会员不存在")
    if member.questionnaire is None:
        raise HTTPException(status_code=400, detail="请先提交问卷")

    if member.plan is None:
        plan = TrainingPlan(member_id=member.id, status="generating", coach_notes=data.coach_notes)
        db.add(plan)
        await db.commit()
        await db.refresh(plan)
    else:
        member.plan.status = "generating"
        member.plan.coach_notes = data.coach_notes
        await db.commit()

    background_tasks.add_task(_generate_plan_task, member.id, db)

    p = member.plan or plan
    return PlanOut(
        id=p.id,
        member_id=p.member_id,
        version=p.version,
        plan_type=p.plan_type,
        split_type=p.split_type,
        weekly_frequency=p.weekly_frequency,
        session_duration=p.session_duration,
        status=p.status,
        raw_content=p.raw_content,
        coach_edited_content=p.coach_edited_content,
        coach_notes=p.coach_notes,
        delivered_at=p.delivered_at,
        created_at=p.created_at,
    )
