from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.database import get_db
from app.models.user import User
from app.models.member import Member
from app.models.order import Order
from app.schemas.order import CreateOrderOut, OrderOut, PayNotify
from app.services.wechat_pay import generate_order_no, get_miniapp_pay_params
from app.config import settings
from app.utils.auth import get_current_user

router = APIRouter(prefix="/api/payments", tags=["支付"])


@router.post("/create-order", response_model=CreateOrderOut)
async def create_order(
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

    existing = await db.execute(
        select(Order).where(Order.member_id == member_id, Order.status == "paid")
    )
    if existing.scalar_one_or_none():
        raise HTTPException(status_code=400, detail="该会员已付费")

    order_no = generate_order_no()
    order = Order(
        user_id=current_user.id,
        member_id=member_id,
        order_no=order_no,
        total_fee=settings.PLAN_PRICE,
    )
    db.add(order)
    await db.commit()
    await db.refresh(order)

    prepay_id = f"mock_{order_no}"
    order.prepay_id = prepay_id
    await db.commit()

    pay_params = get_miniapp_pay_params(prepay_id)

    return CreateOrderOut(
        order_id=order.id,
        order_no=order_no,
        total_fee=order.total_fee,
        prepay_id=prepay_id,
        pay_params=pay_params,
    )


@router.post("/notify")
async def pay_notify(notify: PayNotify):
    return {"return_code": "SUCCESS", "return_msg": "OK"}


@router.get("/orders", response_model=list[OrderOut])
async def list_orders(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    result = await db.execute(
        select(Order).where(Order.user_id == current_user.id).order_by(Order.created_at.desc())
    )
    orders = result.scalars().all()
    return [
        OrderOut(
            id=o.id, order_no=o.order_no, total_fee=o.total_fee,
            status=o.status, paid_at=o.paid_at, created_at=o.created_at,
        )
        for o in orders
    ]
