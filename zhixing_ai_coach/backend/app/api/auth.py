import httpx
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.database import get_db
from app.models.user import User
from app.schemas.user import UserLogin, TokenOut, UserOut, UserUpdate
from app.utils.auth import create_access_token, get_current_user
from app.config import settings

router = APIRouter(prefix="/api/auth", tags=["认证"])


async def wechat_code_to_openid(code: str) -> str:
    """Call WeChat jscode2session API to exchange code for openid."""
    if not settings.WECHAT_APP_ID:
        return code  # fallback: use code as openid in dev mode

    url = "https://api.weixin.qq.com/sns/jscode2session"
    params = {
        "appid": settings.WECHAT_APP_ID,
        "secret": settings.WECHAT_APP_SECRET,
        "js_code": code,
        "grant_type": "authorization_code",
    }
    async with httpx.AsyncClient() as client:
        resp = await client.get(url, params=params)
        data = resp.json()
        if "openid" in data:
            return data["openid"]
        raise HTTPException(status_code=400, detail=f"微信登录失败: {data.get('errmsg', '未知错误')}")


@router.post("/login", response_model=TokenOut)
async def login(data: UserLogin, db: AsyncSession = Depends(get_db)):
    openid = await wechat_code_to_openid(data.code)

    result = await db.execute(select(User).where(User.openid == openid))
    user = result.scalar_one_or_none()

    if user is None:
        user = User(openid=openid, nickname=data.nickname, avatar_url=data.avatar_url)
        db.add(user)
        await db.commit()
        await db.refresh(user)
    else:
        if data.nickname:
            user.nickname = data.nickname
        if data.avatar_url:
            user.avatar_url = data.avatar_url
        await db.commit()

    token = create_access_token(user.id)
    return TokenOut(access_token=token, user=UserOut.model_validate(user))


@router.get("/me", response_model=UserOut)
async def get_me(current_user: User = Depends(get_current_user)):
    return UserOut.model_validate(current_user)


@router.put("/me", response_model=UserOut)
async def update_me(data: UserUpdate, db: AsyncSession = Depends(get_db), current_user: User = Depends(get_current_user)):
    if data.nickname is not None:
        current_user.nickname = data.nickname
    if data.avatar_url is not None:
        current_user.avatar_url = data.avatar_url
    if data.phone is not None:
        current_user.phone = data.phone
    await db.commit()
    await db.refresh(current_user)
    return UserOut.model_validate(current_user)
