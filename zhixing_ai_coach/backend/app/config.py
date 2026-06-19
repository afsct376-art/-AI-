from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    APP_NAME: str = "知行AI教练"
    APP_VERSION: str = "0.1.0"
    DEBUG: bool = True

    DATABASE_URL: str = "sqlite+aiosqlite:///./zhixing.db"

    SECRET_KEY: str = "change-me-in-production"
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 60 * 24 * 7

    WECHAT_APP_ID: str = ""
    WECHAT_APP_SECRET: str = ""
    WECHAT_MCH_ID: str = ""
    WECHAT_API_KEY: str = ""
    WECHAT_NOTIFY_URL: str = ""

    AI_API_KEY: str = ""
    AI_API_URL: str = "https://api.anthropic.com/v1/messages"
    AI_MODEL: str = "claude-sonnet-4-20250514"

    PLAN_PRICE: int = 990  # 分, ¥9.90

    class Config:
        env_file = ".env"


settings = Settings()
