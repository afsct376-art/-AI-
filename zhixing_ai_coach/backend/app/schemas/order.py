from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class CreateOrderOut(BaseModel):
    order_id: str
    order_no: str
    total_fee: int
    prepay_id: str
    pay_params: dict  # WeChat mini-program pay parameters


class OrderOut(BaseModel):
    id: str
    order_no: str
    total_fee: int
    status: str
    paid_at: Optional[datetime]
    created_at: datetime

    class Config:
        from_attributes = True


class PayNotify(BaseModel):
    return_code: str = "SUCCESS"
    return_msg: str = "OK"
