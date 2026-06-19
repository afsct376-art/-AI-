"""WeChat Pay integration for mini-program.

In production, use wechatpay-python SDK.
This is a simplified version for illustration.
"""
import hashlib
import time
import random
import string
from typing import Any

from app.config import settings


def generate_order_no() -> str:
    ts = int(time.time())
    rand = "".join(random.choices(string.digits, k=6))
    return f"ZX{ts}{rand}"


def generate_nonce_str() -> str:
    return "".join(random.choices(string.ascii_letters + string.digits, k=32))


def get_miniapp_pay_params(prepay_id: str) -> dict[str, Any]:
    """Return the pay parameters needed by WeChat mini-program for wx.requestPayment."""
    if not settings.WECHAT_APP_ID:
        return {
            "timeStamp": str(int(time.time())),
            "nonceStr": generate_nonce_str(),
            "package": f"prepay_id={prepay_id}",
            "signType": "MD5",
            "paySign": "simulated-sign",
        }

    nonce_str = generate_nonce_str()
    timestamp = str(int(time.time()))
    package = f"prepay_id={prepay_id}"

    sign_str = (
        f"appId={settings.WECHAT_APP_ID}"
        f"&nonceStr={nonce_str}"
        f"&package={package}"
        f"&signType=MD5"
        f"&timeStamp={timestamp}"
        f"&key={settings.WECHAT_API_KEY}"
    )
    pay_sign = hashlib.md5(sign_str.encode()).hexdigest().upper()

    return {
        "timeStamp": timestamp,
        "nonceStr": nonce_str,
        "package": package,
        "signType": "MD5",
        "paySign": pay_sign,
    }
