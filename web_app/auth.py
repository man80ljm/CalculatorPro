"""全站共享密码 + 签名 Cookie。密码只来自环境变量 APP_PASSWORD。"""
import hashlib
import os

from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

COOKIE_NAME = "cp_session"
SESSION_MAX_AGE = 12 * 60 * 60


def app_password() -> str:
    return os.environ.get("APP_PASSWORD", "").strip()


def _serializer() -> URLSafeTimedSerializer:
    secret = hashlib.sha256(
        f"calculatorpro-session:{app_password()}".encode("utf-8")
    ).hexdigest()
    return URLSafeTimedSerializer(secret, salt="cp-session")


def make_session_token() -> str:
    return _serializer().dumps({"ok": 1})


def read_session_token(token: str | None) -> bool:
    if not token or not app_password():
        return False
    try:
        data = _serializer().loads(token, max_age=SESSION_MAX_AGE)
    except (BadSignature, SignatureExpired):
        return False
    return bool(isinstance(data, dict) and data.get("ok"))
