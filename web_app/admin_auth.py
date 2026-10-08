"""运维后台登录。与教师 users 表、cp_session 完全分开。"""
from __future__ import annotations

import hashlib
import os
import secrets

from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

from web_app.auth import SESSION_MAX_AGE, cookie_secure, get_secret_key, verify_password

ADMIN_COOKIE = "cp_admin_session"
ADMIN_SALT = "cp-admin-session-v1"
MAX_ADMIN_PASSWORD_LENGTH = 512
_NOT_CONFIGURED = "未配置管理员密码"


class AdminAuthError(ValueError):
    def __init__(self, message: str, status: int):
        super().__init__(message)
        self.status = status


def is_admin_path(path: str) -> bool:
    return path == "/admin" or path.startswith("/admin/")


def admin_username() -> str:
    raw = os.environ.get("ADMIN_USERNAME", "").strip()
    return (raw or "admin").casefold()


def _password_hash() -> str:
    return os.environ.get("ADMIN_PASSWORD_HASH", "").strip()


def _password_plain() -> str:
    return os.environ.get("ADMIN_PASSWORD", "").strip()


def admin_configured() -> bool:
    return bool(_password_hash() or _password_plain())


def admin_host() -> str:
    return normalize_host(os.environ.get("ADMIN_HOST", ""))


def normalize_host(value: str) -> str:
    host = (value or "").split(",")[0].strip().lower()
    if not host:
        return ""
    if host.startswith("["):
        end = host.find("]")
        if end != -1:
            return host[1:end]
    if host.count(":") == 1:
        host = host.split(":", 1)[0]
    return host.rstrip(".")


def _config_tag() -> str:
    material = _password_hash() or _password_plain()
    payload = f"{admin_username()}\n{material}".encode()
    return hashlib.sha256(payload).hexdigest()[:24]


def _serializer() -> URLSafeTimedSerializer:
    return URLSafeTimedSerializer(get_secret_key(), salt=ADMIN_SALT)


def _same(left: str, right: str) -> bool:
    """按字节比较。hmac 的字符串比较不接受非 ASCII，中文口令不能因此变成 500。"""
    return secrets.compare_digest(left.encode("utf-8"), right.encode("utf-8"))


def _hash_matches(password_hash: str, password: str) -> bool:
    try:
        return verify_password(password_hash, password)
    except (ValueError, TypeError):
        return False


def authenticate_admin(username: str, password: str) -> bool:
    """只核对环境变量里的管理员口令，不查教师表。"""
    if not admin_configured():
        raise AdminAuthError(_NOT_CONFIGURED, 503)
    if not isinstance(username, str) or not isinstance(password, str):
        return False
    if len(username) > 128 or len(password) > MAX_ADMIN_PASSWORD_LENGTH:
        return False
    supplied_user = username.strip().casefold()
    expected_user = admin_username()
    password_hash = _password_hash()
    if password_hash:
        password_ok = _hash_matches(password_hash, password)
    else:
        password_ok = _same(password, _password_plain())
    user_ok = _same(supplied_user, expected_user)
    return bool(password_ok and user_ok)


def start_admin_session() -> str:
    if not admin_configured():
        raise AdminAuthError(_NOT_CONFIGURED, 503)
    return _serializer().dumps({"u": admin_username(), "tag": _config_tag()})


def read_admin_session(token: str | None) -> str | None:
    if not token or not admin_configured():
        return None
    try:
        get_secret_key()
        data = _serializer().loads(token, max_age=SESSION_MAX_AGE)
    except (BadSignature, SignatureExpired, RuntimeError):
        return None
    if not isinstance(data, dict):
        return None
    username = data.get("u")
    tag = data.get("tag")
    if not isinstance(username, str) or not isinstance(tag, str):
        return None
    if not _same(tag, _config_tag()):
        return None
    if not _same(username, admin_username()):
        return None
    return username


def set_admin_cookie(response, token: str) -> None:
    response.set_cookie(
        ADMIN_COOKIE,
        token,
        max_age=SESSION_MAX_AGE,
        httponly=True,
        samesite="lax",
        secure=cookie_secure(),
        path="/",
    )


def clear_admin_cookie(response) -> None:
    response.delete_cookie(
        ADMIN_COOKIE,
        path="/",
        httponly=True,
        samesite="lax",
        secure=cookie_secure(),
    )
