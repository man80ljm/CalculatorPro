"""教师账号、argon2 密码和可作废的签名会话 Cookie。"""
from __future__ import annotations

import os
import secrets

from argon2 import PasswordHasher
from argon2.exceptions import VerificationError
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from sqlalchemy import select
from sqlalchemy.exc import IntegrityError

from web_app.db import User, UserSession, SessionHead, lock_user, session_expiry, session_scope, utcnow

COOKIE_NAME = "cp_session"
SESSION_MAX_AGE = 12 * 60 * 60
MIN_PASSWORD_LENGTH = 8
MAX_PASSWORD_LENGTH = 128

_hasher = PasswordHasher(time_cost=2, memory_cost=19456, parallelism=1)
_dummy_hash: str | None = None


class AuthError(ValueError):
    def __init__(self, message: str, status: int = 400, code: str = ""):
        super().__init__(message)
        self.status = status
        self.code = code


def get_secret_key() -> str:
    key = os.environ.get("SECRET_KEY", "").strip()
    if not key:
        raise RuntimeError("SECRET_KEY is required")
    return key


def cookie_secure() -> bool:
    return os.environ.get("COOKIE_SECURE", "false").strip().lower() in {"1", "true", "yes", "on"}


def _serializer() -> URLSafeTimedSerializer:
    return URLSafeTimedSerializer(get_secret_key(), salt="cp-session-v2")


def normalize_username(value: str, *, strict: bool) -> str:
    name = (value or "").strip().casefold()
    if any(ord(ch) < 32 or ch.isspace() for ch in name):
        if strict:
            raise AuthError("用户名或邮箱不能包含空白或控制字符")
        return ""
    if strict and (len(name) < 3 or len(name) > 128):
        raise AuthError("用户名或邮箱长度需要在 3 到 128 之间")
    if len(name) > 128:
        return ""
    return name


def validate_new_password(password: str) -> None:
    if not isinstance(password, str) or len(password) < MIN_PASSWORD_LENGTH or len(password) > MAX_PASSWORD_LENGTH:
        raise AuthError(f"密码至少需要 {MIN_PASSWORD_LENGTH} 位，且不超过 {MAX_PASSWORD_LENGTH} 位")


def hash_password(password: str) -> str:
    return _hasher.hash(password)


def verify_password(password_hash: str, password: str) -> bool:
    try:
        return bool(_hasher.verify(password_hash, password))
    except VerificationError:
        return False


def _dummy() -> str:
    global _dummy_hash
    if _dummy_hash is None:
        _dummy_hash = hash_password("dummy-password-for-timing")
    return _dummy_hash


def register_user(username: str, password: str) -> tuple[int, str]:
    name = normalize_username(username, strict=True)
    validate_new_password(password)
    user = User(username=name, password_hash=hash_password(password))
    try:
        with session_scope() as db:
            db.add(user)
            db.flush()
            return user.id, user.username
    except IntegrityError as exc:
        raise AuthError("用户名已存在", status=409) from exc


def authenticate(username: str, password: str) -> tuple[int, str] | None:
    name = normalize_username(username, strict=False)
    if not name or not isinstance(password, str) or len(password) > MAX_PASSWORD_LENGTH:
        verify_password(_dummy(), password if isinstance(password, str) and len(password) <= MAX_PASSWORD_LENGTH else "wrong-password")
        return None
    with session_scope() as db:
        user = db.scalar(select(User).where(User.username == name))
        if user is None:
            verify_password(_dummy(), password)
            return None
        if not verify_password(user.password_hash, password):
            return None
        if _hasher.check_needs_rehash(user.password_hash):
            user.password_hash = hash_password(password)
        return user.id, user.username


def change_password(user_id: int, current_password: str, new_password: str) -> None:
    validate_new_password(new_password)
    with session_scope() as db:
        user = db.get(User, int(user_id))
        if user is None or not verify_password(user.password_hash, current_password or ""):
            raise AuthError("当前密码不正确", status=400)
        user.password_hash = hash_password(new_password)


def start_session(user_id: int) -> str:
    session_id = secrets.token_urlsafe(32)
    with session_scope() as db:
        lock_user(db, user_id)
        head = db.get(SessionHead, int(user_id))
        if head is None:
            db.add(SessionHead(user_id=int(user_id), session_id=session_id))
        else:
            head.session_id = session_id
        for old in db.scalars(select(UserSession).where(UserSession.user_id == int(user_id))).all():
            db.delete(old)
        db.add(
            UserSession(
                id=session_id,
                user_id=int(user_id),
                expires_at=session_expiry(),
            )
        )
    return _serializer().dumps({"uid": int(user_id), "sid": session_id})


def read_session_token(token: str | None) -> tuple[int, str] | None:
    if not token:
        return None
    try:
        get_secret_key()
    except RuntimeError:
        return None
    try:
        data = _serializer().loads(token, max_age=SESSION_MAX_AGE)
    except (BadSignature, SignatureExpired):
        return None
    if not isinstance(data, dict):
        return None
    user_id = data.get("uid")
    session_id = data.get("sid")
    if not isinstance(user_id, int) or not isinstance(session_id, str) or not session_id:
        return None
    return user_id, session_id


def session_is_active(user_id: int, session_id: str) -> bool:
    with session_scope() as db:
        head = db.get(SessionHead, user_id)
        if head is not None and head.session_id != session_id:
            return False
        row = db.get(UserSession, session_id)
        if row is None or row.user_id != int(user_id):
            return False
        if not row.alive(utcnow()):
            db.delete(row)
            return False
        return True


def session_was_replaced(user_id: int, session_id: str) -> bool:
    with session_scope() as db:
        head = db.get(SessionHead, user_id)
        return bool(head and head.session_id != session_id)


def revoke_session(session_id: str) -> None:
    if not session_id:
        return
    with session_scope() as db:
        row = db.get(UserSession, session_id)
        if row is not None:
            db.delete(row)


def revoke_other_sessions(user_id: int, keep_session_id: str) -> None:
    """改密码后让该用户其他设备上的会话失效。"""
    with session_scope() as db:
        rows = db.scalars(select(UserSession).where(UserSession.user_id == int(user_id))).all()
        for row in rows:
            if row.id != keep_session_id:
                db.delete(row)
