"""正式邀请码：服务端生成、数据库唯一约束、事务内扣次。"""
from __future__ import annotations

import os
import re
import secrets

from sqlalchemy import select, text
from sqlalchemy.exc import IntegrityError

from web_app.db import InvitationCode, User, lock_user, session_scope, utcnow

ALPHABET = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
MAX_REGISTRATIONS = 1_000_000


class InvitationError(ValueError):
    def __init__(self, message: str, status: int = 400):
        super().__init__(message)
        self.status = status


def validate_limit(value, *, used_count: int = 0) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value <= MAX_REGISTRATIONS:
        raise InvitationError(f"可注册总次数应为 0 到 {MAX_REGISTRATIONS} 的整数，或选择不限次数")
    if value < used_count:
        raise InvitationError(f"可注册总次数不能小于已使用次数（{used_count} 次）")
    return value


def default_user_limit() -> int | None:
    value = os.environ.get("INVITE_USER_DEFAULT_USES", "5").strip().lower()
    if value == "unlimited":
        return None
    try:
        return validate_limit(int(value))
    except (ValueError, InvitationError) as exc:
        raise RuntimeError("INVITE_USER_DEFAULT_USES 应为 0 到 1000000 的整数，或 unlimited") from exc


def normalize_code(value) -> str:
    if not isinstance(value, str) or not value.strip():
        raise InvitationError("请填写有效邀请码后注册；若页面没有邀请码输入框，请刷新页面")
    code = value.strip().upper()
    if not re.fullmatch(r"[A-Z0-9]{6}", code):
        raise InvitationError("邀请码应为 6 位字母与数字")
    return code


def _random_code() -> str:
    for _ in range(100):
        value = "".join(secrets.choice(ALPHABET) for _ in range(6))
        if re.search(r"[A-Z]", value) and re.search(r"[0-9]", value):
            return value
    raise InvitationError("邀请码生成失败，请稍后重试", status=503)


def _sqlite_write_lock(db) -> None:
    if db.bind.dialect.name == "sqlite" and not db.connection().connection.driver_connection.in_transaction:
        db.execute(text("BEGIN IMMEDIATE"))


def new_invitation(db, *, max_uses: int | None, owner_user_id: int | None = None) -> InvitationCode:
    """保存点处理并发随机碰撞，不回滚外层的注册事务。"""
    validate_limit(max_uses)
    _sqlite_write_lock(db)
    for _ in range(100):
        code = _random_code()
        if db.scalar(select(InvitationCode.id).where(InvitationCode.code == code)) is not None:
            continue
        row = InvitationCode(code=code, owner_user_id=owner_user_id, max_uses=max_uses, used_count=0, enabled=1)
        try:
            with db.begin_nested():
                db.add(row)
                db.flush()
            return row
        except IntegrityError:
            if db.scalar(select(InvitationCode.id).where(InvitationCode.code == code)) is None:
                raise
    raise InvitationError("邀请码生成失败，请稍后重试", status=503)


def _locked(db, *, code: str | None = None, invitation_id: int | None = None) -> InvitationCode | None:
    # SQLite 不支持行锁，在任何读取前取得写锁；PostgreSQL 使用行锁。
    _sqlite_write_lock(db)
    query = select(InvitationCode)
    query = query.where(InvitationCode.code == code) if code is not None else query.where(InvitationCode.id == invitation_id)
    return db.scalar(query.with_for_update())


def claim_invitation(db, code: str) -> InvitationCode:
    row = _locked(db, code=code)
    if row is None or row.deleted_at or not row.enabled or (row.max_uses is not None and row.used_count >= row.max_uses):
        raise InvitationError("邀请码无效、已停用、已删除或注册次数已用完")
    row.used_count += 1
    return row


def invitation_status(row: InvitationCode) -> str:
    if row.deleted_at:
        return "deleted"
    if not row.enabled:
        return "disabled"
    if row.max_uses is not None and row.used_count >= row.max_uses:
        return "exhausted"
    return "active"


def public_invitation(row: InvitationCode, owner: str = "") -> dict:
    return {
        "id": row.id,
        "code": row.code,
        "owner_user_id": row.owner_user_id,
        "owner": owner if row.owner_user_id is not None else "管理员",
        "max_uses": row.max_uses,
        "used_count": row.used_count,
        "remaining": None if row.max_uses is None else max(0, row.max_uses - row.used_count),
        "enabled": bool(row.enabled),
        "status": invitation_status(row),
        "created_at": row.created_at.isoformat(timespec="seconds"),
    }


def create_admin_invitation(max_uses: int | None = 5) -> dict:
    validate_limit(max_uses)
    with session_scope() as db:
        return public_invitation(new_invitation(db, max_uses=max_uses))


def list_invitations() -> list[dict]:
    with session_scope() as db:
        rows = db.execute(
            select(InvitationCode, User.username)
            .outerjoin(User, User.id == InvitationCode.owner_user_id)
            .where(InvitationCode.deleted_at.is_(None))
            .order_by(InvitationCode.created_at.desc(), InvitationCode.id.desc())
        ).all()
        return [public_invitation(row, owner or "") for row, owner in rows]


def update_invitation(invitation_id: int, changes: dict) -> dict:
    if not changes or set(changes) - {"max_uses", "enabled"}:
        raise InvitationError("只可修改可注册总次数和启用状态")
    if "enabled" in changes and not isinstance(changes["enabled"], bool):
        raise InvitationError("启用状态格式不正确")
    with session_scope() as db:
        row = _locked(db, invitation_id=invitation_id)
        if row is None or row.deleted_at:
            raise InvitationError("邀请码不存在或已删除", status=404)
        if "max_uses" in changes:
            row.max_uses = validate_limit(changes["max_uses"], used_count=row.used_count)
        if "enabled" in changes:
            row.enabled = int(changes["enabled"])
        owner = db.get(User, row.owner_user_id) if row.owner_user_id is not None else None
        return public_invitation(row, owner.username if owner else "")


def delete_invitation(invitation_id: int) -> None:
    with session_scope() as db:
        row = _locked(db, invitation_id=invitation_id)
        if row is None or row.deleted_at:
            raise InvitationError("邀请码不存在或已删除", status=404)
        row.deleted_at = utcnow()
        row.enabled = 0


def personal_invitation(user_id: int) -> dict:
    """旧账号首次查看时补发专属码；已删除的个人码不重新生成。"""
    with session_scope() as db:
        user = lock_user(db, user_id)
        if user is None:
            raise InvitationError("用户不存在", status=404)
        row = db.scalar(select(InvitationCode).where(InvitationCode.owner_user_id == user_id))
        if row is None:
            row = new_invitation(db, max_uses=default_user_limit(), owner_user_id=user_id)
        return public_invitation(row, user.username)
