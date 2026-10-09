"""SQLAlchemy 模型。PostgreSQL 与测试用的 SQLite 共用这一套表。"""
from __future__ import annotations

import os
from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime, timedelta, timezone

from sqlalchemy import UniqueConstraint, DateTime, ForeignKey, Integer, String, Text, create_engine, event, text
from sqlalchemy.engine import Engine, make_url
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column
from sqlalchemy.pool import StaticPool

_engine: Engine | None = None
_engine_url: str | None = None
write_context: ContextVar[dict | None] = ContextVar("calculatorpro_write_context", default=None)


class NotFound(LookupError):
    """课程或文件不属于当前用户，或根本不存在。对外一律 404。"""


class Base(DeclarativeBase):
    pass


def utcnow() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def sqlite_allowed() -> bool:
    return os.environ.get("ALLOW_SQLITE", "").strip().lower() in {"1", "true", "yes", "on"}


def database_url() -> str:
    """部署必须显式配置 DATABASE_URL；SQLite 只在 ALLOW_SQLITE=1 时允许（测试、本地试用）。"""
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        if sqlite_allowed():
            return "sqlite:///./calculatorpro.db"
        raise RuntimeError(
            "DATABASE_URL is required (e.g. postgresql+psycopg://user:pass@postgres:5432/calculatorpro). "
            "Set ALLOW_SQLITE=1 to explicitly opt in to a local SQLite database for tests or trials."
        )
    if url.startswith("sqlite") and not sqlite_allowed():
        raise RuntimeError("DATABASE_URL points to SQLite; set ALLOW_SQLITE=1 to explicitly allow it.")
    return url


def get_engine() -> Engine:
    global _engine, _engine_url
    url = database_url()
    if _engine is not None and _engine_url == url:
        return _engine
    if _engine is not None:
        _engine.dispose()
    if url.startswith("sqlite"):
        parsed = make_url(url)
        options = {"connect_args": {"check_same_thread": False, "timeout": 30}}
        # 内存库需要共享连接才能保留数据；文件库让每个并行事务独立借用连接。
        # StaticPool 用在文件库会让报告线程互相提交/回滚，甚至丢失文件记录。
        if parsed.database in (None, "", ":memory:") or parsed.query.get("mode") == "memory":
            options["poolclass"] = StaticPool
        engine = create_engine(url, **options)

        @event.listens_for(engine, "connect")
        def _enable_sqlite_fk(connection, _record):  # noqa: ANN001
            cursor = connection.cursor()
            cursor.execute("PRAGMA foreign_keys=ON")
            cursor.close()
    else:
        engine = create_engine(url, pool_pre_ping=True)
    _engine = engine
    _engine_url = url
    return engine


MIGRATION_HINT = (
    "数据库还没有 v2 学期结构（缺少 terms 表或 course_files.term_id）。"
    "请先备份，再运行 python -m web_app.migrate_v2 --backup-dir <目录>。"
    "启动时不会自动改数据。"
)


def needs_v2_migration(engine: Engine | None = None) -> bool:
    """已有旧库但还没有学期表时返回 True。空库返回 False，交给 create_all。"""
    engine = engine or get_engine()
    from sqlalchemy import inspect

    insp = inspect(engine)
    tables = set(insp.get_table_names())
    if "courses" not in tables:
        return False
    if "terms" not in tables:
        return True
    if "course_files" not in tables:
        return False
    columns = {column["name"] for column in insp.get_columns("course_files")}
    return "term_id" not in columns


def init_db() -> None:
    """创建缺失的表。已有表不改结构；新表（如 report_job_events）在启动时补上。"""
    engine = get_engine()
    if needs_v2_migration(engine):
        raise RuntimeError(MIGRATION_HINT)
    Base.metadata.create_all(engine)


@contextmanager
def session_scope():
    db = Session(get_engine(), expire_on_commit=False)
    try:
        guard = write_context.get()
        if guard:
            # 登录切换与写入使用同一把用户锁，已进入处理的旧请求也不能越过切换。
            lock_user(db, guard["user_id"])
            head = db.get(SessionHead, guard["user_id"])
            session = db.get(UserSession, guard["session_id"])
            if session is None or not session.alive() or (head and head.session_id != guard["session_id"]):
                from web_app.auth import AuthError
                raise AuthError("账号已在别处登录，本页已停止保存。填写内容已保留。", status=401, code="session_replaced")
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


def lock_user(db, user_id: int):
    from sqlalchemy import select
    if db.bind.dialect.name == "sqlite":
        db.execute(text("BEGIN IMMEDIATE"))
    return db.scalar(select(User).where(User.id == int(user_id)).with_for_update())


def ping_db() -> None:
    with session_scope() as db:
        db.execute(text("SELECT 1"))


class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    username: Mapped[str] = mapped_column(String(128), unique=True, index=True)
    password_hash: Mapped[str] = mapped_column(String(512))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class UserSession(Base):
    __tablename__ = "user_sessions"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    expires_at: Mapped[datetime] = mapped_column(DateTime, index=True)

    def alive(self, now: datetime | None = None) -> bool:
        return self.expires_at > (now or utcnow())


class SessionHead(Base):
    """每账号当前有效会话。单独建表，兼容已有数据库。"""
    __tablename__ = "session_heads"
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), primary_key=True)
    session_id: Mapped[str] = mapped_column(String(64))


def session_expiry(now: datetime | None = None) -> datetime:
    return (now or utcnow()) + timedelta(hours=12)


class Course(Base):
    __tablename__ = "courses"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    name: Mapped[str] = mapped_column(String(120))
    settings_json: Mapped[str] = mapped_column(Text, default="{}")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)


class CourseVersion(Base):
    """教学大纲版本。旧学期另存当时实际使用的设置，不改现有表结构。"""
    __tablename__ = "course_versions"
    __table_args__ = (UniqueConstraint("course_id", "number"),)
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    course_id: Mapped[int] = mapped_column(ForeignKey("courses.id", ondelete="CASCADE"), index=True)
    number: Mapped[int] = mapped_column(Integer)
    settings_json: Mapped[str] = mapped_column(Text, default="{}")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class Term(Base):
    """一门课下的一条学期记录。关系表仍在课程上，各学期共用。"""

    __tablename__ = "terms"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    course_id: Mapped[int] = mapped_column(ForeignKey("courses.id"), index=True)
    label: Mapped[str] = mapped_column(String(160), default="")
    year_start: Mapped[str] = mapped_column(String(16), default="")
    year_end: Mapped[str] = mapped_column(String(16), default="")
    semester: Mapped[str] = mapped_column(String(8), default="")
    school_year_term: Mapped[str] = mapped_column(String(80), default="")
    teacher: Mapped[str] = mapped_column(String(80), default="")
    class_name: Mapped[str] = mapped_column(String(120), default="")
    student_count: Mapped[int] = mapped_column(Integer, default=0)
    exam_count: Mapped[int] = mapped_column(Integer, default=0)
    settings_json: Mapped[str] = mapped_column(Text, default="{}")
    is_current: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)


class CourseFile(Base):
    __tablename__ = "course_files"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    course_id: Mapped[int] = mapped_column(ForeignKey("courses.id"), index=True)
    term_id: Mapped[int | None] = mapped_column(ForeignKey("terms.id"), nullable=True, index=True)
    kind: Mapped[str] = mapped_column(String(32), index=True)
    stored_name: Mapped[str] = mapped_column(String(80))
    original_name: Mapped[str] = mapped_column(String(255))
    size: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class FileBatch(Base):
    """一次成功生成的资料；旧文件无需迁移，仍可通过原压缩包下载。"""

    __tablename__ = "file_batches"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    course_id: Mapped[int] = mapped_column(ForeignKey("courses.id", ondelete="CASCADE"), index=True)
    term_id: Mapped[int] = mapped_column(ForeignKey("terms.id", ondelete="CASCADE"), index=True)
    archive_file_id: Mapped[int] = mapped_column(ForeignKey("course_files.id", ondelete="CASCADE"), unique=True)
    kind: Mapped[str] = mapped_column(String(32))
    file_ids_json: Mapped[str] = mapped_column(Text, default="[]")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class ReportJobEvent(Base):
    """报告任务结束时追加的只读统计。不挂外键，课程删了也留着。"""

    __tablename__ = "report_job_events"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    course_id: Mapped[int] = mapped_column(Integer, index=True)
    term_id: Mapped[int | None] = mapped_column(Integer, nullable=True)
    status: Mapped[str] = mapped_column(String(16), index=True)
    error: Mapped[str] = mapped_column(Text, default="")
    duration_ms: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, index=True)
