"""SQLAlchemy 模型。PostgreSQL 与测试用的 SQLite 共用这一套表。"""
from __future__ import annotations

import os
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone

from sqlalchemy import DateTime, ForeignKey, Integer, String, Text, create_engine, event, text
from sqlalchemy.engine import Engine
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column
from sqlalchemy.pool import StaticPool

_engine: Engine | None = None
_engine_url: str | None = None


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
        engine = create_engine(
            url,
            connect_args={"check_same_thread": False, "timeout": 30},
            poolclass=StaticPool,
        )

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


def init_db() -> None:
    get_engine()
    Base.metadata.create_all(get_engine())


@contextmanager
def session_scope():
    db = Session(get_engine(), expire_on_commit=False)
    try:
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


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


class CourseFile(Base):
    __tablename__ = "course_files"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id"), index=True)
    course_id: Mapped[int] = mapped_column(ForeignKey("courses.id"), index=True)
    kind: Mapped[str] = mapped_column(String(32), index=True)
    stored_name: Mapped[str] = mapped_column(String(80))
    original_name: Mapped[str] = mapped_column(String(255))
    size: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
