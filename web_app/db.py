"""SQLAlchemy 模型。PostgreSQL 与测试用的 SQLite 共用这一套表。"""
from __future__ import annotations

import os
from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime, timedelta, timezone

from sqlalchemy import CheckConstraint, UniqueConstraint, DateTime, ForeignKey, Integer, String, Text, create_engine, event, text
from sqlalchemy.engine import Engine, make_url
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column
from sqlalchemy.pool import StaticPool
from sqlalchemy.exc import IntegrityError

_engine: Engine | None = None
_engine_url: str | None = None
write_context: ContextVar[dict | None] = ContextVar("calculatorpro_write_context", default=None)
job_engine: ContextVar[Engine | None] = ContextVar("calculatorpro_job_engine", default=None)


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
    fixed = job_engine.get()
    if fixed is not None:
        return fixed
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
    with Session(engine) as db:
        if db.get(JobQueueLock, 1) is None:
            db.add(JobQueueLock(id=1))
            try:
                db.commit()
            except IntegrityError:
                db.rollback()
                if db.get(JobQueueLock, 1) is None:
                    raise


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
        for path in db.info.get("new_blob_paths", []):
            path.unlink(missing_ok=True)
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


class JobQueueLock(Base):
    __tablename__ = "job_queue_lock"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)


class ReportTask(Base):
    """输入放磁盘，状态放数据库；网页刷新和服务重启后仍可找回。"""
    __tablename__ = "report_tasks"
    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    course_id: Mapped[int] = mapped_column(ForeignKey("courses.id", ondelete="CASCADE"), index=True)
    term_id: Mapped[int] = mapped_column(ForeignKey("terms.id", ondelete="CASCADE"), index=True)
    active_key: Mapped[str | None] = mapped_column(String(120), unique=True, nullable=True)
    status: Mapped[str] = mapped_column(String(16), index=True, default="queued")
    settings_json: Mapped[str] = mapped_column(Text, default="{}")
    meta_json: Mapped[str] = mapped_column(Text, default="{}")
    public_json: Mapped[str] = mapped_column(Text, default="{}")
    claim_token: Mapped[str] = mapped_column(String(64), default="")
    lease_until: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    archive_file_id: Mapped[int | None] = mapped_column(ForeignKey("course_files.id", ondelete="SET NULL"), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, index=True)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)

class InvitationCode(Base):
    """注册邀请码；删除保留码值和历史，避免重新派发。"""
    __tablename__ = "invitation_codes"
    __table_args__ = (
        UniqueConstraint("code", name="uq_invitation_code"),
        UniqueConstraint("owner_user_id", name="uq_invitation_owner"),
        CheckConstraint("used_count >= 0", name="ck_invitation_used"),
        CheckConstraint("max_uses IS NULL OR max_uses >= used_count", name="ck_invitation_limit"),
    )
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    code: Mapped[str] = mapped_column(String(6))
    owner_user_id: Mapped[int | None] = mapped_column(ForeignKey("users.id"), nullable=True)
    max_uses: Mapped[int | None] = mapped_column(Integer, nullable=True)
    used_count: Mapped[int] = mapped_column(Integer, default=0)
    enabled: Mapped[int] = mapped_column(Integer, default=1)
    deleted_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class InvitationRegistration(Base):
    """注册准入与邀请来源，和账号/扣次在同一事务保存。"""
    __tablename__ = "invitation_registrations"
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), primary_key=True)
    invitation_id: Mapped[int] = mapped_column(ForeignKey("invitation_codes.id"), index=True)
    invited_by: Mapped[str] = mapped_column(String(128))
    inviter_user_id: Mapped[int | None] = mapped_column(ForeignKey("users.id"), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class AdminApiConfiguration(Base):
    __tablename__ = "admin_api_configuration"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    encrypted_config: Mapped[str] = mapped_column(Text)
    last_test: Mapped[str | None] = mapped_column(Text, nullable=True)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)


class ScheduleDocument(Base):
    __tablename__ = "schedule_documents"
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    filename: Mapped[str] = mapped_column(String(200))
    digest: Mapped[str] = mapped_column(String(64), unique=True)
    term: Mapped[str] = mapped_column(Text)
    model: Mapped[str] = mapped_column(String(100))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class ScheduleDraft(Base):
    __tablename__ = "schedule_drafts"
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    filename: Mapped[str] = mapped_column(String(200))
    digest: Mapped[str] = mapped_column(String(64), index=True)
    model: Mapped[str] = mapped_column(String(100))
    status: Mapped[str] = mapped_column(String(20))
    result: Mapped[str | None] = mapped_column(Text, nullable=True)
    error: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class ScheduleRecord(Base):
    __tablename__ = "schedule_records"
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    document_id: Mapped[str] = mapped_column(ForeignKey("schedule_documents.id"), index=True)
    term_id: Mapped[str] = mapped_column(String(20), index=True)
    course_name: Mapped[str] = mapped_column(String(160), index=True)
    payload: Mapped[str] = mapped_column(Text)


class ScheduleDeletion(Base):
    """Recoverable timetable removals; active tables retain their existing schema."""
    __tablename__ = "schedule_deletions"
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    scope: Mapped[str] = mapped_column(String(20))
    label: Mapped[str] = mapped_column(String(255))
    snapshot_json: Mapped[str] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    restored_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True, index=True)

class ScheduleRecognitionChunk(Base):
    __tablename__ = "schedule_recognition_chunks"
    cache_key: Mapped[str] = mapped_column(String(64), primary_key=True)
    payload: Mapped[str] = mapped_column(Text)
    usage: Mapped[str] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class LessonPlanSource(Base):
    __tablename__ = "lesson_plan_sources"
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    course_id: Mapped[int | None] = mapped_column(ForeignKey("courses.id", ondelete="CASCADE"), nullable=True, index=True)
    filename: Mapped[str] = mapped_column(String(255))
    stored_path: Mapped[str] = mapped_column(String(255))
    digest: Mapped[str] = mapped_column(String(64), index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)


class LessonPlan(Base):
    __tablename__ = "lesson_plans"
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    course_id: Mapped[int | None] = mapped_column(ForeignKey("courses.id", ondelete="CASCADE"), nullable=True, index=True)
    source_id: Mapped[str] = mapped_column(ForeignKey("lesson_plan_sources.id", ondelete="CASCADE"))
    context_json: Mapped[str] = mapped_column(Text, default="{}")
    revision: Mapped[int] = mapped_column(Integer, default=1)
    warnings_json: Mapped[str] = mapped_column(Text, default="[]")
    pages_json: Mapped[str] = mapped_column(Text, default="{}")
    checked_revision: Mapped[int | None] = mapped_column(Integer, nullable=True)
    output_path: Mapped[str] = mapped_column(String(255), default="")
    output_revision: Mapped[int | None] = mapped_column(Integer, nullable=True)
    output_date: Mapped[str] = mapped_column(String(20), default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)


class LessonPlanTask(Base):
    __tablename__ = "lesson_plan_tasks"
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    plan_id: Mapped[str] = mapped_column(ForeignKey("lesson_plans.id", ondelete="CASCADE"), index=True)
    action: Mapped[str] = mapped_column(String(20))
    lesson: Mapped[int | None] = mapped_column(Integer, nullable=True)
    active_key: Mapped[str | None] = mapped_column(String(64), unique=True, nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="queued", index=True)
    input_revision: Mapped[int] = mapped_column(Integer)
    checkpoint_json: Mapped[str] = mapped_column(Text, default="{}")
    public_json: Mapped[str] = mapped_column(Text, default="{}")
    claim_token: Mapped[str] = mapped_column(String(32), default="")
    lease_until: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)


class TeachingAssignmentSelection(Base):
    """Calendar/lesson-plan teacher selection; never changes grade-register metadata."""
    __tablename__ = "teaching_assignment_selections"
    __table_args__ = (UniqueConstraint("user_id", "course_key", "term_id"),)
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    course_key: Mapped[str] = mapped_column(String(64))
    term_id: Mapped[str] = mapped_column(String(20))
    teacher: Mapped[str] = mapped_column(String(160))
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)


class TeachingMetadataOverride(Base):
    """User-confirmed teaching metadata, isolated from timetable imports and grades."""
    __tablename__ = "teaching_metadata_overrides"
    __table_args__ = (UniqueConstraint("user_id", "scope_key", "term_id"),)
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    course_id: Mapped[int | None] = mapped_column(ForeignKey("courses.id", ondelete="CASCADE"), nullable=True, index=True)
    plan_id: Mapped[str | None] = mapped_column(ForeignKey("lesson_plans.id", ondelete="CASCADE"), nullable=True)
    scope_key: Mapped[str] = mapped_column(String(64))
    term_id: Mapped[str] = mapped_column(String(20))
    fields_json: Mapped[str] = mapped_column(Text, default="{}")
    revision: Mapped[int] = mapped_column(Integer, default=1)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)


class SyllabusImportTask(Base):
    """Durable syllabus repair jobs; existing report tasks are unaffected."""
    __tablename__ = "syllabus_import_tasks"
    expected_version: Mapped[int | None] = mapped_column(Integer, nullable=True)
    source_id: Mapped[str | None] = mapped_column(ForeignKey("lesson_plan_sources.id", ondelete="SET NULL"), nullable=True)
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    course_id: Mapped[int | None] = mapped_column(ForeignKey("courses.id", ondelete="CASCADE"), nullable=True, index=True)
    filename: Mapped[str] = mapped_column(String(255))
    digest: Mapped[str] = mapped_column(String(64))
    input_path: Mapped[str] = mapped_column(String(255), default="")
    allow_completion: Mapped[int] = mapped_column(Integer, default=1)
    active_key: Mapped[str | None] = mapped_column(String(64), unique=True, nullable=True)
    status: Mapped[str] = mapped_column(String(16), default="queued", index=True)
    result_json: Mapped[str] = mapped_column(Text, default="{}")
    public_json: Mapped[str] = mapped_column(Text, default="{}")
    claim_token: Mapped[str] = mapped_column(String(32), default="")
    lease_until: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)


class TeacherImport(Base):
    __tablename__ = "teacher_imports"
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    filename: Mapped[str] = mapped_column(String(200))
    digest: Mapped[str] = mapped_column(String(64), unique=True)
    source_date: Mapped[str] = mapped_column(String(20), default="")
    payload: Mapped[str] = mapped_column(Text)
    status: Mapped[str] = mapped_column(String(20), default="preview")
    revision: Mapped[int] = mapped_column(Integer, default=1)
    base_revision: Mapped[str] = mapped_column(String(64))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    saved_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)


class TeacherProfile(Base):
    """Staff directory metadata, separate from teacher login accounts and grades."""
    __tablename__ = "teacher_profiles"
    __table_args__ = (CheckConstraint("active IN (0, 1)", name="ck_teacher_profile_active"),)
    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    name: Mapped[str] = mapped_column(String(160))
    name_key: Mapped[str] = mapped_column(String(160), index=True)
    college: Mapped[str] = mapped_column(String(160), default="", index=True)
    department: Mapped[str] = mapped_column(String(160), default="", index=True)
    teaching_unit: Mapped[str] = mapped_column(String(160), default="", index=True)
    group_name: Mapped[str] = mapped_column(String(160), default="")
    mobile: Mapped[str] = mapped_column(String(60), default="")
    office_phone: Mapped[str] = mapped_column(String(100), default="")
    notes: Mapped[str] = mapped_column(Text, default="")
    active: Mapped[int] = mapped_column(Integer, default=1, index=True)
    revision: Mapped[int] = mapped_column(Integer, default=1)
    source_import_id: Mapped[str | None] = mapped_column(ForeignKey("teacher_imports.id"), nullable=True)
    source_page: Mapped[int | None] = mapped_column(Integer, nullable=True)
    source_group: Mapped[str] = mapped_column(String(160), default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=utcnow, onupdate=utcnow)
