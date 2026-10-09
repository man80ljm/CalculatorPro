"""运维后台只读统计。写入只发生在报告任务结束时，不在这里。"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone

from sqlalchemy import func, select

from web_app.db import Course, CourseFile, ReportJobEvent, Term, User, UserSession, session_scope
from web_app.deepseek_pool import queue_depth
from web_app.report_jobs import list_jobs_for_admin

_SHANGHAI = timezone(timedelta(hours=8))
_EVENT_LIMIT = 100
_ERROR_LIMIT = 180


def _clip(text: str, limit: int = _ERROR_LIMIT) -> str:
    cleaned = " ".join((text or "").split())
    if len(cleaned) <= limit:
        return cleaned
    return cleaned[: limit - 1] + "…"


def _iso(value: datetime | None) -> str | None:
    if value is None:
        return None
    return value.isoformat(timespec="seconds")


def _bounds(now: datetime | None = None) -> tuple[datetime, datetime]:
    """返回（今日 0 点、本周一 0 点），都是去掉时区的 UTC，方便和库里的时间比较。"""
    current = now or datetime.now(timezone.utc)
    if current.tzinfo is None:
        current = current.replace(tzinfo=timezone.utc)
    local = current.astimezone(_SHANGHAI)
    today_local = local.replace(hour=0, minute=0, second=0, microsecond=0)
    week_local = today_local - timedelta(days=today_local.weekday())

    def as_utc(value: datetime) -> datetime:
        return value.astimezone(timezone.utc).replace(tzinfo=None)

    return as_utc(today_local), as_utc(week_local)


def _event_counts(db, start: datetime) -> dict:
    rows = db.execute(
        select(ReportJobEvent.status, func.count(ReportJobEvent.id))
        .where(ReportJobEvent.created_at >= start, ReportJobEvent.status.in_(("success", "fail")))
        .group_by(ReportJobEvent.status)
    ).all()
    found = {status: int(count) for status, count in rows}
    return {"success": found.get("success", 0), "fail": found.get("fail", 0)}


def _active_teacher_ids(db, week_start: datetime) -> set[int]:
    ids: set[int] = set()
    queries = (
        select(UserSession.user_id).where(UserSession.created_at >= week_start),
        select(Course.user_id).where(Course.updated_at >= week_start),
        select(CourseFile.user_id).where(CourseFile.created_at >= week_start),
    )
    for query in queries:
        ids.update(int(item) for item in db.scalars(query).all())
    return ids


def _count_by(db, model, column) -> dict[int, int]:
    rows = db.execute(select(column, func.count(model.id)).group_by(column)).all()
    return {int(key): int(count) for key, count in rows}


def build_overview() -> dict:
    live = list_jobs_for_admin()
    report_queued = sum(1 for row in live if row["status"] == "queued")
    deepseek_waiting = int(queue_depth())
    today_start, week_start = _bounds()
    with session_scope() as db:
        teacher_count = int(db.scalar(select(func.count(User.id))) or 0)
        course_count = int(db.scalar(select(func.count(Course.id))) or 0)
        active = _active_teacher_ids(db, week_start)
        today = _event_counts(db, today_start)
        week = _event_counts(db, week_start)
    return {
        "teacher_count": teacher_count,
        "course_count": course_count,
        "active_teachers_week": len(active),
        "reports_today": today,
        "reports_week": week,
        "deepseek_queue": deepseek_waiting,
        "report_queued": report_queued,
        "queue_depth": int(queue_depth(exclude_kind="report")) + report_queued,
    }


def build_users() -> dict:
    with session_scope() as db:
        users = list(db.scalars(select(User).order_by(User.created_at.desc(), User.id.desc())).all())
        course_counts = _count_by(db, Course, Course.user_id)
        term_counts = _count_by(db, Term, Term.user_id)
        last_login = dict(
            db.execute(select(UserSession.user_id, func.max(UserSession.created_at)).group_by(UserSession.user_id)).all()
        )
        last_report = dict(
            db.execute(
                select(CourseFile.user_id, func.max(CourseFile.created_at))
                .where(CourseFile.kind == "report")
                .group_by(CourseFile.user_id)
            ).all()
        )
        rows = []
        for user in users:
            rows.append(
                {
                    "id": user.id,
                    "username": user.username,
                    "created_at": _iso(user.created_at),
                    "course_count": course_counts.get(user.id, 0),
                    "term_count": term_counts.get(user.id, 0),
                    "last_login_at": _iso(last_login.get(user.id)),
                    "last_report_at": _iso(last_report.get(user.id)),
                }
            )
    return {"users": rows}


def _names(db, user_ids: set[int], course_ids: set[int]) -> tuple[dict[int, str], dict[int, str]]:
    users: dict[int, str] = {}
    courses: dict[int, str] = {}
    if user_ids:
        for user in db.scalars(select(User).where(User.id.in_(user_ids))).all():
            users[user.id] = user.username
    if course_ids:
        for course in db.scalars(select(Course).where(Course.id.in_(course_ids))).all():
            courses[course.id] = course.name
    return users, courses


def _public_job(row: dict, users: dict[int, str], courses: dict[int, str]) -> dict:
    user_id = row.get("user_id")
    course_id = row.get("course_id")
    return {
        "source": row.get("source") or "",
        "job_id": row.get("job_id"),
        "event_id": row.get("event_id"),
        "user_id": user_id,
        "username": users.get(int(user_id), "") if user_id is not None else "",
        "course_id": course_id,
        "course_name": courses.get(int(course_id), "") if course_id is not None else "",
        "term_id": row.get("term_id"),
        "status": row.get("status") or "",
        "duration_ms": int(row.get("duration_ms") or 0),
        "error": _clip(str(row.get("error") or "")),
        "created_at": row.get("created_at"),
    }


def build_jobs() -> dict:
    live = list_jobs_for_admin()
    with session_scope() as db:
        events = list(
            db.scalars(
                select(ReportJobEvent)
                .order_by(ReportJobEvent.created_at.desc(), ReportJobEvent.id.desc())
                .limit(_EVENT_LIMIT)
            ).all()
        )
        stored = []
        for event in events:
            stored.append(
                {
                    "source": "event",
                    "job_id": None,
                    "event_id": event.id,
                    "user_id": event.user_id,
                    "course_id": event.course_id,
                    "term_id": event.term_id,
                    "status": event.status,
                    "error": event.error or "",
                    "duration_ms": event.duration_ms or 0,
                    "created_at": _iso(event.created_at),
                }
            )
        user_ids = {int(row["user_id"]) for row in live + stored if row.get("user_id") is not None}
        course_ids = {int(row["course_id"]) for row in live + stored if row.get("course_id") is not None}
        users, courses = _names(db, user_ids, course_ids)
    jobs = [_public_job(row, users, courses) for row in live]
    jobs.extend(_public_job(row, users, courses) for row in stored)
    return {"jobs": jobs}
