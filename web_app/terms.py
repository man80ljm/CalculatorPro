"""学期记录：学年学期、教师、班级、人数属于学期，不写回课程层。

课程 settings_json 只留课程层（名称、目标、关系表等）和各学期可继承的运行参数。
计算、导出、报告前用当前学期把学期字段合并进一份临时设置，不写回课程。
每个学期保存实际使用的大纲设置；新大纲供之后的新学期使用。
"""
from __future__ import annotations

import json
from typing import Any

from sqlalchemy import select
from sqlalchemy.orm import Session

from web_app.db import Course, CourseFile, Term, utcnow

TERM_BLOB_KEYS = (
    "mode",
    "spread_mode",
    "distribution",
    "noise_config",
    "report_style",
    "word_limit",
    "student_count",
)

# 不放进 TERM_BLOB_KEYS：那一组只在入参出现时才写入，缺了就会被整表丢掉。
_KEPT_BLOB_KEYS = ("last_achievement", "previous_achievement", "syllabus_settings", "syllabus_version", "syllabus_version_id", "legacy_report_dirty")


def _as_dict(value: Any) -> dict:
    return value if isinstance(value, dict) else {}


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _count(value: Any) -> int:
    try:
        number = int(float(str(value).strip()))
    except (TypeError, ValueError):
        return 0
    if number < 0:
        return 0
    return number


def parse_settings(raw: str | None) -> dict:
    try:
        data = json.loads(raw or "{}")
    except json.JSONDecodeError:
        return {}
    return data if isinstance(data, dict) else {}


def _json_copy(value: Any) -> Any:
    return json.loads(json.dumps(value, ensure_ascii=False))


NEW_TERM_LABEL = "新学期（请填写学年学期）"


def term_label(year_start: str, year_end: str, semester: str, school_year_term: str) -> str:
    if school_year_term:
        return school_year_term[:160]
    if year_start and year_end and semester:
        return f"{year_start}-{year_end}学年第{semester}学期"
    return NEW_TERM_LABEL


def apply_term_fields(term: Term, settings: dict, *, clear_identity: bool = False) -> None:
    """把设置里出现的学期字段写入学期行。没出现的键保持学期原值。

    clear_identity 用于新建另一学期：学年学期、教师、班级、上课人数、考核人数全部清空。
    """
    basic = _as_dict(settings.get("course_basic_info"))
    open_info = _as_dict(settings.get("course_open_info"))
    if clear_identity:
        term.year_start = ""
        term.year_end = ""
        term.semester = ""
        term.school_year_term = ""
        term.teacher = ""
        term.class_name = ""
        term.exam_count = 0
        term.student_count = 0
    else:
        if "year_start" in open_info:
            term.year_start = _text(open_info.get("year_start"))[:16]
        if "year_end" in open_info:
            term.year_end = _text(open_info.get("year_end"))[:16]
        if "semester" in open_info or "term" in open_info:
            term.semester = _text(open_info.get("semester") or open_info.get("term"))[:8]
        if "school_year_term" in basic:
            term.school_year_term = _text(basic.get("school_year_term"))[:80]
        if any(key in open_info for key in ("year_start", "year_end", "semester", "term")):
            if term.year_start and term.year_end and term.semester:
                term.school_year_term = f"{term.year_start}-{term.year_end}学年第{term.semester}学期"[:80]
        if "teacher" in open_info or "teacher" in basic:
            term.teacher = _text(open_info.get("teacher") or basic.get("teacher"))[:80]
        if "class_name" in basic:
            term.class_name = _text(basic.get("class_name"))[:120]
        if "exam_count" in basic:
            term.exam_count = _count(basic.get("exam_count"))
        if "student_count" in settings or "student_count" in basic:
            count = settings.get("student_count") if "student_count" in settings else None
            if count in (None, ""):
                count = basic.get("student_count")
            term.student_count = _count(count)
    old_blob = parse_settings(term.settings_json)
    blob = {}
    for key in TERM_BLOB_KEYS:
        if clear_identity and key == "student_count":
            continue
        if key in settings:
            blob[key] = settings.get(key)
    if "mode" not in blob:
        blob["mode"] = "forward"
    if not clear_identity and "major" in basic:
        blob["major"] = _text(basic.get("major"))[:120]
    # 本学期达成度和自动填上的上一轮，不跟开课信息一起被整表清掉。
    for key in _KEPT_BLOB_KEYS:
        if key in settings:
            blob[key] = _json_copy(settings.get(key))
        elif key in old_blob:
            blob[key] = _json_copy(old_blob[key])
    term.settings_json = json.dumps(blob, ensure_ascii=False)
    term.label = term_label(term.year_start, term.year_end, term.semester, term.school_year_term)
    term.updated_at = utcnow()


_TERM_BASIC_KEYS = ("school_year_term", "teacher", "major", "class_name", "student_count", "exam_count")
_TERM_OPEN_KEYS = ("year_start", "year_end", "semester", "term", "teacher")


def without_term_identity(settings: dict) -> dict:
    """课程层不保留学期身份字段的当前值，避免新建学期时串到别的学期。"""
    stored = json.loads(json.dumps(settings or {}, ensure_ascii=False))
    if not isinstance(stored, dict):
        return {}
    stored.pop("student_count", None)
    basic = stored.get("course_basic_info")
    if isinstance(basic, dict):
        for key in _TERM_BASIC_KEYS:
            basic.pop(key, None)
    opened = stored.get("course_open_info")
    if isinstance(opened, dict):
        for key in _TERM_OPEN_KEYS:
            opened.pop(key, None)
    return stored


def assign_course_settings(course: Course, settings: dict, *, term: Term | None = None) -> None:
    from web_app.course_versions import curriculum, SNAPSHOT, VERSION, VERSION_ID
    latest = parse_settings(course.settings_json)
    stored = curriculum(settings)
    number, version_id = latest.get(VERSION, 1), latest.get(VERSION_ID)
    if term is not None:
        blob = parse_settings(term.settings_json)
        number, version_id = blob.get(VERSION, number), blob.get(VERSION_ID, version_id)
        blob[SNAPSHOT] = stored
        blob[VERSION], blob[VERSION_ID] = number, version_id
        term.settings_json = json.dumps(blob, ensure_ascii=False)
        if version_id != latest.get(VERSION_ID):
            return
    stored[VERSION], stored[VERSION_ID] = number, version_id
    course.settings_json = json.dumps(stored, ensure_ascii=False)


def settings_from_term_row(row: dict) -> dict:
    """迁移退回时，把学期行写回课程设置。row 来自查询结果。"""
    settings = parse_settings(row.get("course_settings"))
    basic = dict(_as_dict(settings.get("course_basic_info")))
    open_info = dict(_as_dict(settings.get("course_open_info")))
    if row.get("year_start"):
        open_info["year_start"] = row["year_start"]
    if row.get("year_end"):
        open_info["year_end"] = row["year_end"]
    if row.get("semester"):
        open_info["semester"] = row["semester"]
        open_info["term"] = row["semester"]
    if row.get("teacher"):
        open_info["teacher"] = row["teacher"]
        basic["teacher"] = row["teacher"]
    if row.get("school_year_term"):
        basic["school_year_term"] = row["school_year_term"]
    if row.get("class_name"):
        basic["class_name"] = row["class_name"]
    if row.get("exam_count"):
        basic["exam_count"] = str(row["exam_count"])
    if row.get("student_count"):
        basic["student_count"] = str(row["student_count"])
        settings["student_count"] = int(row["student_count"])
    blob = parse_settings(row.get("term_settings"))
    if blob.get("major"):
        basic["major"] = str(blob["major"])
    for key in TERM_BLOB_KEYS:
        if key in blob and blob[key] not in (None, ""):
            settings[key] = blob[key]
    settings["course_basic_info"] = basic
    settings["course_open_info"] = open_info
    return settings


def merge_settings(settings: dict, term: Term) -> dict:
    """计算前把当前学期的身份字段叠进一份临时设置。空学期也覆盖，不用课程层里的旧值。"""
    from web_app.course_versions import term_curriculum
    merged = term_curriculum(settings, term)
    basic = dict(_as_dict(merged.get("course_basic_info")))
    open_info = dict(_as_dict(merged.get("course_open_info")))
    open_info["year_start"] = term.year_start or ""
    open_info["year_end"] = term.year_end or ""
    open_info["semester"] = term.semester or ""
    open_info["term"] = term.semester or ""
    open_info["teacher"] = term.teacher or ""
    basic["teacher"] = term.teacher or ""
    basic["school_year_term"] = term.school_year_term or ""
    basic["class_name"] = term.class_name or ""
    basic["exam_count"] = str(term.exam_count or "")
    basic["student_count"] = str(term.student_count or "")
    merged["student_count"] = int(term.student_count or 0)
    blob = parse_settings(term.settings_json)
    basic["major"] = _text(blob.get("major"))
    for key in ("mode", "spread_mode", "distribution", "report_style", "word_limit", "noise_config"):
        if key in blob and blob[key] not in (None, ""):
            merged[key] = blob[key]
    merged["course_basic_info"] = basic
    merged["course_open_info"] = open_info
    return merged


def term_public(term: Term, file_count: int = 0) -> dict:
    blob = parse_settings(term.settings_json)
    return {
        "id": term.id,
        "course_id": term.course_id,
        "label": term.label or NEW_TERM_LABEL,
        "is_current": bool(term.is_current),
        "year_start": term.year_start or "",
        "year_end": term.year_end or "",
        "semester": term.semester or "",
        "school_year_term": term.school_year_term or "",
        "teacher": term.teacher or "",
        "class_name": term.class_name or "",
        "major": blob.get("major") or "",
        "student_count": term.student_count or 0,
        "exam_count": term.exam_count or 0,
        "mode": blob.get("mode") or "forward",
        "spread_mode": blob.get("spread_mode") or "",
        "distribution": blob.get("distribution") or "",
        "report_style": blob.get("report_style") or "",
        "word_limit": blob.get("word_limit") or 0,
        "noise_config": blob.get("noise_config"),
        "file_count": file_count,
        "syllabus_version": blob.get("syllabus_version", 1),
    }


def list_terms(db: Session, course_id: int) -> list[Term]:
    return list(db.scalars(select(Term).where(Term.course_id == int(course_id)).order_by(Term.id)).all())


def _clear_current(db: Session, course_id: int) -> None:
    for term in list_terms(db, course_id):
        term.is_current = 0


def create_term(db: Session, course: Course, settings: dict | None = None, *, make_current: bool = True, clear_identity: bool = False) -> Term:
    from web_app.course_versions import freeze_existing_terms, curriculum, SNAPSHOT, VERSION, VERSION_ID
    version = freeze_existing_terms(db, course)
    if make_current:
        _clear_current(db, course.id)
    term = Term(
        user_id=course.user_id,
        course_id=course.id,
        is_current=1 if make_current else 0,
        created_at=utcnow(),
        updated_at=utcnow(),
    )
    apply_term_fields(term, settings or parse_settings(course.settings_json), clear_identity=clear_identity)
    blob = parse_settings(term.settings_json)
    blob[SNAPSHOT] = curriculum(settings or parse_settings(course.settings_json))
    blob[VERSION], blob[VERSION_ID] = version.number, version.id
    term.settings_json = json.dumps(blob, ensure_ascii=False)
    db.add(term)
    db.flush()
    return term


def current_term(db: Session, course: Course) -> Term | None:
    """当前学期。还没有导入过登记表时返回 None，不自动补一条空学期。"""
    term = db.scalar(
        select(Term).where(Term.course_id == course.id, Term.is_current == 1).order_by(Term.id.desc())
    )
    if term is not None:
        return term
    term = db.scalar(select(Term).where(Term.course_id == course.id).order_by(Term.id))
    if term is not None:
        _clear_current(db, course.id)
        term.is_current = 1
        return term
    return None


def select_term(db: Session, course: Course, term_id: int) -> Term:
    chosen = None
    for term in list_terms(db, course.id):
        if term.id == int(term_id):
            chosen = term
        term.is_current = 0
    if chosen is None:
        from web_app.db import NotFound

        raise NotFound()
    chosen.is_current = 1
    chosen.updated_at = utcnow()
    return chosen


def file_count(db: Session, term_id: int) -> int:
    rows = db.scalars(select(CourseFile.id).where(CourseFile.term_id == int(term_id))).all()
    return len(list(rows))
