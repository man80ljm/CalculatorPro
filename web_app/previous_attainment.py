"""上一轮教学达成度：算完写入本学期，导入更早学期的成绩时自动带出来。"""
from __future__ import annotations

import json
import math
import re
from io import BytesIO

from openpyxl import Workbook
from sqlalchemy import select

from web_app.db import CourseFile, NotFound, utcnow
from web_app.storage import resolve_stored, save_blob
from web_app.terms import list_terms, parse_settings

_OBJECTIVE = re.compile(r"^课程目标(\d+)$")
_TOTAL_SOURCE_KEYS = ("总达成度", "课程总目标", "课程目标达成值", "课程总达成值", "total_value")
AUTO_PREVIOUS_NAME = "上一轮达成度-自动.xlsx"


def _finite_number(value):
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        number = float(value)
    elif isinstance(value, str):
        text = value.strip()
        if not text or text == "—":
            return None
        try:
            number = float(text)
        except ValueError:
            return None
    else:
        return None
    if not math.isfinite(number):
        return None
    return number


def _optional_int(value) -> int | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else None
    text = str(value).strip()
    if not text or not re.fullmatch(r"\d+", text):
        return None
    return int(text)


def normalize_achievement(raw) -> dict:
    """分目标用「课程目标N」，总体用「课程总目标」。源数据里的「总达成度」映射过去。"""
    if not isinstance(raw, dict):
        return {}
    normalized = {}
    for key, value in raw.items():
        name = str(key).strip()
        if not _OBJECTIVE.fullmatch(name):
            continue
        number = _finite_number(value)
        if number is None:
            continue
        normalized[name] = number
    for key in _TOTAL_SOURCE_KEYS:
        if key not in raw:
            continue
        number = _finite_number(raw.get(key))
        if number is None:
            continue
        normalized["课程总目标"] = number
        break
    return normalized


def _academic_key(term) -> tuple[int, int, int] | None:
    start = _optional_int(getattr(term, "year_start", None))
    end = _optional_int(getattr(term, "year_end", None))
    semester = _optional_int(getattr(term, "semester", None))
    if start is None or end is None or semester is None:
        return None
    return (start, end, semester)


def term_sort_key(term) -> tuple:
    """按学年起、学年止、学期排序。缺少年学期时退回学期 id。"""
    academic = _academic_key(term)
    if academic is None:
        return (int(getattr(term, "id", 0) or 0),)
    return academic


def _term_id(term) -> int:
    return int(getattr(term, "id", 0) or 0)


def _is_strictly_earlier(other, current) -> bool:
    """同学年学期不算更早。缺少年学期时不能靠更小的 id 把自己当成上一轮。"""
    if _term_id(other) == _term_id(current):
        return False
    other_year = _academic_key(other)
    current_year = _academic_key(current)
    if other_year is not None and current_year is not None:
        return other_year < current_year
    if other_year is None and current_year is None:
        return False
    return term_sort_key(other) < term_sort_key(current)


def _recency(term) -> tuple:
    academic = _academic_key(term)
    if academic is None:
        return (0, _term_id(term))
    return (1, academic[0], academic[1], academic[2], _term_id(term))


def _stored_last_achievement(term) -> dict:
    blob = parse_settings(getattr(term, "settings_json", None))
    raw = blob.get("last_achievement")
    if not isinstance(raw, dict):
        return {}
    return normalize_achievement(raw)


def find_prior_term_with_achievement(db, course, current_term):
    """同一门课里，严格早于当前学期、且已有达成度的最近一条。

    同学年学期有多条（不同班级）时，它们彼此都不是上一轮；相对更早的学年里若有多条，取 id 较大的一条。
    """
    best = None
    best_rank = None
    best_achievement = None
    for term in list_terms(db, course.id):
        if not _is_strictly_earlier(term, current_term):
            continue
        achievement = _stored_last_achievement(term)
        if not achievement:
            continue
        rank = _recency(term)
        if best is None or rank > best_rank:
            best = term
            best_rank = rank
            best_achievement = achievement
    if best is None:
        return None, None
    return best, best_achievement


def _objective_rows(achievement: dict) -> list[tuple[str, float]]:
    numbered = []
    for name, value in achievement.items():
        match = _OBJECTIVE.fullmatch(name)
        if not match:
            continue
        numbered.append((int(match.group(1)), name, float(value)))
    numbered.sort()
    return [(name, value) for _, name, value in numbered]


def achievement_to_previous_xlsx(achievement: dict) -> bytes:
    """生成 load_previous_achievement 能读的表：课程分目标 + 分目标达成值。"""
    normalized = normalize_achievement(achievement)
    book = Workbook()
    sheet = book.active
    sheet.title = "上一轮达成度"
    sheet.append(["课程分目标", "分目标达成值"])
    for name, value in _objective_rows(normalized):
        sheet.append([name, value])
    if "课程总目标" in normalized:
        sheet.append(["课程目标达成值", float(normalized["课程总目标"])])
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def _put_term_setting(term, key: str, value) -> None:
    blob = parse_settings(term.settings_json)
    blob[key] = value
    term.settings_json = json.dumps(blob, ensure_ascii=False)
    term.updated_at = utcnow()


def store_last_achievement(term, achievement) -> dict:
    """计算或报告成功后，把本学期达成度留在该学期上。空结果不覆盖已有值。"""
    normalized = normalize_achievement(achievement)
    if not normalized:
        return {}
    _put_term_setting(term, "last_achievement", normalized)
    return normalized


def record_last_achievement(user_id: int, course_id: int, term_id: int, achievement) -> dict:
    """报告任务在自己的会话里写入本学期达成度。"""
    normalized = normalize_achievement(achievement)
    if not normalized or not term_id:
        return {}
    from web_app.db import Term, session_scope

    with session_scope() as db:
        term = db.get(Term, int(term_id))
        if term is None or int(term.user_id) != int(user_id) or int(term.course_id) != int(course_id):
            return {}
        return store_last_achievement(term, normalized)


def _replace_previous(db, user_id: int, course_id: int, term_id: int, filename: str, data: bytes):
    saved = save_blob(db, user_id, course_id, filename, data, "previous", term_id=term_id)
    older = list(
        db.scalars(
            select(CourseFile).where(
                CourseFile.user_id == int(user_id),
                CourseFile.course_id == int(course_id),
                CourseFile.term_id == int(term_id),
                CourseFile.kind == "previous",
                CourseFile.id != saved.id,
            )
        ).all()
    )
    for row in older:
        try:
            resolve_stored(row.user_id, row.course_id, row.stored_name).unlink(missing_ok=True)
        except (NotFound, OSError):
            pass
        db.delete(row)
    return saved


def _not_filled() -> dict:
    return {"previous_filled": False, "prior_term_id": None, "previous_achievement": {}}


def autofill_previous_from_prior_term(db, user_id: int, course, term) -> dict:
    """有更早学期的达成度就写成当前学期的上一轮表。没有则不写，也不拿本学期自己的结果充数。"""
    prior, achievement = find_prior_term_with_achievement(db, course, term)
    if prior is None or not achievement or _term_id(prior) == _term_id(term):
        return _not_filled()
    _replace_previous(
        db,
        user_id,
        course.id,
        term.id,
        AUTO_PREVIOUS_NAME,
        achievement_to_previous_xlsx(achievement),
    )
    _put_term_setting(term, "previous_achievement", achievement)
    return {
        "previous_filled": True,
        "prior_term_id": int(prior.id),
        "previous_achievement": achievement,
    }
