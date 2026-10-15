"""Only timetable data supplies teaching-calendar and lesson-plan teacher/classes."""
import hashlib
import json
from difflib import SequenceMatcher
import re
import unicodedata
import uuid

from sqlalchemy import select
from web_app.db import ScheduleDocument, ScheduleRecord, TeachingAssignmentSelection
from web_app.service import ServiceError
from web_app.terms import parse_settings
from web_app.teaching_calendar.academic_term import current_academic_term


class AssignmentError(ServiceError):
    pass


def course_name(course, outline_name=""):
    basic = parse_settings(course.settings_json).get("course_basic_info") or {}
    return basic.get("course_name") or outline_name or course.name


def course_college(course):
    basic = parse_settings(course.settings_json).get("course_basic_info") or {}
    return basic.get("college") or ""


def name_key(value):
    # Exact name comparison after full-width/whitespace normalization; no fuzzy matching.
    value = unicodedata.normalize("NFKC", str(value)).strip().strip("《》")
    return re.sub(r"\s+", "", value).casefold()


def academic_period(year_start=None, year_end=None, semester=None):
    if year_start is None and year_end is None and semester is None:
        return current_academic_term()
    start, end, term = str(year_start or ""), str(year_end or ""), str(semester or "")
    if (not re.fullmatch(r"\d{4}", start) or not re.fullmatch(r"\d{4}", end)
            or int(end) != int(start) + 1 or term not in {"1", "2"}):
        raise AssignmentError("请填写相邻两个年份和第 1 或第 2 学期。", status=400)
    return {"year_start": start, "year_end": end, "semester": term,
            "label": f"{start}-{end} 学年第 {term} 学期"}


def _json(raw):
    try:
        value = json.loads(raw)
        return value if isinstance(value, dict) else {}
    except (TypeError, ValueError):
        return {}


def _names(value):
    if not isinstance(value, list): return []
    return list(dict.fromkeys(item.strip() for item in value
                             if isinstance(item, str) and item.strip() and len(item) <= 160))


def _same_period(document, period):
    term = _json(document.term)
    return (str(term.get("academicYear", "")).replace("—", "-").replace("–", "-")
            == f"{period['year_start']}-{period['year_end']}"
            and str(term.get("semester")) == period["semester"])


def suggest_courses(db, name, period):
    """Offer similar timetable titles for a teacher to confirm; never assign one silently."""
    term_id = f"{period['year_start']}-{period['semester']}"
    rows = db.execute(select(ScheduleRecord.course_name, ScheduleDocument)
                      .join(ScheduleDocument, ScheduleRecord.document_id == ScheduleDocument.id)
                      .where(ScheduleRecord.term_id == term_id)).all()
    wanted = name_key(name)
    scores = {}
    if not wanted:
        return []
    for title, document in rows:
        if not _same_period(document, period):
            continue
        found = name_key(title)
        if found == wanted or not found:
            continue
        ratio = SequenceMatcher(None, wanted, found).ratio()
        if ratio < 0.72 and not (min(len(wanted), len(found)) >= 4
                                 and (wanted in found or found in wanted)):
            continue
        scores[title] = max(scores.get(title, 0), ratio)
    return [title for title, _ in sorted(scores.items(), key=lambda item: (-item[1], item[0]))[:6]]


def lookup(db, user_id, name, period, *, college="", matched_course_name=""):
    from .teacher_directory import teacher_info
    term_id = f"{period['year_start']}-{period['semester']}"
    key = name_key(name)
    target = name_key(matched_course_name or name)
    result = {"course_name": name, "matched_course_name": matched_course_name,
              "course_matches": [], "year_start": period["year_start"], "year_end": period["year_end"],
              "semester": period["semester"], "status": "missing", "teacher": "", "classes": [],
              "options": [], "message": f"后台课表中未找到《{name}》{period['label']}的授课安排，请联系管理员核对课表。"}
    records = db.execute(select(ScheduleRecord, ScheduleDocument)
                         .join(ScheduleDocument, ScheduleRecord.document_id == ScheduleDocument.id)
                         .where(ScheduleRecord.term_id == term_id)
                         .order_by(ScheduleDocument.created_at.desc(), ScheduleDocument.id.desc(), ScheduleRecord.id)).all()
    matches = []
    for row, document in records:
        if name_key(row.course_name) == target and _same_period(document, period):
            matches.append((row, document))
    if not matches or matched_course_name:
        result["course_matches"] = suggest_courses(db, name, period)
    if not matches:
        if result["course_matches"] and not matched_course_name:
            result["message"] = "课表中有相近课程名，请先确认对应关系，再选择本学期的任课教师。"
        elif matched_course_name:
            result["message"] = f"已关联的课表课程《{matched_course_name}》本学期没有可用记录，请重新选择。"
        return result
    # A later import for this course/semester supersedes that course's previous import.
    latest_document = matches[0][1].id
    groups = {}
    for row, document in matches:
        if document.id != latest_document: continue
        data = _json(row.payload)
        teachers, classes = _names(data.get("teachers")), _names(data.get("classes"))
        issues = data.get("issues") or []
        uncertain = not data.get("reviewed") and any(
            isinstance(issue, str) and any(word in issue for word in ("教师", "班级", "课程名称"))
            for issue in issues)
        if len(teachers) != 1 or not classes or uncertain:
            result.update(status="review_required", message="后台课表中的任课教师或班级尚待核对，已留空，请联系管理员确认。")
            return result
        teacher = teachers[0]
        groups.setdefault(teacher, [])
        for label in classes:
            if label not in groups[teacher]: groups[teacher].append(label)
    options = [{"teacher": teacher, "classes": sorted(groups[teacher])} for teacher in sorted(groups)]
    for item in options: item["teacher_info"] = teacher_info(db, item["teacher"], college)
    result["options"] = options
    chosen = db.scalar(select(TeachingAssignmentSelection.teacher).where(
        TeachingAssignmentSelection.user_id == user_id,
        TeachingAssignmentSelection.course_key == hashlib.sha256(key.encode()).hexdigest(),
        TeachingAssignmentSelection.term_id == term_id))
    selected = options[0] if len(options) == 1 else next((item for item in options if item["teacher"] == chosen), None)
    if selected:
        result.update(status="ready", teacher=selected["teacher"], classes=selected["classes"], teacher_info=selected["teacher_info"], message="任课教师、授课对象已从后台课表自动填入。")
    else:
        result.update(status="selection_required", message="此课程本学期有多位任课教师，请选择一次；本学期的教学日历和教案会沿用选择，新学期需重新确认。")
    if matched_course_name:
        result["message"] = f"已对应后台课表《{matched_course_name}》。" + result["message"]
    return result


def choose(db, user_id, name, period, teacher, *, college="", matched_course_name=""):
    result = lookup(db, user_id, name, period, college=college, matched_course_name=matched_course_name)
    if not isinstance(teacher, str) or not any(item["teacher"] == teacher for item in result["options"]):
        raise AssignmentError("请选择后台课表中这门课、本学期的任课教师。", status=400, code="invalid_teaching_assignment")
    key = hashlib.sha256(name_key(name).encode()).hexdigest()
    term_id = f"{period['year_start']}-{period['semester']}"
    row = db.scalar(select(TeachingAssignmentSelection).where(
        TeachingAssignmentSelection.user_id == user_id, TeachingAssignmentSelection.course_key == key,
        TeachingAssignmentSelection.term_id == term_id))
    if row is None:
        row = TeachingAssignmentSelection(id=uuid.uuid4().hex, user_id=user_id, course_key=key, term_id=term_id, teacher=teacher)
        db.add(row)
    else: row.teacher = teacher
    db.flush()
    return lookup(db, user_id, name, period, college=college, matched_course_name=matched_course_name)


def require_choice(result):
    if result["status"] == "selection_required":
        raise AssignmentError(result["message"], status=409, code="teaching_assignment_required", teaching_assignment=result)


def fill_context(context, assignment):
    context["授课教师"] = assignment["teacher"]
    context["授课对象"] = "、".join(assignment["classes"])
    context["教研室"] = (assignment.get("teacher_info") or {}).get("teaching_unit", "")
    return context
