"""Manual fallback for teaching documents, scoped to an account, course and semester."""
import json
import re
import uuid

from sqlalchemy import select
from web_app.db import TeachingMetadataOverride, TeacherProfile
from . import course_assignments as assignments
from .teacher_directory import teacher_info, key

FIELDS = {"teacher": "授课教师", "audience": "授课对象", "teaching_unit": "教研室"}


def scope(course_id=None, plan_id=None):
    if course_id is not None:
        return f"course:{int(course_id)}"
    if plan_id:
        return f"plan:{plan_id}"
    raise assignments.AssignmentError("请先保存课程或教案。")


def override(db, user_id, period, *, course_id=None, plan_id=None):
    return db.scalar(select(TeachingMetadataOverride).where(
        TeachingMetadataOverride.user_id == user_id,
        TeachingMetadataOverride.scope_key == scope(course_id, plan_id),
        TeachingMetadataOverride.term_id == f"{period['year_start']}-{period['semester']}"))


def values(result):
    return {"teacher": result["teacher"], "audience": "、".join(result["classes"]),
            "teaching_unit": (result.get("teacher_info") or {}).get("teaching_unit", "")}


def resolve(db, user_id, name, period, *, college="", course_id=None, plan_id=None):
    row = override(db, user_id, period, course_id=course_id, plan_id=plan_id)
    stored = json.loads(row.fields_json) if row else {}
    matched_name = stored.get("_course_match", "")
    matched_name = matched_name if isinstance(matched_name, str) else ""
    result = assignments.lookup(db, user_id, name, period, college=college,
                                matched_course_name=matched_name)
    manual = {field: stored[field] for field in FIELDS if field in stored}
    sources = {field: "timetable" if field != "teaching_unit" else "directory" for field in FIELDS}
    if "teacher" in manual:
        result["teacher"] = manual["teacher"]
        result["teacher_info"] = teacher_info(db, result["teacher"], college)
    if "audience" in manual:
        result["classes"] = [manual["audience"]] if manual["audience"] else []
    if "teaching_unit" in manual:
        result["teacher_info"] = {**(result.get("teacher_info") or {}), "teaching_unit": manual["teaching_unit"], "unit_source": "manual"}
    for field in manual:
        sources[field] = "manual"
    if manual:
        if result["teacher"] and result["classes"]:
            result["status"] = "manual"
        result["message"] = "已使用本课程、本学期保存的手填信息；后台课表更新不会覆盖这些内容。"
    elif result["status"] in {"missing", "review_required"}:
        result["message"] += " 也可以手动填写教师、授课对象和教研室。"
    result.update(college=college, has_manual=bool(manual), manual_fields=list(manual), sources=sources,
                  override_revision=row.revision if row else 0,
                  editable={field: field in manual or result["status"] in {"missing", "review_required", "manual"}
                            or (field == "teaching_unit" and not values(result)[field]) for field in FIELDS})
    return result


def save(db, user_id, name, period, edit, *, college="", course_id=None, plan_id=None):
    if not isinstance(edit, dict) or edit.get("mode") not in {"manual", "timetable", "course_match"}:
        raise assignments.AssignmentError("授课信息的填写方式不正确。")
    row = override(db, user_id, period, course_id=course_id, plan_id=plan_id)
    expected = edit.get("revision")
    if type(expected) is not int or expected != (row.revision if row else 0):
        raise assignments.AssignmentError("本学期授课信息已在其他页面修改，请重新打开后保存。", status=409, code="stale_teaching_metadata")
    previous = json.loads(row.fields_json) if row else {}
    if edit["mode"] == "timetable":
        kept = {"_course_match": previous["_course_match"]} if previous.get("_course_match") else {}
        if row:
            if kept:
                if kept != previous:
                    row.fields_json = json.dumps(kept, ensure_ascii=False)
                    row.revision += 1
            else:
                db.delete(row)
        db.flush()
        return resolve(db, user_id, name, period, college=college, course_id=course_id, plan_id=plan_id)
    if edit["mode"] == "course_match":
        title = edit.get("course_name")
        if not isinstance(title, str) or len(title) > 160:
            raise assignments.AssignmentError("请选择课表中的课程名称。")
        title = title.strip()
        candidates = assignments.suggest_courses(db, name, period)
        if title and not any(assignments.name_key(title) == assignments.name_key(item) for item in candidates):
            raise assignments.AssignmentError("此课程不在本学期的相近课表课程中，请重新查询。", status=409)
        updated = {**previous}
        if title:
            updated["_course_match"] = next(item for item in candidates
                                            if assignments.name_key(title) == assignments.name_key(item))
        else:
            updated.pop("_course_match", None)
        if row is None and updated:
            row = TeachingMetadataOverride(id=uuid.uuid4().hex, user_id=user_id, course_id=course_id,
                plan_id=plan_id if course_id is None else None, scope_key=scope(course_id, plan_id),
                term_id=f"{period['year_start']}-{period['semester']}", fields_json=json.dumps(updated, ensure_ascii=False))
            db.add(row)
        elif row and updated != previous:
            if updated:
                row.fields_json = json.dumps(updated, ensure_ascii=False)
                row.revision += 1
            else:
                db.delete(row)
        db.flush()
        return resolve(db, user_id, name, period, college=college, course_id=course_id, plan_id=plan_id)
    fields = edit.get("fields")
    if not isinstance(fields, dict) or not fields or set(fields) - set(FIELDS):
        raise assignments.AssignmentError("请填写需要保存的授课信息。")
    clean = {}
    for field, value in fields.items():
        limit = 2000 if field == "audience" else 160
        if not isinstance(value, str) or len(value) > limit or re.search(r"[\x00-\x1f]", value):
            raise assignments.AssignmentError(FIELDS[field] + "格式不正确或内容过长。")
        clean[field] = value.strip().removesuffix("系") if field == "teaching_unit" else value.strip()
    # An automatically matched unit stays directory-sourced; only explicit edits are pinned.
    updated = {**previous, **clean}
    if row is None:
        row = TeachingMetadataOverride(id=uuid.uuid4().hex, user_id=user_id, course_id=course_id,
            plan_id=plan_id if course_id is None else None, scope_key=scope(course_id, plan_id),
            term_id=f"{period['year_start']}-{period['semester']}", fields_json=json.dumps(updated, ensure_ascii=False))
        db.add(row)
    elif updated != previous:
        row.fields_json = json.dumps(updated, ensure_ascii=False)
        row.revision += 1
    db.flush()
    return resolve(db, user_id, name, period, college=college, course_id=course_id, plan_id=plan_id)


def search_teachers(db, query, college=""):
    query = key(query)
    if not query or len(query) > 160:
        return []
    rows = list(db.scalars(select(TeacherProfile).where(TeacherProfile.active == 1,
        TeacherProfile.name_key.contains(query, autoescape=True)).order_by(TeacherProfile.name).limit(25)))
    if college:
        rows = [row for row in rows if key(row.college) == key(college)]
    return [{"name": row.name, "college": row.college, "department": row.department,
             "teaching_unit": (row.teaching_unit or row.department).removesuffix("系")} for row in rows]
