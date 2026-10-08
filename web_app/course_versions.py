"""大纲版本和学期设置快照：更新课程后，旧学期仍用原目标、比例和对应关系。"""
from __future__ import annotations

import json
from sqlalchemy import select
from web_app.db import CourseVersion, Term

SNAPSHOT = "syllabus_settings"
VERSION = "syllabus_version"
VERSION_ID = "syllabus_version_id"


def curriculum(settings: dict) -> dict:
    from web_app.terms import without_term_identity
    copied = without_term_identity(settings)
    for key in (SNAPSHOT, VERSION, VERSION_ID, "last_achievement", "previous_achievement", "term_id"):
        copied.pop(key, None)
    return copied


def term_curriculum(settings: dict, term) -> dict:
    from web_app.terms import parse_settings
    blob = parse_settings(term.settings_json) if term is not None else {}
    snapshot = blob.get(SNAPSHOT)
    if isinstance(snapshot, dict):
        result = curriculum(snapshot)
        result[VERSION] = blob.get(VERSION, 1)
        result[VERSION_ID] = blob.get(VERSION_ID)
        return result
    return json.loads(json.dumps(settings or {}, ensure_ascii=False))


def freeze_existing_terms(db, course) -> CourseVersion:
    from web_app.terms import parse_settings
    settings = parse_settings(course.settings_json)
    latest = db.scalar(select(CourseVersion).where(CourseVersion.course_id == course.id).order_by(CourseVersion.number.desc()))
    if latest is None:
        latest = CourseVersion(user_id=course.user_id, course_id=course.id, number=1,
                               settings_json=json.dumps(curriculum(settings), ensure_ascii=False))
        db.add(latest)
        db.flush()
    settings[VERSION] = latest.number
    settings[VERSION_ID] = latest.id
    course.settings_json = json.dumps(settings, ensure_ascii=False)
    for term in db.scalars(select(Term).where(Term.course_id == course.id)).all():
        blob = parse_settings(term.settings_json)
        if not isinstance(blob.get(SNAPSHOT), dict):
            blob[SNAPSHOT] = curriculum(settings)
            blob[VERSION] = latest.number
            blob[VERSION_ID] = latest.id
            term.settings_json = json.dumps(blob, ensure_ascii=False)
    return latest


def new_syllabus_version(db, course, settings: dict, *, expected_version=None) -> CourseVersion:
    from web_app.service import ServiceError
    from web_app.terms import parse_settings
    old = freeze_existing_terms(db, course)
    if expected_version is not None and str(expected_version) != str(old.number):
        raise ServiceError("大纲已在其他窗口更新，请重新打开课程后再确认。", status=409, code="stale_version")
    # 保存旧版最新的课程设置；历史学期的独立快照不改动。
    old.settings_json = json.dumps(curriculum(parse_settings(course.settings_json)), ensure_ascii=False)
    version = CourseVersion(user_id=course.user_id, course_id=course.id, number=old.number + 1,
                            settings_json=json.dumps(curriculum(settings), ensure_ascii=False))
    db.add(version)
    db.flush()
    stored = curriculum(settings)
    stored[VERSION] = version.number
    stored[VERSION_ID] = version.id
    course.settings_json = json.dumps(stored, ensure_ascii=False)
    return version
