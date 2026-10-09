"""报告是否与当前输入一致，只用于界面状态，不参与计算。"""
from __future__ import annotations

import hashlib
import json
from functools import lru_cache
from pathlib import Path

from sqlalchemy import select

from web_app.db import CourseFile, ReportTask
from web_app.terms import merge_settings, parse_settings


def _clean(value):
    if isinstance(value, dict):
        cleaned = {key: _clean(item) for key, item in value.items()}
        return {key: item for key, item in cleaned.items() if item not in (None, "", {}, [])}
    if isinstance(value, list):
        return [_clean(item) for item in value]
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return value


def _rows(value):
    return [row for row in value if isinstance(row, dict)] if isinstance(value, list) else []


def _number(value, default=None):
    try:
        return float(value)
    except (TypeError, ValueError):
        return value if value not in (None, "") else default


def settings_signature(settings):
    from web_app.service import DIST_MAP, SPREAD_MAP, _ratios
    settings = settings if isinstance(settings, dict) else {}
    payload = settings.get("relation_payload")
    payload = payload if isinstance(payload, dict) else {}
    links = [{"name": link.get("name"), "ratio": _number(link.get("ratio")),
              "methods": [{"name": method.get("name"), "supports": method.get("supports"),
                           "subtotal": _number(method.get("subtotal"))} for method in _rows(link.get("methods"))]}
             for link in _rows(payload.get("links"))]
    values = {key: settings.get(key) for key in ("course_open_info", "course_basic_info", "student_count", "course_description", "objective_requirements")}
    # 表单以字符串保存基本信息；数字和字符串的相同写法视为同一内容。
    for key in ("course_open_info", "course_basic_info"):
        if isinstance(values[key], dict):
            values[key] = {name: str(value).strip() if isinstance(value, (str, int, float)) else value
                           for name, value in values[key].items()}
    values["student_count"] = _number(values["student_count"])
    # 自动保存会补空字段、行表头和派生合计；这些不是输入内容的变化。
    values.update(mode=settings.get("mode") or "forward", ratios=_ratios(settings),
                  relation_payload={"objectives_count": _number(payload.get("objectives_count")), "links": links},
                  grad_req_map=[{k: item.get(k) for k in ("requirement", "indicator", "strength")} for item in _rows(settings.get("grad_req_map"))],
                  spread_mode=SPREAD_MAP.get(str(settings.get("spread_mode")), "medium"),
                  distribution=DIST_MAP.get(str(settings.get("distribution")), "normal"),
                  noise_config=settings.get("noise_config") if settings.get("mode") == "reverse" else None,
                  report_style=settings.get("report_style") or "专业", word_limit=_number(settings.get("word_limit") or 200))
    text = json.dumps(_clean(values), ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def input_signature(settings, grade_digest, previous_digest, course_name):
    values = [settings_signature(settings), grade_digest, previous_digest or "", course_name]
    return hashlib.sha256(json.dumps(values, ensure_ascii=False).encode("utf-8")).hexdigest()


def submitted_signature(settings, excel, previous, course_name):
    return input_signature(settings, hashlib.sha256(excel).hexdigest(),
                           hashlib.sha256(previous).hexdigest() if previous else "", course_name)


@lru_cache(maxsize=512)
def _file_digest(filename, size, modified):
    digest = hashlib.sha256()
    with Path(filename).open("rb") as handle:
        while block := handle.read(1024 * 1024):
            digest.update(block)
    return digest.hexdigest()


def current_signature(db, course, term):
    from web_app.storage import resolve_stored
    digests = {}
    for kind in ("grade", "previous"):
        row = db.scalar(select(CourseFile).where(CourseFile.course_id == course.id, CourseFile.term_id == term.id,
                        CourseFile.user_id == course.user_id, CourseFile.kind == kind).order_by(CourseFile.id.desc()).limit(1))
        if row is None:
            digests[kind] = ""
            continue
        path = resolve_stored(row.user_id, row.course_id, row.stored_name)
        if not path.is_file():
            digests[kind] = "missing:" + str(row.id)
            continue
        stat = path.stat()
        digests[kind] = _file_digest(str(path), stat.st_size, stat.st_mtime_ns)
    return input_signature(merge_settings(parse_settings(course.settings_json), term), digests["grade"], digests["previous"], course.name)


def mark_legacy_changed(db, term):
    if term is None:
        return
    blob = parse_settings(term.settings_json)
    blob["legacy_report_dirty"] = True
    term.settings_json = json.dumps(blob, ensure_ascii=False)


def course_report_state(db, course, term):
    if term is None:
        return None
    from web_app.storage import resolve_stored
    archives = list(db.scalars(select(CourseFile).where(CourseFile.course_id == course.id,
                    CourseFile.user_id == course.user_id, CourseFile.term_id == term.id,
                    CourseFile.kind == "report").order_by(CourseFile.id.desc())))
    archive = next((row for row in archives if row.original_name.lower().endswith(".zip")
                    and resolve_stored(row.user_id, row.course_id, row.stored_name).is_file()), None)
    if archive is None:
        return None
    task = db.scalar(select(ReportTask).where(ReportTask.user_id == course.user_id, ReportTask.course_id == course.id,
                     ReportTask.term_id == term.id, ReportTask.status == "success", ReportTask.archive_file_id == archive.id))
    view = parse_settings(task.public_json) if task else {}
    original = view.get("source_signature")
    stale = original != current_signature(db, course, term) if original else bool(parse_settings(term.settings_json).get("legacy_report_dirty"))
    return {"status": "stale" if stale else "ready", "message": "内容已修改，旧报告仍可下载。" if stale else "报告已生成，可以下载。",
            "archive_file_id": archive.id, "download_url": f"/api/courses/{course.id}/files/{archive.id}",
            "created_at": archive.created_at.isoformat() + "Z", "summary": view.get("summary"), "legacy": not bool(original)}
