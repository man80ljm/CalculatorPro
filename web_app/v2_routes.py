"""学期、大纲建课、成绩登记表导入。挂到现有应用上，旧接口仍走当前学期。"""
from __future__ import annotations

import json
import re

from fastapi import Request
from fastapi.responses import JSONResponse

from web_app.db import NotFound, session_scope, utcnow
from web_app.grade_register import (
    detect_register_mode,
    forward_workbook,
    headcount_warnings,
    parse_register,
    ratio_warnings,
    register_course_message,
    reverse_workbook,
)
from web_app.limiter import syllabus_limiter
from web_app.relation_grid import absorb_relation_grid
from web_app.previous_attainment import autofill_previous_from_prior_term
from web_app.storage import resolve_stored, save_blob
from web_app.syllabus.major import derive_major_from_class
from web_app.syllabus.rules import settings_from_wizard, term_parts, validate_relation_grid
from web_app.syllabus.service import extract_draft
from web_app.terms import (
    apply_term_fields,
    assign_course_settings,
    create_term,
    current_term,
    file_count,
    list_terms,
    parse_settings,
    select_term,
    term_public,
    without_term_identity,
)


MIXED_CLASS_NAME = "专业任选"

def _replace_term_grade(db, user_id: int, course_id: int, term_id: int, filename: str, data: bytes):
    """同一学期只保留这一份成绩文件。"""
    from sqlalchemy import select

    from web_app.db import CourseFile, NotFound

    saved = save_blob(db, user_id, course_id, filename, data, "grade", term_id=term_id)
    older = list(
        db.scalars(
            select(CourseFile).where(
                CourseFile.user_id == user_id,
                CourseFile.course_id == course_id,
                CourseFile.term_id == term_id,
                CourseFile.kind == "grade",
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


def apply_parsed_register(db, course, term, parsed: dict, *, confirmed: bool, user_id: int, mode: str = "auto") -> dict:
    """把已解析的成绩登记表写入当前学期。与主页导入同一套规则，不调用模型。"""
    from web_app.service import ServiceError

    requested = str(mode or "auto").strip().lower()
    if requested not in {"auto", "forward", "reverse"}:
        raise ServiceError("导入模式只能是 auto、forward 或 reverse")
    from web_app.course_versions import term_curriculum
    settings = term_curriculum(parse_settings(course.settings_json), term)
    mismatch = register_course_message(settings, course.name, parsed)
    if mismatch and not confirmed:
        raise ServiceError(mismatch, status=409)
    payload = settings.get("relation_payload") or {}
    links = payload.get("links") if isinstance(payload, dict) else []
    detection = detect_register_mode(parsed, links or [])
    if requested == "forward" and detection["mode"] != "forward":
        raise ServiceError(detection["reason"] or "无法按正向导入：考核方式列不完整")
    applied = detection["mode"] if requested == "auto" else requested
    if applied == "forward":
        workbook, map_notes = forward_workbook(parsed, links or [], detection)
        grade_name = "成绩登记表-正向.xlsx"
    else:
        workbook, map_notes = reverse_workbook(parsed, links or [])
        grade_name = "成绩登记表-逆向.xlsx"
    ratios = settings.get("ratios") or {}
    warnings = ratio_warnings(parsed["percents"], ratios) + map_notes
    warnings.extend(headcount_warnings(term.student_count, term.exam_count, parsed["student_count"], parsed.get("exam_count")))
    conflicts = []
    filled = []

    def put(new: str, label: str) -> str:
        text = str(new or "").strip()
        if text:
            filled.append(label)
        return text

    basic = dict(settings.get("course_basic_info") or {})
    opened = dict(settings.get("course_open_info") or {})
    # 登记表有课程性质时优先于大纲；缺失时保留现有课程性质。
    if parsed.get("course_type"):
        basic["course_type"] = put(parsed["course_type"], "课程性质")
    # 登记表是这一学期的事实来源：有值就覆盖，不再留「未覆盖」冲突。
    if parsed.get("year_start"):
        opened["year_start"] = put(parsed["year_start"], "学年起")
    if parsed.get("year_end"):
        opened["year_end"] = put(parsed["year_end"], "学年止")
    if parsed.get("semester"):
        opened["semester"] = put(parsed["semester"], "学期")
    if parsed.get("teacher"):
        opened["teacher"] = put(parsed["teacher"], "任课教师")
        basic["teacher"] = opened["teacher"]
    if parsed.get("school_year_term"):
        basic["school_year_term"] = put(parsed["school_year_term"], "学年学期")
    classes = [str(name).strip() for name in (parsed.get("classes") or []) if str(name).strip()]
    if len(classes) == 1:
        basic["class_name"] = put(classes[0], "上课班级")
        major = derive_major_from_class(classes[0]) or ""
        basic["major"] = major
        if major:
            filled.append("上课专业")
    elif len(classes) > 1:
        basic["class_name"] = put(MIXED_CLASS_NAME, "上课班级")
        # 混合上课只统计整张表，不从其中一个班级推断上课专业。
        basic["major"] = ""
    settings["student_count"] = parsed["student_count"]
    basic["student_count"] = str(parsed["student_count"])
    filled.append("上课人数")
    if parsed.get("exam_count"):
        basic["exam_count"] = str(parsed["exam_count"])
        filled.append("考核人数")
    settings["mode"] = applied
    settings["course_basic_info"] = basic
    settings["course_open_info"] = opened
    apply_term_fields(term, settings)
    assign_course_settings(course, settings, term=term)
    _replace_term_grade(db, user_id, course.id, term.id, grade_name, workbook)
    course.updated_at = utcnow()
    reason = detection["reason"]
    if requested in {"forward", "reverse"} and requested != detection["mode"]:
        forced = "正向" if requested == "forward" else "逆向"
        reason = f"已按指定改为{forced}导入。{detection['reason']}"
    return {
        "student_count": parsed["student_count"],
        "exam_count": parsed["exam_count"],
        "classes": parsed["classes"],
        "class_name": basic.get("class_name", ""),
        "mixed_classes": len(classes) > 1,
        "percents": parsed["percents"],
        "ratios": ratios,
        "warnings": warnings,
        "conflicts": conflicts,
        "filled": filled,
        "mode": applied,
        "reason": reason,
        "detection": {
            "mode": detection["mode"],
            "reason": detection["reason"],
            "matched": detection["matched"],
        },
        "term_id": term.id,
        "consistent": not warnings,
    }


_DASH_TERM = re.compile(r"(\d{4})\s*[-–—]\s*(\d{4})\s*[-–—]\s*([12])\b")


def normalize_term_parts(year_start: str, year_end: str, semester: str, school_year_term: str = "") -> tuple[str, str, str]:
    """学年起、学年止、学期。支持「2024-2025学年第1学期」和「2024-2025-1」。"""
    start, end, sem = str(year_start or "").strip(), str(year_end or "").strip(), str(semester or "").strip()
    if start and end and sem:
        return start[:16], end[:16], sem[:8]
    text = str(school_year_term or "").strip()
    parsed = term_parts(text)
    if parsed[0]:
        return parsed
    match = _DASH_TERM.search(text)
    if match:
        return match.group(1), match.group(2), match.group(3)
    return start[:16], end[:16], sem[:8]


def _term_parts_of(term) -> tuple[str, str, str]:
    return normalize_term_parts(term.year_start, term.year_end, term.semester, term.school_year_term)


def _replace_prompt(year_start: str, year_end: str, semester: str, class_name: str) -> str:
    label = f"{year_start}-{year_end}学年第{semester}学期"
    if class_name:
        label = f"{label} · {class_name}"
    return f"将替换该学期（{label}）的成绩，确定吗？"


def import_parsed_register(
    db,
    course,
    parsed: dict,
    *,
    confirmed: bool,
    user_id: int,
    mode: str = "auto",
    year_start: str = "",
    year_end: str = "",
    semester: str = "",
    class_name: str = "",
) -> dict:
    """整表导入，混合班级统一按「专业任选」匹配学期。"""
    from web_app.service import ServiceError

    start, end, sem = normalize_term_parts(
        year_start or parsed.get("year_start") or "",
        year_end or parsed.get("year_end") or "",
        semester or parsed.get("semester") or "",
        parsed.get("school_year_term") or "",
    )
    classes = [str(name).strip() for name in (parsed.get("classes") or []) if str(name).strip()]
    # 兼容旧客户端的 class_name 参数，但它不再筛选学生。
    chosen = MIXED_CLASS_NAME if len(classes) > 1 else (classes[0] if classes else "")
    needs = []
    if not (start and end and sem):
        needs.append("term")
    if needs:
        detail = []
        if "term" in needs:
            detail.append("登记表没有学年或学期，请补填学年起、学年止和学期")
        raise ServiceError("；".join(detail) + "。", status=409, code="need_input", needs=needs, classes=classes)

    match_class = bool(classes)
    imported = dict(parsed)
    imported["year_start"] = start
    imported["year_end"] = end
    imported["semester"] = sem
    imported["school_year_term"] = f"{start}-{end}学年第{sem}学期"

    terms = list_terms(db, course.id)
    if match_class:
        matches = [item for item in terms if _term_parts_of(item) == (start, end, sem) and (item.class_name or "") == chosen]
    else:
        matches = [item for item in terms if _term_parts_of(item) == (start, end, sem)]
        if len(matches) > 1:
            raise ServiceError("该学年学期已有多个班级，登记表没有班级，无法确定要替换哪一条。")

    from web_app.course_versions import term_curriculum
    settings = term_curriculum(parse_settings(course.settings_json), matches[0]) if matches else parse_settings(course.settings_json)
    mismatch = register_course_message(settings, course.name, imported)
    if mismatch and not confirmed:
        raise ServiceError(mismatch, status=409)
    if matches and not confirmed:
        raise ServiceError(
            _replace_prompt(start, end, sem, chosen or matches[0].class_name or ""),
            status=409,
            code="replace",
            term_id=matches[0].id,
        )
    if matches:
        term = matches[0]
        select_term(db, course, term.id)
    else:
        term = create_term(db, course, settings, make_current=True, clear_identity=True)
    summary = apply_parsed_register(db, course, term, imported, confirmed=confirmed, user_id=user_id, mode=mode)
    summary["created_term"] = not matches
    summary.update(autofill_previous_from_prior_term(db, user_id, course, term))
    if summary.get("previous_filled"):
        prior = next((item for item in terms if item.id == summary.get("prior_term_id")), None)
        if prior and parse_settings(prior.settings_json).get("syllabus_version", 1) != parse_settings(term.settings_json).get("syllabus_version", 1):
            summary["warnings"].append("本学期大纲已更新，请核对上一轮达成度是否对应相同课程目标。")
            summary["consistent"] = False
    return summary


def _read_limited(upload, limit: int) -> bytes:
    from web_app.service import ServiceError

    if upload is None or not getattr(upload, "filename", ""):
        raise ServiceError("请选择文件")
    data = upload.file.read(limit + 1) if hasattr(upload, "file") else None
    return data


def attach(application) -> None:
    from web_app.app import (
        _api,
        _as_upload,
        _close_form,
        _course_dict,
        _identity,
        _json_body,
        _json_error,
        _optional_form,
        _owned_course,
        max_upload_bytes,
    )
    from web_app.service import ServiceError

    async def _bytes(upload, required: bool) -> bytes | None:
        if upload is None:
            if required:
                raise ServiceError("请选择文件")
            return None
        limit = max_upload_bytes()
        data = await upload.read(limit + 1)
        if len(data) > limit:
            raise ServiceError(f"上传文件不能超过 {limit // (1024 * 1024)}MB")
        if not data:
            raise ServiceError("文件是空的")
        return data

    @_api
    async def list_terms_route(course_id: int, request: Request):
        user_id, _session = _identity(request)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            current = current_term(db, course)
            return JSONResponse(
                {
                    "current_term_id": current.id if current is not None else None,
                    "terms": [term_public(item, file_count(db, item.id)) for item in list_terms(db, course.id)],
                }
            )

    application.get("/api/courses/{course_id}/terms")(list_terms_route)

    @_api
    async def create_term_route(course_id: int, request: Request):
        user_id, _session = _identity(request)
        body = await _json_body(request)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            settings = without_term_identity(parse_settings(course.settings_json))
            if isinstance(body.get("settings"), dict):
                settings = without_term_identity(body["settings"])
            term = create_term(db, course, settings, make_current=True, clear_identity=True)
            if body:
                basic = dict(settings.get("course_basic_info") or {})
                opened = dict(settings.get("course_open_info") or {})
                for key in ("year_start", "year_end", "semester", "teacher", "school_year_term", "class_name"):
                    if key not in body:
                        continue
                    value = str(body.get(key) or "")
                    if key in {"school_year_term", "class_name"}:
                        basic[key] = value
                    elif key == "teacher":
                        opened["teacher"] = value
                        basic["teacher"] = value
                    else:
                        opened[key] = value
                settings = dict(settings)
                settings["course_basic_info"] = basic
                settings["course_open_info"] = opened
                if "student_count" in body:
                    settings["student_count"] = int(body.get("student_count") or 0)
                if "exam_count" in body:
                    basic["exam_count"] = str(body.get("exam_count") or "")
                    settings["course_basic_info"] = basic
                apply_term_fields(term, settings)
            course.updated_at = utcnow()
            return JSONResponse(_course_dict(db, course))

    application.post("/api/courses/{course_id}/terms")(create_term_route)

    @_api
    async def update_term_route(course_id: int, term_id: int, request: Request):
        user_id, _session = _identity(request)
        body = await _json_body(request)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            term = select_term(db, course, term_id) if body.get("select") else None
            if term is None:
                term = next((item for item in list_terms(db, course.id) if item.id == int(term_id)), None)
            if term is None:
                raise NotFound()
            from web_app.course_versions import freeze_existing_terms, term_curriculum
            freeze_existing_terms(db, course)
            settings = term_curriculum(parse_settings(course.settings_json), term)
            basic = dict(settings.get("course_basic_info") or {})
            opened = dict(settings.get("course_open_info") or {})
            for key in ("year_start", "year_end", "semester", "teacher"):
                if key in body:
                    opened[key] = str(body.get(key) or "")
            if "teacher" in body:
                basic["teacher"] = str(body.get("teacher") or "")
            if "major" in body:
                basic["major"] = str(body.get("major") or "")
            for key in ("school_year_term", "class_name"):
                if key in body:
                    basic[key] = str(body.get(key) or "")
            if "student_count" in body:
                settings["student_count"] = int(body.get("student_count") or 0)
                basic["student_count"] = str(body.get("student_count") or "")
            if "exam_count" in body:
                basic["exam_count"] = str(body.get("exam_count") or "")
            settings["course_basic_info"] = basic
            settings["course_open_info"] = opened
            apply_term_fields(term, settings)
            if term.is_current:
                assign_course_settings(course, settings, term=term)
            course.updated_at = utcnow()
            return JSONResponse(_course_dict(db, course))

    application.patch("/api/courses/{course_id}/terms/{term_id}")(update_term_route)

    @_api
    async def select_term_route(course_id: int, term_id: int, request: Request):
        user_id, _session = _identity(request)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            select_term(db, course, term_id)
            return JSONResponse(_course_dict(db, course))

    application.post("/api/courses/{course_id}/terms/{term_id}/select")(select_term_route)

    @_api
    async def delete_term_route(course_id: int, term_id: int, request: Request):
        user_id, _session = _identity(request)
        body = {}
        if request.headers.get("content-type", "").startswith("application/json"):
            body = await _json_body(request)
        confirm = str(request.query_params.get("confirm") or body.get("confirm") or "") in {"1", "true", "yes"}
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            terms = list_terms(db, course.id)
            if len(terms) <= 1:
                raise ServiceError("至少保留一条学期记录")
            term = next((item for item in terms if item.id == int(term_id)), None)
            if term is None:
                raise NotFound()
            from web_app.db import CourseFile
            from sqlalchemy import select

            files = list(db.scalars(select(CourseFile).where(CourseFile.term_id == term.id)).all())
            if files and not confirm:
                raise ServiceError("该学期已有成绩或报告，确认后才会删除")
            for row in files:
                db.delete(row)
            was_current = bool(term.is_current)
            db.delete(term)
            db.flush()
            if was_current:
                remaining = list_terms(db, course.id)
                if remaining:
                    remaining[0].is_current = 1
            return JSONResponse(_course_dict(db, course))

    application.delete("/api/courses/{course_id}/terms/{term_id}")(delete_term_route)

    @_api
    async def syllabus_extract(request: Request):
        user_id, _session = _identity(request)
        if not syllabus_limiter.allow(f"user:{user_id}"):
            return _json_error(429, "大纲读取过于频繁，请稍后再试")
        form = await _optional_form(request)
        try:
            if form is None:
                raise ServiceError("请上传教学大纲")
            syllabus = _as_upload(form.get("syllabus"))
            register = _as_upload(form.get("register"))
            syllabus_bytes = await _bytes(syllabus, required=True)
            register_bytes = await _bytes(register, required=False) if register else None
            draft = extract_draft(
                syllabus.filename or "syllabus.docx",
                syllabus_bytes,
                register.filename if register else None,
                register_bytes,
            )
            return JSONResponse(draft)
        finally:
            await _close_form(form)

    application.post("/api/syllabus/extract")(syllabus_extract)

    @_api
    async def validate_relation(request: Request):
        _identity(request)
        body = await _json_body(request)
        grid = body.get("relation_grid") or body.get("grid")
        checked = validate_relation_grid(grid)
        return JSONResponse(
            {
                "ok": checked["ok"],
                "errors": checked["errors"],
                "ratios": checked["ratios"],
                "weights": checked["weights"],
            }
        )

    application.post("/api/syllabus/validate-relation")(validate_relation)

    async def _wizard_request(request: Request):
        content_type = request.headers.get("content-type", "")
        if "multipart/form-data" not in content_type:
            return await _json_body(request), None, None
        form = await _optional_form(request)
        try:
            if form is None:
                raise ServiceError("确认内容格式不正确")
            raw = form.get("payload")
            if hasattr(raw, "read"):
                raw = (await raw.read()).decode("utf-8")
            if not isinstance(raw, str) or not raw.strip():
                raise ServiceError("确认内容格式不正确")
            try:
                body = json.loads(raw)
            except json.JSONDecodeError as exc:
                raise ServiceError("确认内容格式不正确") from exc
            upload = _as_upload(form.get("register") or form.get("file"))
            if upload is None:
                return body, None, None
            data = await _bytes(upload, required=True)
            return body, upload.filename or "register.xlsx", data
        finally:
            await _close_form(form)

    @_api
    async def from_syllabus(request: Request):
        user_id, _session = _identity(request)
        body, register_name, register_bytes = await _wizard_request(request)
        settings, extra = settings_from_wizard(body)
        settings = absorb_relation_grid(settings, strict=True)
        from web_app.db import Course

        with session_scope() as db:
            from web_app.course_versions import reject_existing_course
            reject_existing_course(db, user_id, extra["course_name"][:120], settings)
            course = Course(
                user_id=user_id,
                name=extra["course_name"][:120],
                settings_json="{}",
                updated_at=utcnow(),
            )
            db.add(course)
            db.flush()
            assign_course_settings(course, settings)
            from web_app.course_versions import freeze_existing_terms
            freeze_existing_terms(db, course)
            grade_import = None
            grade_import_error = ""
            # 前端不再上传登记表。仍收到文件时走和主页一样的匹配/新建，不先占一条空学期。
            if register_bytes:
                try:
                    parsed = parse_register(register_name or "register.xlsx", register_bytes)
                    grade_import = import_parsed_register(db, course, parsed, confirmed=False, user_id=user_id)
                except ServiceError as exc:
                    grade_import_error = str(exc)
            payload = _course_dict(db, course)
            if grade_import is not None:
                payload["grade_import"] = grade_import
            if grade_import_error:
                payload["grade_import_error"] = grade_import_error
            return JSONResponse(payload)

    application.post("/api/courses/from-syllabus")(from_syllabus)

    @_api
    async def update_syllabus_version(course_id: int, request: Request):
        user_id, _session = _identity(request)
        body = await _json_body(request)
        settings, _extra = settings_from_wizard(body)
        settings = absorb_relation_grid(settings, strict=True)
        from web_app.course_versions import new_syllabus_version
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            new_syllabus_version(db, course, settings, expected_version=body.get("expected_version"))
            course.updated_at = utcnow()
            return JSONResponse(_course_dict(db, course))

    application.post("/api/courses/{course_id}/syllabus-version")(update_syllabus_version)

    def _form_text(form, key: str) -> str:
        if form is None:
            return ""
        raw = form.get(key)
        if raw is None or hasattr(raw, "read"):
            return ""
        return str(raw).strip()

    async def _import_register(course_id: int, request: Request, _term_id: int | None):
        user_id, _session = _identity(request)
        form = await _optional_form(request)
        try:
            if form is None:
                raise ServiceError("请上传成绩登记表")
            upload = _as_upload(form.get("file") or form.get("register"))
            data = await _bytes(upload, required=True)
            parsed = parse_register(upload.filename or "register.xlsx", data)
            confirmed = str(request.query_params.get("confirm") or "") in {"1", "true", "yes"}
            if form.get("confirm") is not None and not hasattr(form.get("confirm"), "read"):
                confirmed = confirmed or str(form.get("confirm")).strip().lower() in {"1", "true", "yes"}
            raw_mode = form.get("mode")
            if hasattr(raw_mode, "read"):
                raw_mode = None
            requested_mode = str(raw_mode or request.query_params.get("mode") or "auto")
            with session_scope() as db:
                course = _owned_course(db, user_id, course_id)
                summary = import_parsed_register(
                    db,
                    course,
                    parsed,
                    confirmed=confirmed,
                    user_id=user_id,
                    mode=requested_mode,
                    year_start=_form_text(form, "year_start"),
                    year_end=_form_text(form, "year_end"),
                    semester=_form_text(form, "semester"),
                    class_name=_form_text(form, "class_name"),
                )
                return JSONResponse(summary)
        finally:
            await _close_form(form)

    @_api
    async def import_current(course_id: int, request: Request):
        return await _import_register(course_id, request, None)

    application.post("/api/courses/{course_id}/grade-register")(import_current)

    @_api
    async def import_term(course_id: int, term_id: int, request: Request):
        return await _import_register(course_id, request, term_id)

    application.post("/api/courses/{course_id}/terms/{term_id}/grade-register")(import_term)
