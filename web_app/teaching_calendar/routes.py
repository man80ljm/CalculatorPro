"""教师登录保护下的教学日历直接下载和底层兼容 API。"""
from web_app.teaching.request_origin import same_origin
import asyncio
import json

from fastapi import Request
from starlette.datastructures import UploadFile
from starlette.exceptions import HTTPException
from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse

from web_app.limiter import calendar_ai_limiter
from web_app.db import session_scope
from web_app.terms import parse_settings
from .errors import CalendarError
from .service import MAX_FILE_BYTES, generate, generate_for_course, preview


def attach_calendar(application):
    from web_app.app import _api, _attachment, _identity, _close_form, _owned_course

    async def form_data(request):
        if not request.headers.get("content-type", "").lower().startswith("multipart/form-data"):
            raise CalendarError("请上传文件并提交预览内容。", code="invalid_form")
        # Cookie 写操作仅允许同源；不要求浏览器以外的测试客户端提供 Origin。
        origin = request.headers.get("origin")
        if not same_origin(request):
            raise CalendarError("请从本站页面提交文件。", 403, code="invalid_origin")
        try:
            return await request.form(max_files=2, max_fields=3, max_part_size=1024 * 1024)
        except HTTPException:
            raise CalendarError("上传表单无效或内容过大，请检查文件和预览内容。", 400,
                                code="invalid_form") from None

    async def read_upload(form, name):
        upload = form.get(name)
        if not isinstance(upload, UploadFile) or not upload.filename:
            raise CalendarError("请选择课程教学大纲和教学日历模板。", code="file_missing")
        data = await upload.read(MAX_FILE_BYTES + 1)
        if len(data) > MAX_FILE_BYTES:
            raise CalendarError("每个上传文件不能超过 5MB。", 413, code="file_too_large")
        return upload.filename, data

    @application.get("/teaching-calendar")
    def calendar_page():
        return RedirectResponse("/", status_code=303)

    def calendar_download(filename, content, summary, output_format):
        media = "application/vnd.ms-excel" if output_format == "xls" else (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response = _attachment(filename, content, media)
        response.headers["X-Calendar-Hours-Match"] = json.dumps(summary["hours_match"])
        response.headers["X-Calendar-Total-Hours"] = str(summary["total_hours"])
        if summary["template_total_hours"] is not None:
            response.headers["X-Calendar-Course-Hours"] = str(summary["template_total_hours"])
        term = summary["academic_term"]
        response.headers["X-Calendar-School-Year"] = f"{term['year_start']}-{term['year_end']}"
        response.headers["X-Calendar-Semester"] = term["semester"]
        return response

    @_api
    async def course_calendar(course_id: int, request: Request):
        user_id, _session = _identity(request)
        origin = request.headers.get("origin")
        if not same_origin(request):
            raise CalendarError("请从本站页面下载教学日历。", 403, code="invalid_origin")
        output_format = request.query_params.get("format", "xls")
        if output_format != "xls":
            raise CalendarError("教学日历只提供 .xls 下载。", code="invalid_format")
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            # 下载最新学年学期时使用课程最新大纲，不受正在查看的历史学期影响。
            name, settings = course.name, parse_settings(course.settings_json)
            from web_app.teaching import course_assignments as assignments
            from web_app.teaching import metadata
            teaching = metadata.resolve(db, user_id, assignments.course_name(course), assignments.academic_period(), college=assignments.course_college(course), course_id=course.id)
            if (teaching["status"] == "missing" and teaching["course_matches"]
                    and request.query_params.get("skip_course_match") != "1"):
                raise assignments.AssignmentError(
                    "课表中有名称相近的课程，请先确认对应关系或选择跳过。",
                    status=409, code="teaching_course_match_required",
                    teaching_assignment=teaching)
            assignments.require_choice(teaching)
        if request.query_params.get("progress") == "1":
            from .progress import calendar_events
            return StreamingResponse(calendar_events(name, settings, teaching), media_type="application/x-ndjson",
                headers={"Cache-Control": "no-store", "X-Accel-Buffering": "no"})
        filename, content, summary = await asyncio.to_thread(generate_for_course, name, settings, output_format, teaching_assignment=teaching)
        return calendar_download(filename, content, summary, output_format)

    application.post("/api/courses/{course_id}/teaching-calendar")(course_calendar)

    @_api
    async def preview_route(request: Request):
        user_id, _session = _identity(request)
        form = await form_data(request)
        try:
            raw = form.get("use_ai", "false")
            if raw not in {"true", "false"}:
                raise CalendarError("使用 AI 整理开关格式不正确。", code="invalid_form")
            use_ai = raw == "true"
            if use_ai and not calendar_ai_limiter.allow(f"user:{user_id}"):
                raise CalendarError("AI 整理提交过于频繁，请稍后重试或改用规则模式。", 429,
                                    code="ai_rate_limited", fallback_available=True)
            outline_name, outline_bytes = await read_upload(form, "outline")
            template_name, template_bytes = await read_upload(form, "template")
            result = await asyncio.to_thread(preview, outline_name, outline_bytes, template_name,
                                             template_bytes, use_ai)
            return JSONResponse(result, headers={"Cache-Control": "no-store"})
        finally:
            await _close_form(form)

    @_api
    async def generate_route(request: Request):
        _identity(request)
        form = await form_data(request)
        try:
            template_name, template_bytes = await read_upload(form, "template")
            output_format = form.get("format", "xls")
            payload = form.get("payload", "")
            filename, content, summary = await asyncio.to_thread(generate, template_name, template_bytes,
                                                                 payload, output_format)
            return calendar_download(filename, content, summary, output_format)
        finally:
            await _close_form(form)

    application.post("/api/teaching-calendar/preview")(preview_route)
    application.post("/api/teaching-calendar/generate")(generate_route)
