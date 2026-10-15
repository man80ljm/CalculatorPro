"""Narrow teacher/class lookup for an authenticated user's own course."""
from web_app.teaching.request_origin import same_origin
from fastapi import Request
from fastapi.responses import JSONResponse
from web_app.db import session_scope
from . import course_assignments as assignments, metadata


def attach_assignments(application):
    from web_app.app import _api, _identity, _owned_course, _json_body

    @_api
    async def read(course_id: int, request: Request):
        user_id, _ = _identity(request)
        period = assignments.academic_period(request.query_params.get("year_start"),
            request.query_params.get("year_end"), request.query_params.get("semester"))
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            result = metadata.resolve(db, user_id, assignments.course_name(course), period, college=assignments.course_college(course), course_id=course.id)
        return JSONResponse(result, headers={"Cache-Control": "no-store"})

    @_api
    async def select_teacher(course_id: int, request: Request):
        user_id, _ = _identity(request)
        origin = request.headers.get("origin")
        if not same_origin(request):
            raise assignments.AssignmentError("请从本站页面选择任课教师。", status=403)
        body = await _json_body(request)
        period = assignments.academic_period(body.get("year_start"), body.get("year_end"), body.get("semester"))
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            name = assignments.course_name(course)
            current = metadata.resolve(db,user_id,name,period,college=assignments.course_college(course),course_id=course.id)
            assignments.choose(db, user_id, name, period, body.get("teacher"), college=assignments.course_college(course),
                               matched_course_name=current["matched_course_name"])
            result = metadata.resolve(db,user_id,assignments.course_name(course),period,college=assignments.course_college(course),course_id=course.id)
        return JSONResponse(result, headers={"Cache-Control": "no-store"})

    @_api
    async def save_metadata(course_id: int, request: Request):
        user_id, _ = _identity(request)
        if not same_origin(request):
            raise assignments.AssignmentError("请从本站页面保存授课信息。", status=403)
        body = await _json_body(request)
        period = assignments.academic_period(body.get("year_start"), body.get("year_end"), body.get("semester"))
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            result = metadata.save(db, user_id, assignments.course_name(course), period, body,
                college=assignments.course_college(course), course_id=course.id)
        return JSONResponse(result, headers={"Cache-Control": "no-store"})

    @_api
    async def teachers(request: Request):
        _identity(request)
        with session_scope() as db:
            rows = metadata.search_teachers(db, request.query_params.get("q", ""),request.query_params.get("college", "")[:160])
        return JSONResponse({"teachers": rows}, headers={"Cache-Control": "no-store"})

    application.get("/api/teaching/teachers")(teachers)
    application.patch("/api/courses/{course_id}/teaching-metadata")(save_metadata)
    application.get("/api/courses/{course_id}/teaching-assignment")(read)
    application.post("/api/courses/{course_id}/teaching-assignment")(select_teacher)
