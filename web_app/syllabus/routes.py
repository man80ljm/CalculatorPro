"""Authenticated background syllabus imports and progress, including saved original DOCX reuse."""
import asyncio
from pathlib import Path

from fastapi import Request
from fastapi.responses import JSONResponse
from starlette.datastructures import UploadFile
from sqlalchemy import select
from web_app.db import LessonPlanSource, SyllabusImportTask, NotFound, session_scope
from web_app.service import ServiceError
from web_app.limiter import RateLimiter
from web_app.teaching.request_origin import same_origin
from web_app.lesson_plans.source import source_path
from . import compatibility, jobs

limiter = RateLimiter(6,60)


def attach_syllabus_imports(application):
    from web_app.app import _api, _identity, _owned_course, _json_body, _close_form

    @_api
    async def start(request: Request):
        user_id, _ = _identity(request)
        if not same_origin(request):
            raise ServiceError("请从本站页面读取大纲。",status=403)
        if not limiter.allow(f"user:{user_id}"):
            raise ServiceError("大纲补全请求过于频繁，请稍后再试。",status=429)
        course_id = None
        if request.headers.get("content-type","").startswith("multipart/form-data"):
            form = await request.form(max_files=1,max_fields=2,max_part_size=compatibility.MAX_BYTES)
            try:
                upload = form.get("syllabus")
                if not isinstance(upload,UploadFile) or not upload.filename:
                    raise ServiceError("请先选择教学大纲。")
                filename = upload.filename
                data = await upload.read(compatibility.MAX_BYTES+1)
                raw = form.get("allow_completion","true")
                if raw not in {"true","false"}:
                    raise ServiceError("AI 补全选项不正确。")
                allow_completion = raw == "true"
            finally:
                await _close_form(form)
        else:
            body = await _json_body(request)
            course_id = body.get("course_id")
            if type(course_id) is not int:
                raise ServiceError("请选择需要补全的课程。")
            allow_completion = body.get("allow_completion",True)
            if type(allow_completion) is not bool:
                raise ServiceError("AI 补全选项不正确。")
            with session_scope() as db:
                _owned_course(db,user_id,course_id)
                row = db.scalar(select(LessonPlanSource).where(LessonPlanSource.user_id==user_id,LessonPlanSource.course_id==course_id)
                    .order_by(LessonPlanSource.created_at.desc()).limit(1))
                if row is None:
                    raise ServiceError("此课程尚未保存原始大纲，请上传一次后再补全。")
                filename, data = row.filename, await asyncio.to_thread(source_path(row).read_bytes)
        with session_scope() as db:
            result = jobs.launch(db,user_id,filename,data,allow_completion,course_id)
        jobs.start_runtime()
        return JSONResponse(result,status_code=202,headers={"Cache-Control":"no-store"})

    @_api
    async def read(identity: str, request: Request):
        user_id, _ = _identity(request)
        with session_scope() as db:
            task = db.get(SyllabusImportTask,identity)
            if task is None or task.user_id != user_id:
                raise NotFound()
            result = jobs.view(task)
        return JSONResponse(result,headers={"Cache-Control":"no-store"})

    @_api
    async def validate_draft(request: Request):
        user_id, _ = _identity(request)
        body = await _json_body(request)
        with session_scope() as db:
            task = db.get(SyllabusImportTask,str(body.get("completion_job_id") or ""))
            if task is None or task.user_id != user_id:
                raise NotFound()
            if task.status != "success":
                raise ServiceError("大纲尚未读取完成。",status=409)
            import json
            outline = json.loads(task.result_json)["outline"]
        from . import completion
        grid = body.get("relation_grid")
        if not isinstance(grid,list) or len(grid)>101 or any(not isinstance(row,list) or len(row)>24 or any(not isinstance(cell,str) or len(cell)>500 for cell in row) for row in grid):
            raise ServiceError("考核对应关系表格式不正确。")
        if not isinstance(body.get("fields"),dict) or not isinstance(body.get("objectives"),list):
            raise ServiceError("课程信息或目标格式不正确。")
        relation = compatibility.validate_relation_grid(grid)
        raw_grad = body.get("grad_req_map") or []
        if not isinstance(raw_grad,list) or any(not isinstance(item,dict) for item in raw_grad):
            raise ServiceError("毕业要求对应关系格式不正确。")
        grad = [{key:completion.text(item.get(key,""),key,1 if key=="strength" else 1000) for key in ("requirement","indicator","strength")} for item in raw_grad]
        draft = {"fields":body.get("fields") or {},"objectives":body.get("objectives") or [],
                 "grad_req_map":grad,"relation":relation}
        errors = completion.draft_errors(draft,outline)
        return JSONResponse({"ok":not errors,"errors":errors,"relation_errors":relation["errors"],
                             "ratios":relation["ratios"],"weights":relation["weights"]})

    application.post("/api/syllabus/validate-draft")(validate_draft)
    application.post("/api/syllabus/imports")(start)
    application.get("/api/syllabus/imports/{identity}")(read)
