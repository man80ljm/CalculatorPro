"""教师拥有的原大纲、教案草稿、后台任务和 Word 下载。"""
from web_app.teaching.request_origin import same_origin
import asyncio
import json
import re
import uuid
from datetime import datetime
from pathlib import Path
from sqlalchemy import select
from starlette.datastructures import UploadFile
from fastapi import Request
from fastapi.responses import JSONResponse, FileResponse
from web_app.db import Course, LessonPlan, LessonPlanSource, LessonPlanTask, NotFound, session_scope
from web_app.storage import upload_root
from web_app.teaching_calendar.academic_term import CHINA_TIMEZONE
from web_app.limiter import RateLimiter
from . import source, service, jobs
from web_app.teaching import course_assignments as assignments, metadata
from .errors import LessonPlanError

limiter=RateLimiter(6,60)


def attach_lesson_plans(application):
    from web_app.app import _api, _identity, _owned_course, _json_body, _close_form
    def origin(request):
        value=request.headers.get("origin")
        if not same_origin(request): raise LessonPlanError("请从本站页面提交教案。",403)

    def owned_plan(db,user_id,identity):
        plan=db.get(LessonPlan,identity)
        if plan is None or plan.user_id!=user_id: raise NotFound()
        return plan

    def teaching_for_plan(db, plan, context, teacher_choice=None, metadata_edit=None):
        period=assignments.academic_period(context["学年起"],context["学年止"],context["学期"])
        course=db.get(Course,plan.course_id) if plan.course_id else None
        name=assignments.course_name(course,context["课程名称"]) if course else context["课程名称"]
        college=assignments.course_college(course) if course else ""
        if teacher_choice is not None:
            current=metadata.resolve(db,plan.user_id,name,period,college=college,course_id=plan.course_id,plan_id=plan.id)
            assignments.choose(db,plan.user_id,name,period,teacher_choice,college=college,
                               matched_course_name=current["matched_course_name"])
        if metadata_edit is not None:
            return metadata.save(db,plan.user_id,name,period,metadata_edit,college=college,course_id=plan.course_id,plan_id=plan.id)
        return metadata.resolve(db,plan.user_id,name,period,college=college,course_id=plan.course_id,plan_id=plan.id)

    def synchronize(db,plan):
        context=json.loads(plan.context_json)
        teaching=teaching_for_plan(db,plan,context)
        original=context.copy()
        assignments.fill_context(context,teaching)
        if context!=original:
            plan.context_json=json.dumps(context,ensure_ascii=False); plan.revision+=1
        return teaching

    def template_current(db,plan):
        if not plan.output_path: return False
        task = db.get(LessonPlanTask,Path(plan.output_path).stem)
        return bool(task and task.status=="success" and json.loads(task.checkpoint_json).get("template_digest")==service.template_digest())

    def public(db,plan):
        active=db.scalar(select(LessonPlanTask).where(LessonPlanTask.active_key==plan.id))
        original=json.loads(plan.context_json)
        context=original.copy()
        teaching=teaching_for_plan(db,plan,context)
        assignments.fill_context(context,teaching)
        metadata_current=context==original
        current_template=template_current(db,plan)
        checked=db.scalar(select(LessonPlanTask).where(LessonPlanTask.plan_id==plan.id,LessonPlanTask.status=="success").order_by(LessonPlanTask.updated_at.desc()).limit(1))
        pages_template_current=bool(checked and json.loads(checked.checkpoint_json).get("template_digest")==service.template_digest())
        path=(upload_root()/plan.output_path).resolve()
        ready=current_template and metadata_current and teaching["status"]!="selection_required" and plan.output_revision==plan.revision and plan.output_date==datetime.now(CHINA_TIMEZONE).date().isoformat() and upload_root() in path.parents and path.is_file()
        notices=json.loads(plan.warnings_json)
        if plan.output_path and not current_template:
            notices.append("教案模板已更新，请重新确认生成 Word 后下载；已生成的教案内容会保留。")
        if not metadata_current and plan.output_revision is not None:
            notices.append("后台课表或教师资料已更新，请重新确认生成 Word 后下载。")
        return {"id":plan.id,"course_id":plan.course_id,"source_id":plan.source_id,"revision":plan.revision,"context":context,
                "counts":service.counts(context),"limits":service.engine.LIMITS,"pages":json.loads(plan.pages_json),
                "pages_current":pages_template_current and metadata_current and plan.checked_revision==plan.revision,"warnings":notices,
                "teaching_assignment":teaching,"outline_version":context.get("_outline",{}).get("completion_id",""),
                "job":jobs.task_view(active) if active else None,"download_url":f"/api/lesson-plans/{plan.id}/download" if ready else None}

    async def upload(request):
        origin(request)
        if not request.headers.get("content-type","").startswith("multipart/form-data"): raise LessonPlanError("请上传 Word 大纲。")
        form=await request.form(max_files=1,max_fields=3,max_part_size=source.MAX_BYTES)
        try:
            file=form.get("outline")
            if not isinstance(file,UploadFile) or not file.filename: raise LessonPlanError("请选择 .docx 教学大纲。")
            data=await file.read(source.MAX_BYTES+1)
            if len(data)>source.MAX_BYTES: raise LessonPlanError("教学大纲不能超过 5MB。",413)
            return file.filename,data,form.get("course_id")
        finally: await _close_form(form)

    @_api
    async def store_course_outline(course_id:int,request:Request):
        user_id,_=_identity(request); filename,data,_=await upload(request)
        # 验证与保存只使用自己的模块，不触碰现有建课解析逻辑。
        await asyncio.to_thread(source.validate_bytes,filename,data)
        with session_scope() as db:
            _owned_course(db,user_id,course_id)
            row=source.store_source(db,user_id,course_id,filename,data)
            return JSONResponse({"id":row.id,"filename":row.filename,"saved":True})

    @_api
    async def course_outline(course_id:int,request:Request):
        user_id,_=_identity(request)
        with session_scope() as db:
            _owned_course(db,user_id,course_id)
            row=db.scalar(select(LessonPlanSource).where(LessonPlanSource.user_id==user_id,LessonPlanSource.course_id==course_id).order_by(LessonPlanSource.created_at.desc()))
            if row is None: return JSONResponse({"source":None,"message":"此课程尚未保存原始 Word 大纲，请上传一次后复用。"})
            try:
                result=await asyncio.to_thread(source.public_source,row)
            except LessonPlanError as exc:
                return JSONResponse({"source":None,"message":str(exc)+" 请重新上传符合格式的大纲。"})
            return JSONResponse({"source":result})

    @_api
    async def create(request:Request):
        user_id,_=_identity(request); origin(request)
        if request.headers.get("content-type","").startswith("multipart/form-data"):
            filename,data,course_id=await upload(request)
            if course_id and not str(course_id).isdigit(): raise LessonPlanError("课程编号不正确。")
            course_id=int(course_id) if course_id else None
            outline=await asyncio.to_thread(source.parse_bytes,filename,data)
            with session_scope() as db:
                if course_id is not None: _owned_course(db,user_id,course_id)
                row=source.store_source(db,user_id,course_id,filename,data)
                if not outline["课程名称"]: outline["课程名称"]=Path(filename).stem
                plan=LessonPlan(id=uuid.uuid4().hex,user_id=user_id,course_id=course_id,source_id=row.id,
                                context_json=json.dumps(service.initial_context(outline),ensure_ascii=False))
                db.add(plan); db.flush(); return JSONResponse(public(db,plan))
        body=await _json_body(request)
        with session_scope() as db:
            row=source.owned_source(db,user_id,str(body.get("source_id") or ""))
            if body.get("reuse") is True:
                # 自动打开和多页重复打开都复用已有草稿，保留教师的编辑结果。
                previous=db.scalar(select(LessonPlan).where(LessonPlan.user_id==user_id,LessonPlan.source_id==row.id).order_by(LessonPlan.updated_at.desc()))
                if previous is not None:
                    current_version = (source.saved_snapshot(row) or {}).get("completion_id", "")
                    previous_version = json.loads(previous.context_json).get("_outline", {}).get("completion_id", "")
                    if current_version == previous_version: return JSONResponse(public(db,previous))
            outline=await asyncio.to_thread(source.read_source,row)
            if not outline["课程名称"]: outline["课程名称"]=Path(row.filename).stem
            plan=LessonPlan(id=uuid.uuid4().hex,user_id=user_id,course_id=row.course_id,source_id=row.id,
                            context_json=json.dumps(service.initial_context(outline),ensure_ascii=False))
            db.add(plan); db.flush(); return JSONResponse(public(db,plan))

    @_api
    async def read(identity:str,request:Request):
        user_id,_=_identity(request)
        with session_scope() as db: return JSONResponse(public(db,owned_plan(db,user_id,identity)))

    @_api
    async def latest(course_id:int,request:Request):
        user_id,_=_identity(request)
        with session_scope() as db:
            _owned_course(db,user_id,course_id)
            plan=db.scalar(select(LessonPlan).where(LessonPlan.user_id==user_id,LessonPlan.course_id==course_id).order_by(LessonPlan.updated_at.desc()))
            return JSONResponse({"plan":public(db,plan) if plan else None})

    def revision(plan,body):
        if str(body.get("revision"))!=str(plan.revision): raise LessonPlanError("教案已在其他页面修改，请重新打开后编辑。",409,code="stale_plan")

    @_api
    async def update(identity:str,request:Request):
        user_id,_=_identity(request); origin(request); body=await _json_body(request)
        with session_scope() as db:
            plan=owned_plan(db,user_id,identity); revision(plan,body)
            if db.scalar(select(LessonPlanTask).where(LessonPlanTask.active_key==plan.id)): raise LessonPlanError("后台正在处理此教案，请完成后再编辑。",409)
            original=json.loads(plan.context_json)
            submitted=body.get("context")
            if not isinstance(submitted,dict): raise LessonPlanError("教案内容格式不正确。")
            submitted=dict(submitted)
            submitted.pop("_outline",None)
            if "_outline" in original: submitted["_outline"]=original["_outline"]
            context=service.normalize_context(submitted)
            teaching=teaching_for_plan(db,plan,context,body.get("teacher_choice"),body.get("teaching_metadata"))
            assignments.fill_context(context,teaching)
            if [r["课次"] for r in context["lessons"]]!=[r["课次"] for r in original["lessons"]]: raise LessonPlanError("不能增删或替换大纲课次。")
            for row,old in zip(context["lessons"],original["lessons"]):
                if any(row[k]!=old[k] for k in ("课次","学时","授课题目","授课类型")): raise LessonPlanError("课次、学时、题目和授课类型应保留大纲原文。")
            if context!=original:
                plan.context_json=json.dumps(context,ensure_ascii=False); plan.revision+=1
            db.flush(); return JSONResponse(public(db,plan))

    async def submit(identity,request,action,lesson=None):
        user_id,_=_identity(request); origin(request); body=await _json_body(request)
        with session_scope() as db:
            plan=owned_plan(db,user_id,identity); revision(plan,body)
            active=db.scalar(select(LessonPlanTask).where(LessonPlanTask.active_key==plan.id))
            if not active:
                teaching=synchronize(db,plan)
                assignments.require_choice(teaching)
                if action in {"generate","compress"} and not limiter.allow(f"user:{user_id}"): raise LessonPlanError("AI 教案请求过于频繁，请稍后重试。",429)
            result=jobs.launch(db,user_id,plan,action,lesson)
        jobs.start_runtime()
        return JSONResponse(result,status_code=202)

    @_api
    async def generate(identity:str,request:Request): return await submit(identity,request,"generate")
    @_api
    async def compress(identity:str,lesson:int,request:Request): return await submit(identity,request,"compress",lesson)
    @_api
    async def check(identity:str,request:Request): return await submit(identity,request,"check")
    @_api
    async def export(identity:str,request:Request): return await submit(identity,request,"export")

    @_api
    async def task(identity:str,request:Request):
        user_id,_=_identity(request)
        with session_scope() as db:
            row=db.get(LessonPlanTask,identity)
            if row is None or row.user_id!=user_id: raise NotFound()
            return JSONResponse(jobs.task_view(row))

    @_api
    async def download(identity:str,request:Request):
        user_id,_=_identity(request)
        with session_scope() as db:
            plan=owned_plan(db,user_id,identity)
            context=json.loads(plan.context_json)
            teaching=teaching_for_plan(db,plan,context)
            assignments.require_choice(teaching)
            expected=assignments.fill_context(context.copy(),teaching)
            if context!=expected:
                raise LessonPlanError("后台课表或教师资料已更新，请确认生成 Word 后下载最新版本。",409)
            if plan.output_revision!=plan.revision or plan.output_date!=datetime.now(CHINA_TIMEZONE).date().isoformat():
                raise LessonPlanError("请按当前内容确认生成 Word，再下载当天版本。",409)
            if not template_current(db,plan):
                raise LessonPlanError("教案模板已更新，请重新确认生成 Word 后下载；无需重新调用 AI。",409,code="stale_template")
            root=upload_root(); path=(root/plan.output_path).resolve()
            if root not in path.parents or not path.is_file(): raise LessonPlanError("下载文件已不可用，请重新生成 Word。",404)
            context=json.loads(plan.context_json); name=re.sub(r'[\\/:*?"<>|\r\n]',"_",context["课程名称"])[:80]
            return FileResponse(path,media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",filename=name+"_课程教案.docx")

    @_api
    async def teaching_assignment(identity:str,request:Request):
        user_id,_=_identity(request)
        with session_scope() as db:
            plan=owned_plan(db,user_id,identity)
            context=json.loads(plan.context_json)
            if request.query_params:
                period=assignments.academic_period(request.query_params.get("year_start"),request.query_params.get("year_end"),request.query_params.get("semester"))
                context.update(学年起=period["year_start"],学年止=period["year_end"],学期=period["semester"])
            return JSONResponse(teaching_for_plan(db,plan,context),headers={"Cache-Control":"no-store"})

    application.get("/api/lesson-plans/{identity}/teaching-assignment")(teaching_assignment)
    application.post("/api/courses/{course_id}/lesson-plan-outline")(store_course_outline)
    application.get("/api/courses/{course_id}/lesson-plan-outline")(course_outline)
    application.get("/api/courses/{course_id}/lesson-plans/current")(latest)
    application.post("/api/lesson-plans")(create)
    application.get("/api/lesson-plans/{identity}")(read)
    application.patch("/api/lesson-plans/{identity}")(update)
    application.post("/api/lesson-plans/{identity}/generate")(generate)
    application.post("/api/lesson-plans/{identity}/lessons/{lesson}/compress")(compress)
    application.post("/api/lesson-plans/{identity}/check-pages")(check)
    application.post("/api/lesson-plans/{identity}/export")(export)
    application.get("/api/lesson-plans/{identity}/download")(download)
    application.get("/api/lesson-plan-jobs/{identity}")(task)
