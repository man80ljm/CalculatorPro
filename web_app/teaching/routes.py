"""Administrator-only pages and APIs for timetable and provider configuration."""
import asyncio,threading
from fastapi import Request
from fastapi.responses import JSONResponse
from sqlalchemy.exc import SQLAlchemyError
from web_app.db import get_engine
from web_app.admin_auth import ADMIN_COOKIE,invitation_csrf_token
from web_app.admin_routes import _page,_require_admin,_require_invite_write,_no_store
from web_app.limiter import RateLimiter
from .api_settings import ApiSettings,ApiError
from .schedule_db import ScheduleDatabase
from . import schedule_deletion
from .pdf_schedule import PdfSchedule
import json

def attach_teaching(application):
    service=None
    service_lock=threading.Lock()
    def get_service():
        nonlocal service
        with service_lock:
            if service is None:
                engine=get_engine();service=PdfSchedule(ScheduleDatabase(engine),ApiSettings(engine))
            return service
    application.state.get_teaching_service=get_service
    tests=RateLimiter(max_calls=5,window_seconds=60)
    uploads=RateLimiter(max_calls=3,window_seconds=60)
    async def body(request,limit):
        content=bytearray()
        async for chunk in request.stream():
            if len(content)+len(chunk)>limit:raise ApiError('请求内容过大。',413)
            content.extend(chunk)
        try:data=json.loads(content or b'{}')
        except (ValueError,UnicodeError):raise ApiError('请求不是有效 JSON。') from None
        if not isinstance(data,dict):raise ApiError('请求内容应为 JSON 对象。')
        return data
    def error(status,message):return _no_store(JSONResponse({'message':message,'detail':message},status_code=status))
    @application.get('/admin/schedule')
    def schedule_page(request:Request):return _page(request,'schedule.html')
    @application.get('/admin/settings')
    def settings_page(request:Request):return _page(request,'settings.html')
    @application.get('/admin/api/settings')
    def settings_get(request:Request):
        denied=_require_admin(request)
        if denied is not None:return denied
        try:
            return _no_store(JSONResponse({'csrf':invitation_csrf_token(request.cookies.get(ADMIN_COOKIE,'')),'settings':get_service().settings.metadata()}))
        except ApiError as exc:return error(exc.status,str(exc))
        except SQLAlchemyError:return error(503,'数据库不可用。')
    @application.get('/admin/api/schedule')
    def schedule_get(request:Request):
        denied=_require_admin(request)
        if denied is not None:return denied
        try:return _no_store(JSONResponse(get_service().db.snapshot()))
        except SQLAlchemyError:return error(503,'数据库不可用。')
    @application.get('/admin/api/schedule/trash')
    def schedule_trash(request:Request):
        denied=_require_admin(request)
        if denied is not None:return denied
        try:return _no_store(JSONResponse(schedule_deletion.list_trash(get_service().db,request.query_params.get('before'))))
        except ApiError as exc:return error(exc.status,str(exc))
        except SQLAlchemyError:return error(503,'数据库不可用。')
    @application.get('/admin/api/schedule/jobs/{identifier}')
    def job_get(identifier:str,request:Request):
        denied=_require_admin(request)
        if denied is not None:return denied
        try:
            if len(identifier)!=32 or any(c not in 'abcdef0123456789' for c in identifier):raise ApiError('识别任务不存在。',404)
            return _no_store(JSONResponse(get_service().job(identifier)))
        except ApiError as exc:return error(exc.status,str(exc))
        except SQLAlchemyError:return error(503,'数据库不可用。')
    @application.post('/admin/api/settings')
    async def settings_save(request:Request):return await write(request,'settings')
    @application.post('/admin/api/settings/test')
    async def settings_test(request:Request):return await write(request,'test')
    @application.post('/admin/api/schedule/analyze')
    async def schedule_analyze(request:Request):return await write(request,'analyze')
    @application.post('/admin/api/schedule/save')
    async def schedule_save(request:Request):return await write(request,'save')
    @application.post('/admin/api/schedule/record')
    async def schedule_edit(request:Request):return await write(request,'record')
    @application.post('/admin/api/schedule/delete-preview')
    async def schedule_delete_preview(request:Request):return await write(request,'delete-preview')
    @application.post('/admin/api/schedule/delete')
    async def schedule_delete(request:Request):return await write(request,'delete')
    @application.post('/admin/api/schedule/restore')
    async def schedule_restore(request:Request):return await write(request,'restore')
    async def write(request,operation):
        denied=_require_invite_write(request)
        if denied is not None:return denied
        try:
            data=await body(request,12*1024*1024 if operation=='analyze' else 4*1024*1024 if operation in {'save','record'} else 1024*1024 if operation in {'delete-preview','delete'} else 16384)
            token=request.cookies.get(ADMIN_COOKIE,'')
            if operation in {'settings','test'} and set(data)-{'source','model','key'}:raise ApiError('只可设置密钥来源、密钥和模型。')
            if operation=='test' and not tests.allow(token):raise ApiError('连接测试过于频繁，请稍后重试。',429)
            if operation=='analyze' and not uploads.allow(token):raise ApiError('PDF 识别提交过于频繁，请稍后重试。',429)
            value=get_service()
            if operation=='settings':payload={'settings':await asyncio.to_thread(value.settings.save,data)}
            elif operation=='test':payload=await asyncio.to_thread(value.settings.test,data)
            elif operation=='analyze':payload=await asyncio.to_thread(value.start,data)
            elif operation=='save':payload=await asyncio.to_thread(value.save,data)
            elif operation=='delete-preview':payload=await asyncio.to_thread(schedule_deletion.preview_delete,value.db,data)
            elif operation=='delete':payload=await asyncio.to_thread(schedule_deletion.delete,value.db,data)
            elif operation=='restore':payload=await asyncio.to_thread(schedule_deletion.restore,value.db,data)
            else:payload=await asyncio.to_thread(value.edit,data)
            return _no_store(JSONResponse(payload,status_code=202 if operation=='analyze' else 200))
        except ApiError as exc:return error(exc.status,str(exc))
        except SQLAlchemyError:return error(503,'数据库不可用，操作未完成，请刷新后核对。')