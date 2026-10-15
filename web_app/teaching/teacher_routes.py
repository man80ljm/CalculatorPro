"""Staff management is administrator-only; contact data never reaches user APIs."""
import asyncio
import json
from fastapi import Request, Query
from fastapi.responses import JSONResponse
from sqlalchemy.exc import SQLAlchemyError
from web_app.admin_auth import ADMIN_COOKIE, invitation_csrf_token
from web_app.admin_routes import _page, _require_admin, _require_invite_write, _no_store
from web_app.db import get_engine
from web_app.limiter import RateLimiter
from web_app.service import ServiceError
from .api_settings import ApiError
from .teacher_directory import TeacherDirectory

upload_limiter=RateLimiter(3,60)


def attach_teacher_directory(application):
    def store(): return TeacherDirectory(get_engine())
    def response(value,status=200): return _no_store(JSONResponse(value,status_code=status))
    def failure(error): return response({"detail":str(error)},getattr(error,"status",400) or 400)

    async def body(request,limit):
        content=bytearray()
        async for chunk in request.stream():
            if len(content)+len(chunk)>limit: raise ApiError("请求内容过大。",413)
            content.extend(chunk)
        try: result=json.loads(content or b"{}")
        except (ValueError,UnicodeError): raise ApiError("请求不是有效的 JSON。") from None
        if not isinstance(result,dict): raise ApiError("请求内容应为 JSON 对象。")
        return result

    @application.get("/admin/teachers")
    def page(request:Request): return _page(request,"teachers.html")

    @application.get("/admin/api/teachers")
    async def listing(request:Request,query:str="",department:str="",status:str="active",page:int=Query(1,ge=1)):
        denied=_require_admin(request)
        if denied is not None: return denied
        try:
            value=await asyncio.to_thread(store().snapshot,query,department,status,page)
            value["csrf"]=invitation_csrf_token(request.cookies.get(ADMIN_COOKIE,""))
            return response(value)
        except (ApiError,ServiceError) as error: return failure(error)
        except SQLAlchemyError: return response({"detail":"教师库暂不可用，请稍后重试。"},503)

    @application.get("/admin/api/teachers/import/{identity}")
    async def read_import(identity:str,request:Request):
        denied=_require_admin(request)
        if denied is not None: return denied
        try: return response(await asyncio.to_thread(store().read_import,identity))
        except ApiError as error: return failure(error)
        except SQLAlchemyError: return response({"detail":"教师库暂不可用，请稍后重试。"},503)

    async def write(request,operation):
        denied=_require_invite_write(request)
        if denied is not None: return denied
        try:
            data=await body(request,7*1024*1024 if operation=="upload" else 65536)
            value=store()
            if operation=="upload":
                if set(data)-{"filename","pdf"}: raise ApiError("请提交教师联系表 PDF。")
                if not upload_limiter.allow(request.cookies.get(ADMIN_COOKIE,"")): raise ApiError("导入提交过于频繁，请稍后重试。",429)
                result=await asyncio.to_thread(value.upload,data.get("filename"),data.get("pdf"))
            elif operation=="review": result=await asyncio.to_thread(value.review,data)
            elif operation=="confirm": result=await asyncio.to_thread(value.confirm,data)
            elif operation=="record": result=await asyncio.to_thread(value.save_record,data)
            else: result=await asyncio.to_thread(value.set_status,data)
            return response(result)
        except (ApiError,ServiceError) as error: return failure(error)
        except SQLAlchemyError: return response({"detail":"教师库暂不可用，本次修改未保存。"},503)

    @application.post("/admin/api/teachers/import/preview")
    async def upload(request:Request): return await write(request,"upload")
    @application.post("/admin/api/teachers/import/review")
    async def review(request:Request): return await write(request,"review")
    @application.post("/admin/api/teachers/import/confirm")
    async def confirm(request:Request): return await write(request,"confirm")
    @application.post("/admin/api/teachers/record")
    async def record(request:Request): return await write(request,"record")
    @application.post("/admin/api/teachers/status")
    async def status(request:Request): return await write(request,"status")
