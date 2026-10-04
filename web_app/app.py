"""CalculatorPro 网页服务。所有页面和接口都要先通过共享密码。"""
from __future__ import annotations

import asyncio
import hmac
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles

from web_app.auth import COOKIE_NAME, SESSION_MAX_AGE, app_password, make_session_token, read_session_token
from web_app.limiter import ai_limiter, login_limiter
from web_app.service import ServiceError, ai_status, build_template, run_ai_report, run_calculation, run_export

STATIC_DIR = Path(__file__).resolve().parent / "static"
MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_REQUEST_BYTES = 12 * 1024 * 1024
PUBLIC_PATHS = {"/healthz", "/login", "/api/login", "/favicon.ico"}


def _client_key(request: Request) -> str:
    host = request.client.host if request.client else ""
    if host in {"127.0.0.1", "::1"}:
        forwarded = request.headers.get("x-forwarded-for", "")
        if forwarded:
            return forwarded.split(",")[0].strip() or host
    return host or "unknown"


def _json_error(status: int, message: str) -> JSONResponse:
    return JSONResponse({"detail": message}, status_code=status)


def _attachment(filename: str, content: bytes, media_type: str) -> Response:
    encoded = quote(filename)
    return Response(
        content=content,
        media_type=media_type,
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded}",
            "Cache-Control": "no-store",
        },
    )


async def _read_upload(upload: UploadFile | None, required: bool) -> bytes | None:
    if upload is None or not upload.filename:
        if required:
            raise ServiceError("请上传 Excel 成绩文件（.xlsx）")
        return None
    name = upload.filename.lower()
    if not name.endswith(".xlsx"):
        raise ServiceError("只接受 .xlsx 文件")
    data = await upload.read(MAX_FILE_BYTES + 1)
    if len(data) > MAX_FILE_BYTES:
        raise ServiceError("上传文件不能超过 10MB")
    if not data.startswith(b"PK"):
        raise ServiceError("文件不是有效的 xlsx")
    return data


def _parse_settings(raw: str) -> dict:
    import json

    if not raw or not raw.strip():
        raise ServiceError("缺少课程设置")
    if len(raw.encode("utf-8")) > 1024 * 1024:
        raise ServiceError("课程设置过大")
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        raise ServiceError("课程设置不是有效的 JSON")
    if not isinstance(data, dict):
        raise ServiceError("课程设置格式不正确")
    return data


class _Guard:
    """ASGI 中间件：除健康检查和登录外，全部要求有效会话。"""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        path = scope.get("path") or "/"
        if path in PUBLIC_PATHS:
            await self.app(scope, receive, send)
            return
        headers = {key.decode("latin1").lower(): value.decode("latin1") for key, value in scope.get("headers", [])}
        length = headers.get("content-length")
        if length and length.isdigit() and int(length) > MAX_REQUEST_BYTES:
            response = _json_error(413, "上传内容不能超过 10MB")
            await response(scope, receive, send)
            return
        cookie_header = headers.get("cookie", "")
        token = ""
        for part in cookie_header.split(";"):
            name, _, value = part.strip().partition("=")
            if name == COOKIE_NAME:
                token = value
                break
        if read_session_token(token):
            await self.app(scope, receive, send)
            return
        if path.startswith("/api/"):
            response = _json_error(401, "未登录或登录已过期")
        else:
            response = RedirectResponse("/login", status_code=303)
        await response(scope, receive, send)


def create_app() -> FastAPI:
    application = FastAPI(title="CalculatorPro", docs_url=None, redoc_url=None, openapi_url=None)

    @application.get("/healthz")
    def healthz():
        return {"status": "ok"}

    @application.get("/favicon.ico")
    def favicon():
        return Response(status_code=204)

    @application.get("/login")
    def login_page(request: Request):
        if read_session_token(request.cookies.get(COOKIE_NAME)):
            return RedirectResponse("/", status_code=303)
        return FileResponse(STATIC_DIR / "login.html", media_type="text/html; charset=utf-8")

    @application.post("/api/login")
    async def login(request: Request):
        if not login_limiter.allow(_client_key(request)):
            return _json_error(429, "尝试次数过多，请稍后再试")
        try:
            body = await request.json()
        except Exception:
            body = {}
        password = ""
        if isinstance(body, dict):
            password = str(body.get("password") or "")
        expected = app_password()
        if not expected:
            return _json_error(503, "服务器未配置 APP_PASSWORD，暂时无法登录")
        if not hmac.compare_digest(password.encode("utf-8"), expected.encode("utf-8")):
            return _json_error(401, "密码不正确")
        response = JSONResponse({"ok": True})
        response.set_cookie(
            COOKIE_NAME,
            make_session_token(),
            max_age=SESSION_MAX_AGE,
            httponly=True,
            samesite="lax",
            path="/",
        )
        return response

    @application.post("/api/logout")
    def logout():
        response = JSONResponse({"ok": True})
        response.delete_cookie(COOKIE_NAME, path="/")
        return response

    @application.get("/")
    def home():
        return FileResponse(STATIC_DIR / "index.html", media_type="text/html; charset=utf-8")

    @application.get("/api/ai-status")
    def ai_status_route():
        return ai_status()

    @application.post("/api/template")
    async def template(settings: str = Form(...)):
        try:
            filename, content = await asyncio.to_thread(build_template, _parse_settings(settings))
        except ServiceError as exc:
            return _json_error(400, str(exc))
        except Exception as exc:
            return _json_error(500, f"模板生成失败：{exc}")
        return _attachment(
            filename,
            content,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @application.post("/api/calculate")
    async def calculate(
        settings: str = Form(...),
        file: UploadFile = File(...),
        previous: UploadFile | None = File(None),
    ):
        try:
            excel = await _read_upload(file, required=True)
            previous_bytes = await _read_upload(previous, required=False)
            parsed = _parse_settings(settings)
            summary = await asyncio.to_thread(run_calculation, excel, previous_bytes, parsed)
        except ServiceError as exc:
            return _json_error(400, str(exc))
        except Exception as exc:
            return _json_error(500, f"计算失败：{exc}")
        return JSONResponse(summary)

    @application.post("/api/export")
    async def export(
        settings: str = Form(...),
        file: UploadFile = File(...),
        previous: UploadFile | None = File(None),
    ):
        try:
            excel = await _read_upload(file, required=True)
            previous_bytes = await _read_upload(previous, required=False)
            parsed = _parse_settings(settings)
            filename, content, _summary = await asyncio.to_thread(run_export, excel, previous_bytes, parsed)
        except ServiceError as exc:
            return _json_error(400, str(exc))
        except Exception as exc:
            return _json_error(500, f"导出失败：{exc}")
        return _attachment(filename, content, "application/zip")

    @application.post("/api/ai-report")
    async def ai_report(
        request: Request,
        settings: str = Form(...),
        file: UploadFile = File(...),
        previous: UploadFile | None = File(None),
    ):
        if not ai_limiter.allow(_client_key(request)):
            return _json_error(429, "AI 报告请求过于频繁，请稍后再试")
        try:
            excel = await _read_upload(file, required=True)
            previous_bytes = await _read_upload(previous, required=False)
            parsed = _parse_settings(settings)
            filename, content = await asyncio.to_thread(run_ai_report, excel, previous_bytes, parsed)
        except ServiceError as exc:
            return _json_error(400, str(exc))
        except Exception as exc:
            return _json_error(500, f"AI 报告生成失败：{exc}")
        return _attachment(filename, content, "application/zip")

    application.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
    application.add_middleware(_Guard)
    return application


app = create_app()
