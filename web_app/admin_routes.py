"""只读运维后台。页面和查询都是 GET，写操作只有登录和退出。"""
from __future__ import annotations

import json
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse, Response
from sqlalchemy.exc import SQLAlchemyError

from web_app.admin_auth import (
    ADMIN_COOKIE,
    AdminAuthError,
    admin_configured,
    admin_host,
    authenticate_admin,
    clear_admin_cookie,
    is_admin_path,
    read_admin_session,
    set_admin_cookie,
    start_admin_session,
    normalize_host,
)
from web_app.admin_stats import build_jobs, build_overview, build_users

STATIC_DIR = Path(__file__).resolve().parent / "static" / "admin"
_STATIC_FILES = {
    "admin.css": "text/css; charset=utf-8",
    "admin.js": "text/javascript; charset=utf-8",
}
_PASS_THROUGH = {"/healthz", "/favicon.ico"}


def _json_error(status: int, message: str) -> JSONResponse:
    response = JSONResponse({"detail": message}, status_code=status)
    response.headers["Cache-Control"] = "no-store"
    return response


def _no_store(response: Response) -> Response:
    response.headers["Cache-Control"] = "no-store"
    response.headers["X-Robots-Tag"] = "noindex"
    return response


def _html(name: str) -> FileResponse:
    return FileResponse(
        STATIC_DIR / name,
        media_type="text/html; charset=utf-8",
        headers={"Cache-Control": "no-store", "X-Robots-Tag": "noindex"},
    )


def _page(request: Request, name: str) -> Response:
    if not admin_configured():
        return _html("unconfigured.html")
    logged_in = read_admin_session(request.cookies.get(ADMIN_COOKIE)) is not None
    if name == "login.html":
        if logged_in:
            return RedirectResponse("./", status_code=303)
        return _html("login.html")
    if not logged_in:
        return RedirectResponse("login", status_code=303)
    return _html(name)


def _require_admin(request: Request) -> JSONResponse | None:
    if not admin_configured():
        return _json_error(503, "未配置管理员密码")
    if read_admin_session(request.cookies.get(ADMIN_COOKIE)) is None:
        return _json_error(401, "未登录或登录已过期")
    return None


async def _json_body(request: Request) -> dict:
    raw = await request.body()
    if len(raw) > 65536:
        raise ValueError("请求内容过大")
    if not raw.strip():
        return {}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError("请求不是有效的 JSON") from exc
    if not isinstance(data, dict):
        raise ValueError("请求格式不正确")
    return data


def _scope_host(scope) -> str:
    for key, value in scope.get("headers") or []:
        if key.lower() == b"host":
            return normalize_host(value.decode("latin1"))
    return ""


def _rewrite_admin_host(scope) -> dict:
    """子域根路径转到 /admin。已经是 /admin 前缀的请求保持原样。"""
    if scope.get("type") != "http":
        return scope
    expected = admin_host()
    if not expected or _scope_host(scope) != expected:
        return scope
    path = scope.get("path") or "/"
    if path in _PASS_THROUGH or is_admin_path(path):
        return scope
    new_path = "/admin/" if path == "/" else "/admin" + (path if path.startswith("/") else "/" + path)
    copied = dict(scope)
    copied["path"] = new_path
    copied["raw_path"] = new_path.encode("utf-8")
    return copied


class AdminHostMiddleware:
    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope.get("type") == "http":
            scope = _rewrite_admin_host(scope)
        await self.app(scope, receive, send)


def attach_admin(application: FastAPI) -> None:
    @application.post("/admin/api/login")
    async def admin_login(request: Request):
        if not admin_configured():
            return _json_error(503, "未配置管理员密码")
        try:
            body = await _json_body(request)
        except ValueError as exc:
            return _json_error(400, str(exc))
        username = body.get("username")
        password = body.get("password")
        if not isinstance(username, str):
            username = ""
        if not isinstance(password, str):
            password = ""
        try:
            ok = authenticate_admin(username, password)
        except AdminAuthError as exc:
            return _json_error(exc.status, str(exc))
        if not ok:
            return _json_error(401, "用户名或密码不正确")
        response = JSONResponse({"ok": True})
        set_admin_cookie(response, start_admin_session())
        return _no_store(response)

    @application.post("/admin/api/logout")
    async def admin_logout():
        response = JSONResponse({"ok": True})
        clear_admin_cookie(response)
        return _no_store(response)

    @application.get("/admin/api/overview")
    def admin_overview(request: Request):
        denied = _require_admin(request)
        if denied is not None:
            return denied
        try:
            payload = build_overview()
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")
        return _no_store(JSONResponse(payload))

    @application.get("/admin/api/users")
    def admin_users(request: Request):
        denied = _require_admin(request)
        if denied is not None:
            return denied
        try:
            payload = build_users()
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")
        return _no_store(JSONResponse(payload))

    @application.get("/admin/api/jobs")
    def admin_jobs(request: Request):
        denied = _require_admin(request)
        if denied is not None:
            return denied
        try:
            payload = build_jobs()
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")
        return _no_store(JSONResponse(payload))

    @application.get("/admin/static/{name}")
    def admin_static(name: str):
        media = _STATIC_FILES.get(name)
        if media is None:
            return Response(status_code=404)
        return FileResponse(STATIC_DIR / name, media_type=media, headers={"Cache-Control": "no-cache"})

    @application.get("/admin")
    def admin_root():
        return RedirectResponse("/admin/", status_code=307)

    @application.get("/admin/")
    def admin_home(request: Request):
        return _page(request, "index.html")

    @application.get("/admin/login")
    def admin_login_page(request: Request):
        return _page(request, "login.html")

    @application.get("/admin/users")
    def admin_users_page(request: Request):
        return _page(request, "users.html")

    @application.get("/admin/jobs")
    def admin_jobs_page(request: Request):
        return _page(request, "jobs.html")
