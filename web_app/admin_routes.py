"""运维统计与邀请码管理；邀请码写入沿用独立管理员权限。"""
from __future__ import annotations

import asyncio
import json
import secrets
from urllib.parse import urlsplit
from pathlib import Path

from fastapi import FastAPI, Query, Request
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
    invitation_csrf_token,
)
from web_app.admin_stats import build_jobs, build_overview, build_users
from web_app.invitations import InvitationError, create_admin_invitation, list_invitations, update_invitation, delete_invitation

STATIC_DIR = Path(__file__).resolve().parent / "static" / "admin"
_STATIC_FILES = {
    "admin.css": "text/css; charset=utf-8",
    "admin.js": "text/javascript; charset=utf-8",
    "teaching-ui.js": "text/javascript; charset=utf-8",
    "teaching-api.js": "text/javascript; charset=utf-8",
    "schedule.js": "text/javascript; charset=utf-8",
    "settings.js": "text/javascript; charset=utf-8",
    "teachers.js": "text/javascript; charset=utf-8",
    "teachers.css": "text/css; charset=utf-8",
    "teaching.css": "text/css; charset=utf-8",
    "schedule.css": "text/css; charset=utf-8",
    "settings.css": "text/css; charset=utf-8",
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


def _require_invite_write(request: Request) -> JSONResponse | None:
    denied = _require_admin(request)
    if denied is not None:
        return denied
    origin = request.headers.get("origin")
    if origin:
        try:
            parsed = urlsplit(origin)
        except ValueError:
            return _json_error(403, "请求来源不允许")
        if parsed.scheme not in {"http", "https"} or parsed.netloc.lower() != request.headers.get("host", "").lower():
            return _json_error(403, "请求来源不允许")
    if request.headers.get("sec-fetch-site") == "cross-site":
        return _json_error(403, "请求来源不允许")
    expected = invitation_csrf_token(request.cookies.get(ADMIN_COOKIE, ""))
    provided = request.headers.get("x-admin-csrf", "")
    if not secrets.compare_digest(expected.encode(), provided.encode()):
        return _json_error(403, "页面已失效，请刷新后台后重试")
    if request.method != "DELETE" and request.headers.get("content-type", "").split(";", 1)[0].lower() != "application/json":
        return _json_error(400, "请使用 JSON 请求")
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
    def admin_jobs(request: Request, user_id: int | None = Query(None, ge=1, le=2147483647),
                   page: int = Query(1, ge=1, le=100000000)):
        denied = _require_admin(request)
        if denied is not None:
            return denied
        try:
            payload = build_jobs(user_id=user_id, page=page)
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")
        return _no_store(JSONResponse(payload))

    @application.get("/admin/api/invites")
    def admin_invites(request: Request):
        denied = _require_admin(request)
        if denied is not None:
            return denied
        try:
            payload = {"invites": list_invitations(),
                       "csrf_token": invitation_csrf_token(request.cookies.get(ADMIN_COOKIE, ""))}
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")
        return _no_store(JSONResponse(payload))

    @application.post("/admin/api/invites")
    async def admin_create_invite(request: Request):
        denied = _require_invite_write(request)
        if denied is not None:
            return denied
        try:
            body = await _json_body(request)
            if set(body) - {"max_uses"}:
                raise InvitationError("只可设置可注册总次数")
            invitation = await asyncio.to_thread(create_admin_invitation, body.get("max_uses", 5))
            return _no_store(JSONResponse({"invitation": invitation}, status_code=201))
        except (InvitationError, ValueError) as exc:
            return _json_error(getattr(exc, "status", 400), str(exc))
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")

    @application.patch("/admin/api/invites/{invitation_id}")
    async def admin_update_invite(invitation_id: int, request: Request):
        denied = _require_invite_write(request)
        if denied is not None:
            return denied
        try:
            body = await _json_body(request)
            invitation = await asyncio.to_thread(update_invitation, invitation_id, body)
            return _no_store(JSONResponse({"invitation": invitation}))
        except (InvitationError, ValueError) as exc:
            return _json_error(getattr(exc, "status", 400), str(exc))
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")

    @application.delete("/admin/api/invites/{invitation_id}")
    async def admin_delete_invite(invitation_id: int, request: Request):
        denied = _require_invite_write(request)
        if denied is not None:
            return denied
        try:
            await asyncio.to_thread(delete_invitation, invitation_id)
            return _no_store(JSONResponse({"ok": True}))
        except InvitationError as exc:
            return _json_error(exc.status, str(exc))
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")

    @application.get("/admin/invites")
    def admin_invites_page(request: Request):
        return _page(request, "invites.html")

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
