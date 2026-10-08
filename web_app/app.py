"""CalculatorPro 网页服务。每位教师使用自己的账号，课程资料按人隔离。"""
from __future__ import annotations

import asyncio
import functools
import ipaddress
import json
import os
from pathlib import Path
from urllib.parse import quote, unquote

from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from sqlalchemy import select
from sqlalchemy.exc import SQLAlchemyError
from starlette.datastructures import UploadFile

from web_app.admin_auth import is_admin_path
from web_app.auth import (
    revoke_other_sessions,
    COOKIE_NAME,
    SESSION_MAX_AGE,
    AuthError,
    authenticate,
    change_password,
    cookie_secure,
    get_secret_key,
    normalize_username,
    read_session_token,
    register_user,
    revoke_session,
    session_is_active,
    start_session,
)
from web_app.db import Course, CourseFile, NotFound, Term, User, init_db, ping_db, session_scope, utcnow
from web_app.terms import (
    apply_term_fields,
    assign_course_settings,
    create_term,
    current_term,
    file_count,
    list_terms,
    merge_settings,
    term_public,
)
from web_app.limiter import ai_limiter, allow_login, register_limiter
from web_app.previous_attainment import store_last_achievement
from web_app.relation_grid import absorb_relation_grid
from web_app.report_jobs import launch_report_job, report_job_download, report_job_snapshot
from web_app.service import ServiceError, ai_status, build_template, run_ai_report, run_calculation, run_export
from web_app.storage import (
    download_filename,
    ensure_upload_root,
    read_blob,
    remove_tree,
    save_blob,
)

STATIC_DIR = Path(__file__).resolve().parent / "static"
PUBLIC_PATHS = {"/healthz", "/login", "/register", "/api/login", "/api/register", "/favicon.ico"}
XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def max_upload_bytes() -> int:
    raw = os.environ.get("MAX_UPLOAD_MB", "10").strip() or "10"
    try:
        mb = int(raw)
    except ValueError:
        mb = 10
    mb = min(max(mb, 1), 50)
    return mb * 1024 * 1024


def max_request_bytes() -> int:
    # 大纲和成绩登记表可以同时上传，各不超过单个文件上限。
    return 2 * max_upload_bytes() + 2 * 1024 * 1024


def _trusted_proxies() -> list:
    raw = os.environ.get("TRUSTED_PROXIES", "127.0.0.1/32,::1/128,172.16.0.0/12")
    nets = []
    for part in raw.split(","):
        part = part.strip()
        if part:
            try:
                nets.append(ipaddress.ip_network(part, strict=False))
            except ValueError:
                continue
    return nets


def _client_key(request: Request) -> str:
    """限流用的客户端地址。只有直连方是受信代理时才看 X-Forwarded-For，且取最右一项（代理追加的那一项）。"""
    host = request.client.host if request.client else ""
    try:
        addr = ipaddress.ip_address(host)
    except ValueError:
        return host or "unknown"
    if any(addr in net for net in _trusted_proxies()):
        forwarded = request.headers.get("x-forwarded-for", "")
        parts = [item.strip() for item in forwarded.split(",") if item.strip()]
        if parts:
            return parts[-1]
        real_ip = request.headers.get("x-real-ip", "").strip()
        if real_ip:
            return real_ip
    return host


def _json_error(status: int, message: str, **extra) -> JSONResponse:
    body = {"detail": message}
    for key, value in extra.items():
        if value is not None:
            body[key] = value
    return JSONResponse(body, status_code=status)


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


def _media_type(filename: str) -> str:
    lower = filename.lower()
    if lower.endswith(".xlsx"):
        return XLSX_MEDIA
    if lower.endswith(".docx"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if lower.endswith(".zip"):
        return "application/zip"
    return "application/octet-stream"


def _set_session_cookie(response: Response, token: str) -> None:
    response.set_cookie(
        COOKIE_NAME,
        token,
        max_age=SESSION_MAX_AGE,
        httponly=True,
        samesite="lax",
        secure=cookie_secure(),
        path="/",
    )


def _clear_session_cookie(response: Response) -> None:
    response.delete_cookie(COOKIE_NAME, path="/", httponly=True, samesite="lax", secure=cookie_secure())


def _cookie_from_scope(scope) -> str:
    headers = {key.decode("latin1").lower(): value.decode("latin1") for key, value in scope.get("headers", [])}
    token = ""
    for part in headers.get("cookie", "").split(";"):
        name, _, value = part.strip().partition("=")
        if name == COOKIE_NAME and value:
            token = unquote(value)
    return token


def _content_length(scope) -> int | None:
    headers = {key.decode("latin1").lower(): value.decode("latin1") for key, value in scope.get("headers", [])}
    length = headers.get("content-length")
    if length and length.isdigit():
        return int(length)
    return None


class _Guard:
    """除健康检查、登录和注册外，请求必须带仍有效的会话。"""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        path = scope.get("path") or "/"
        length = _content_length(scope)
        if length is not None and length > max_request_bytes():
            response = _json_error(413, f"上传内容不能超过 {max_upload_bytes() // (1024 * 1024)}MB")
            await response(scope, receive, send)
            return
        # 运维后台走独立 Cookie，不接受教师会话，也不把未登录的人送到教师登录页。
        if path in PUBLIC_PATHS or is_admin_path(path):
            await self.app(scope, receive, send)
            return
        parsed = read_session_token(_cookie_from_scope(scope))
        active = False
        if parsed:
            try:
                active = session_is_active(parsed[0], parsed[1])
            except SQLAlchemyError:
                response = _json_error(503, "数据库不可用")
                await response(scope, receive, send)
                return
        if active and parsed:
            state = scope.setdefault("state", {})
            if isinstance(state, dict):
                state["user_id"] = parsed[0]
                state["session_id"] = parsed[1]
            await self.app(scope, receive, send)
            return
        if path.startswith("/api/"):
            response = _json_error(401, "未登录或登录已过期")
        else:
            response = RedirectResponse("/login", status_code=303)
        await response(scope, receive, send)


def _api(fn):
    @functools.wraps(fn)
    async def wrapper(*args, **kwargs):
        try:
            return await fn(*args, **kwargs)
        except NotFound:
            return _json_error(404, "没有找到")
        except AuthError as exc:
            return _json_error(exc.status, str(exc))
        except ServiceError as exc:
            return _json_error(getattr(exc, "status", 400) or 400, str(exc), **(getattr(exc, "extra", None) or {}))
        except ValueError as exc:
            return _json_error(400, str(exc))
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")

    return wrapper


def _identity(request: Request) -> tuple[int, str]:
    user_id = getattr(request.state, "user_id", None)
    session_id = getattr(request.state, "session_id", None)
    if not user_id or not session_id:
        raise AuthError("未登录或登录已过期", status=401)
    return int(user_id), str(session_id)


def _parse_settings(raw: str) -> dict:
    if not raw or not raw.strip():
        raise ServiceError("缺少课程设置")
    if len(raw.encode("utf-8")) > 1024 * 1024:
        raise ServiceError("课程设置过大")
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ServiceError("课程设置不是有效的 JSON") from exc
    if not isinstance(data, dict):
        raise ServiceError("课程设置格式不正确")
    return data


async def _read_upload(upload: UploadFile | None, required: bool) -> bytes | None:
    if upload is None or not upload.filename:
        if required:
            raise ServiceError("请上传 Excel 成绩文件（.xlsx）")
        return None
    if not upload.filename.lower().endswith(".xlsx"):
        raise ServiceError("只接受 .xlsx 文件")
    limit = max_upload_bytes()
    data = await upload.read(limit + 1)
    if len(data) > limit:
        raise ServiceError(f"上传文件不能超过 {limit // (1024 * 1024)}MB")
    if not data.startswith(b"PK"):
        raise ServiceError("文件不是有效的 xlsx")
    return data


async def _optional_form(request: Request):
    content_type = request.headers.get("content-type", "")
    if "multipart/form-data" not in content_type and "application/x-www-form-urlencoded" not in content_type:
        return None
    return await request.form()


async def _close_form(form) -> None:
    if form is None:
        return
    for value in form.values():
        if isinstance(value, UploadFile):
            await value.close()


def _as_upload(value) -> UploadFile | None:
    if isinstance(value, UploadFile) and value.filename:
        return value
    return None


async def _json_body(request: Request) -> dict:
    raw = await request.body()
    if len(raw) > 1024 * 1024:
        raise ServiceError("请求内容过大")
    if not raw.strip():
        return {}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ServiceError("请求不是有效的 JSON") from exc
    if not isinstance(data, dict):
        raise ServiceError("请求格式不正确")
    return data


def _course_name_ok(name: str) -> str:
    cleaned = (name or "").strip()
    if not cleaned or len(cleaned) > 120:
        raise ServiceError("请填写课程名称（不超过 120 字）")
    return cleaned


def _settings_of(course: Course) -> dict:
    try:
        data = json.loads(course.settings_json or "{}")
    except json.JSONDecodeError:
        return {}
    return data if isinstance(data, dict) else {}


def _file_dict(row: CourseFile) -> dict:
    created = row.created_at.isoformat(timespec="seconds") if row.created_at else None
    return {
        "id": row.id,
        "course_id": row.course_id,
        "term_id": row.term_id,
        "kind": row.kind,
        "original_name": row.original_name,
        "size": row.size,
        "created_at": created,
    }


def _course_files(db, user_id: int, course_id: int, term_id: int | None = None) -> list[CourseFile]:
    query = select(CourseFile).where(CourseFile.user_id == user_id, CourseFile.course_id == course_id)
    if term_id is not None:
        query = query.where(CourseFile.term_id == int(term_id))
    return list(db.scalars(query.order_by(CourseFile.id.desc())).all())


def _term_text_requested(settings: dict) -> bool:
    """只有老师确实填了学期身份时才新建学期。上课人数的默认值不能单独触发。"""
    basic = settings.get("course_basic_info") if isinstance(settings.get("course_basic_info"), dict) else {}
    opened = settings.get("course_open_info") if isinstance(settings.get("course_open_info"), dict) else {}
    for key in ("teacher", "class_name", "major", "school_year_term"):
        if str(basic.get(key) or "").strip():
            return True
    for key in ("year_start", "year_end", "semester", "term", "teacher"):
        if str(opened.get(key) or "").strip():
            return True
    return False


def _course_dict(db, course: Course) -> dict:
    updated = course.updated_at.isoformat(timespec="seconds") if course.updated_at else None
    term = current_term(db, course)
    terms = []
    for item in list_terms(db, course.id):
        terms.append(term_public(item, file_count(db, item.id)))
    return {
        "id": course.id,
        "name": course.name,
        "settings": _settings_of(course),
        "current_term_id": term.id if term is not None else None,
        "current_term": term_public(term, file_count(db, term.id)) if term is not None else None,
        "terms": terms,
        "files": [_file_dict(row) for row in _course_files(db, course.user_id, course.id, term.id)] if term is not None else [],
        "updated_at": updated,
    }


def _owned_course(db, user_id: int, course_id: int) -> Course:
    course = db.scalar(select(Course).where(Course.id == int(course_id), Course.user_id == int(user_id)))
    if course is None:
        raise NotFound()
    return course


def _owned_file(db, user_id: int, course_id: int, file_id: int) -> CourseFile:
    row = db.scalar(
        select(CourseFile).where(
            CourseFile.id == int(file_id),
            CourseFile.course_id == int(course_id),
            CourseFile.user_id == int(user_id),
        )
    )
    if row is None:
        raise NotFound()
    return row


def _latest_bytes(db, user_id: int, course_id: int, kind: str, term_id: int | None = None) -> bytes | None:
    query = select(CourseFile).where(
        CourseFile.user_id == user_id, CourseFile.course_id == course_id, CourseFile.kind == kind
    )
    if term_id is not None:
        query = query.where(CourseFile.term_id == int(term_id))
    row = db.scalar(query.order_by(CourseFile.id.desc()))
    if row is None:
        return None
    try:
        return read_blob(row)
    except NotFound as exc:
        raise ServiceError("已保存的文件丢失，请重新上传") from exc


def _store_settings(course: Course, settings: dict) -> dict:
    prepared = absorb_relation_grid(settings, strict=True)
    assign_course_settings(course, prepared)
    course.updated_at = utcnow()
    return prepared


def _job_settings(settings: dict, term) -> dict:
    merged = merge_settings(settings, term)
    merged["term_id"] = term.id
    return merged


def _persist_files(
    user_id: int, course_id: int, files: list[tuple[str, bytes]], kind: str, term_id: int | None = None
) -> None:
    if not files:
        return
    with session_scope() as db:
        course = _owned_course(db, user_id, course_id)
        if term_id:
            saved_term = int(term_id)
        else:
            current = current_term(db, course)
            if current is None:
                raise ServiceError("请先导入本学期成绩登记表")
            saved_term = current.id
        for name, data in files:
            save_blob(db, user_id, course_id, name, data, kind, term_id=saved_term)


def _browser_logged_in(request: Request) -> bool:
    parsed = read_session_token(request.cookies.get(COOKIE_NAME))
    if not parsed:
        return False
    try:
        return session_is_active(parsed[0], parsed[1])
    except SQLAlchemyError:
        return False


def create_app() -> FastAPI:
    get_secret_key()
    ensure_upload_root()
    init_db()
    application = FastAPI(title="CalculatorPro", docs_url=None, redoc_url=None, openapi_url=None)

    @application.get("/healthz")
    def healthz():
        try:
            ping_db()
        except Exception:
            return JSONResponse({"status": "error"}, status_code=503)
        return {"status": "ok"}

    @application.get("/favicon.ico")
    def favicon():
        return Response(status_code=204)

    @application.get("/login")
    @application.get("/register")
    def login_page(request: Request):
        if _browser_logged_in(request):
            return RedirectResponse("/", status_code=303)
        return FileResponse(STATIC_DIR / "login.html", media_type="text/html; charset=utf-8")

    @application.post("/api/register")
    async def register(request: Request):
        if not register_limiter.allow(_client_key(request)):
            return _json_error(429, "尝试次数过多，请稍后再试")
        try:
            body = await _json_body(request)
            register_user(str(body.get("username") or ""), str(body.get("password") or ""))
        except AuthError as exc:
            return _json_error(exc.status, str(exc))
        except ServiceError as exc:
            return _json_error(400, str(exc))
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")
        return JSONResponse({"ok": True})

    @application.post("/api/login")
    async def login(request: Request):
        try:
            body = await _json_body(request)
        except ServiceError:
            body = {}
        username = str(body.get("username") or "")
        if not allow_login(_client_key(request), normalize_username(username, strict=False) or username):
            return _json_error(429, "尝试次数过多，请稍后再试")
        try:
            user = authenticate(username, str(body.get("password") or ""))
        except SQLAlchemyError:
            return _json_error(503, "数据库不可用")
        if user is None:
            return _json_error(401, "用户名或密码不正确")
        response = JSONResponse({"ok": True, "username": user[1]})
        _set_session_cookie(response, start_session(user[0]))
        return response

    @_api
    async def logout(request: Request):
        _user_id, session_id = _identity(request)
        revoke_session(session_id)
        response = JSONResponse({"ok": True})
        _clear_session_cookie(response)
        return response

    application.post("/api/logout")(logout)

    @_api
    async def me(request: Request):
        user_id, _session_id = _identity(request)
        with session_scope() as db:
            user = db.get(User, user_id)
            if user is None:
                raise AuthError("未登录或登录已过期", status=401)
            return JSONResponse({"id": user.id, "username": user.username})

    application.get("/api/me")(me)

    @_api
    async def update_password(request: Request):
        user_id, _session_id = _identity(request)
        body = await _json_body(request)
        change_password(user_id, str(body.get("current_password") or ""), str(body.get("new_password") or ""))
        revoke_other_sessions(user_id, _session_id)
        return JSONResponse({"ok": True})

    application.post("/api/password")(update_password)

    @application.get("/")
    def home():
        return FileResponse(STATIC_DIR / "index.html", media_type="text/html; charset=utf-8")

    @application.get("/api/ai-status")
    def ai_status_route():
        return ai_status()

    @_api
    async def list_courses(request: Request):
        user_id, _session_id = _identity(request)
        with session_scope() as db:
            rows = list(
                db.scalars(select(Course).where(Course.user_id == user_id).order_by(Course.updated_at.desc(), Course.id.desc())).all()
            )
            return JSONResponse(
                {
                    "courses": [
                        {
                            "id": row.id,
                            "name": row.name,
                            "updated_at": row.updated_at.isoformat(timespec="seconds") if row.updated_at else None,
                        }
                        for row in rows
                    ]
                }
            )

    application.get("/api/courses")(list_courses)

    @_api
    async def create_course(request: Request):
        user_id, _session_id = _identity(request)
        body = await _json_body(request)
        name = _course_name_ok(str(body.get("name") or ""))
        settings = body.get("settings") or {}
        if not isinstance(settings, dict):
            raise ServiceError("课程设置格式不正确")
        settings = absorb_relation_grid(settings, strict=False)
        with session_scope() as db:
            course = Course(
                user_id=user_id,
                name=name,
                settings_json="{}",
                updated_at=utcnow(),
            )
            db.add(course)
            db.flush()
            create_term(db, course, settings, make_current=True)
            assign_course_settings(course, settings)
            return JSONResponse(_course_dict(db, course))

    application.post("/api/courses")(create_course)

    @_api
    async def get_course(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            return JSONResponse(_course_dict(db, course))

    application.get("/api/courses/{course_id}")(get_course)

    @_api
    async def update_course(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        body = await _json_body(request)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            if body.get("name") is not None:
                course.name = _course_name_ok(str(body.get("name") or ""))
            if body.get("settings") is not None:
                settings = body.get("settings")
                if not isinstance(settings, dict):
                    raise ServiceError("课程设置格式不正确")
                settings = absorb_relation_grid(settings, strict=False)
                term = current_term(db, course)
                if term is None and _term_text_requested(settings):
                    term = create_term(db, course, settings, make_current=True, clear_identity=True)
                if term is not None:
                    apply_term_fields(term, settings)
                assign_course_settings(course, settings)
            course.updated_at = utcnow()
            return JSONResponse(_course_dict(db, course))

    application.patch("/api/courses/{course_id}")(update_course)

    @_api
    async def delete_course(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            for row in _course_files(db, user_id, course.id):
                db.delete(row)
            db.flush()
            for term in list_terms(db, course.id):
                db.delete(term)
            db.flush()
            db.delete(course)
        remove_tree(user_id, course_id)
        return JSONResponse({"ok": True})

    application.delete("/api/courses/{course_id}")(delete_course)

    @_api
    async def list_files(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            term = current_term(db, course)
            if term is None:
                return JSONResponse({"files": []})
            rows = _course_files(db, user_id, course_id, term.id)
            return JSONResponse({"files": [_file_dict(row) for row in rows]})

    application.get("/api/courses/{course_id}/files")(list_files)

    @_api
    async def upload_course_file(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        form = await _optional_form(request)
        try:
            if form is None:
                raise ServiceError("请上传 Excel 成绩文件（.xlsx）")
            kind = str(form.get("kind") or "")
            if kind not in {"grade", "previous"}:
                raise ServiceError("文件类型只能是成绩表或上一学年达成度表")
            upload = _as_upload(form.get("file"))
            data = await _read_upload(upload, required=True)
            assert upload is not None and data is not None
            with session_scope() as db:
                course = _owned_course(db, user_id, course_id)
                term = current_term(db, course)
                if term is None:
                    raise ServiceError("请先导入本学期成绩登记表")
                row = save_blob(db, user_id, course_id, upload.filename or "grades.xlsx", data, kind, term_id=term.id)
                return JSONResponse(_file_dict(row))
        finally:
            await _close_form(form)

    application.post("/api/courses/{course_id}/files")(upload_course_file)

    @_api
    async def download_course_file(course_id: int, file_id: int, request: Request):
        user_id, _session_id = _identity(request)
        with session_scope() as db:
            row = _owned_file(db, user_id, course_id, file_id)
            data = read_blob(row)
            filename = download_filename(row.original_name)
        return _attachment(filename, data, _media_type(filename))

    application.get("/api/courses/{course_id}/files/{file_id}")(download_course_file)

    async def _load_settings(request: Request, user_id: int, course_id: int) -> dict:
        form = await _optional_form(request)
        try:
            settings_raw = form.get("settings") if form is not None else None
            with session_scope() as db:
                course = _owned_course(db, user_id, course_id)
                term = current_term(db, course)
                if term is None:
                    raise ServiceError("请先导入本学期成绩登记表")
                if isinstance(settings_raw, str):
                    settings = _parse_settings(settings_raw)
                    settings = _store_settings(course, settings)
                    apply_term_fields(term, settings)
                else:
                    settings = _settings_of(course)
                return _job_settings(settings, term)
        finally:
            await _close_form(form)

    async def _load_grade_job(request: Request, user_id: int, course_id: int):
        form = await _optional_form(request)
        try:
            settings_raw = form.get("settings") if form is not None else None
            grade_up = _as_upload(form.get("file")) if form is not None else None
            prev_up = _as_upload(form.get("previous")) if form is not None else None
            grade_bytes = await _read_upload(grade_up, required=False)
            prev_bytes = await _read_upload(prev_up, required=False)
            with session_scope() as db:
                course = _owned_course(db, user_id, course_id)
                term = current_term(db, course)
                if term is None:
                    raise ServiceError("请先导入本学期成绩登记表")
                if isinstance(settings_raw, str):
                    settings = _parse_settings(settings_raw)
                    settings = _store_settings(course, settings)
                    apply_term_fields(term, settings)
                else:
                    settings = _settings_of(course)
                settings = _job_settings(settings, term)
                if grade_bytes is not None and grade_up is not None:
                    save_blob(db, user_id, course_id, grade_up.filename or "grades.xlsx", grade_bytes, "grade", term_id=term.id)
                else:
                    grade_bytes = _latest_bytes(db, user_id, course_id, "grade", term.id)
                if prev_bytes is not None and prev_up is not None:
                    save_blob(db, user_id, course_id, prev_up.filename or "previous.xlsx", prev_bytes, "previous", term_id=term.id)
                else:
                    prev_bytes = _latest_bytes(db, user_id, course_id, "previous", term.id)
            if not grade_bytes:
                raise ServiceError("请上传 Excel 成绩文件（.xlsx）")
            return settings, grade_bytes, prev_bytes
        finally:
            await _close_form(form)

    @_api
    async def template(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        try:
            settings = await _load_settings(request, user_id, course_id)
            filename, content = await asyncio.to_thread(build_template, settings)
            _persist_files(user_id, course_id, [(filename, content)], "template")
        except (NotFound, ServiceError, ValueError, AuthError):
            raise
        except Exception as exc:
            return _json_error(500, f"模板生成失败：{exc}")
        return _attachment(filename, content, XLSX_MEDIA)

    application.post("/api/courses/{course_id}/template")(template)

    @_api
    async def calculate(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        try:
            settings, excel, previous = await _load_grade_job(request, user_id, course_id)
            captured: list[tuple[str, bytes]] = []
            summary = await asyncio.to_thread(run_calculation, excel, previous, settings, captured)
            _persist_files(user_id, course_id, captured, "output")
            with session_scope() as db:
                course = _owned_course(db, user_id, course_id)
                term = current_term(db, course)
                store_last_achievement(term, summary.get("achievement"))
                summary["term_id"] = term.id
                summary["term_label"] = term.label
                summary["class_name"] = term.class_name
        except (NotFound, ServiceError, ValueError, AuthError):
            raise
        except Exception as exc:
            return _json_error(500, f"计算失败：{exc}")
        return JSONResponse(summary)

    application.post("/api/courses/{course_id}/calculate")(calculate)

    @_api
    async def export(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        try:
            settings, excel, previous = await _load_grade_job(request, user_id, course_id)
            captured: list[tuple[str, bytes]] = []
            filename, content, _summary = await asyncio.to_thread(run_export, excel, previous, settings, captured)
            _persist_files(user_id, course_id, captured + [(filename, content)], "output")
        except (NotFound, ServiceError, ValueError, AuthError):
            raise
        except Exception as exc:
            return _json_error(500, f"导出失败：{exc}")
        return _attachment(filename, content, "application/zip")

    application.post("/api/courses/{course_id}/export")(export)

    @_api
    async def ai_report(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        if not ai_limiter.allow(_client_key(request)):
            return _json_error(429, "AI 报告请求过于频繁，请稍后再试")
        try:
            settings, excel, previous = await _load_grade_job(request, user_id, course_id)
            captured: list[tuple[str, bytes]] = []
            filename, content = await asyncio.to_thread(run_ai_report, excel, previous, settings, captured)
            _persist_files(user_id, course_id, captured + [(filename, content)], "report")
        except (NotFound, ServiceError, ValueError, AuthError):
            raise
        except Exception as exc:
            return _json_error(500, f"AI 报告生成失败：{exc}")
        return _attachment(filename, content, "application/zip")

    application.post("/api/courses/{course_id}/ai-report")(ai_report)

    @_api
    async def start_report_job(course_id: int, request: Request):
        user_id, _session_id = _identity(request)
        settings, excel, previous = await _load_grade_job(request, user_id, course_id)
        with session_scope() as db:
            course = _owned_course(db, user_id, course_id)
            term = current_term(db, course)
            term_id = term.id
            meta = {"term_id": term.id, "term_label": term.label or "", "class_name": term.class_name or ""}

        def persist(files: list[tuple[str, bytes]]) -> None:
            _persist_files(user_id, course_id, files, "report", term_id=term_id)

        view = launch_report_job(
            user_id=user_id,
            course_id=course_id,
            term_id=term_id,
            excel=excel,
            previous=previous,
            settings=settings,
            meta=meta,
            persist=persist,
            allow_ai=lambda: ai_limiter.allow(_client_key(request)),
        )
        return JSONResponse(view)

    application.post("/api/courses/{course_id}/report-jobs")(start_report_job)

    @_api
    async def read_report_job(job_id: str, request: Request):
        user_id, _session_id = _identity(request)
        view = report_job_snapshot(user_id, job_id)
        if view is None:
            raise NotFound()
        return JSONResponse(view)

    application.get("/api/report-jobs/{job_id}")(read_report_job)

    @_api
    async def download_report_job(job_id: str, request: Request):
        user_id, _session_id = _identity(request)
        filename, content = report_job_download(user_id, job_id)
        return _attachment(filename, content, "application/zip")

    application.get("/api/report-jobs/{job_id}/download")(download_report_job)

    from web_app.admin_routes import AdminHostMiddleware, attach_admin
    from web_app.v2_routes import attach

    attach(application)
    attach_admin(application)
    application.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
    application.add_middleware(_Guard)
    application.add_middleware(AdminHostMiddleware)
    return application


app = create_app()
