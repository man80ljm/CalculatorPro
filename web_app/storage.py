"""课程文件只放在 UPLOAD_DIR/<user_id>/<course_id>/<生成的文件名> 下。"""
from __future__ import annotations

import os
import shutil
import uuid
from pathlib import Path
from contextvars import ContextVar

from web_app.db import CourseFile, NotFound

ALLOWED_SUFFIXES = {".xlsx", ".docx", ".zip", ".json"}
job_upload_root: ContextVar[Path | None] = ContextVar("calculatorpro_job_upload_root", default=None)


def upload_root() -> Path:
    fixed = job_upload_root.get()
    if fixed is not None:
        return fixed
    raw = os.environ.get("UPLOAD_DIR", "/data/uploads").strip() or "/data/uploads"
    return Path(raw).expanduser().resolve()


def ensure_upload_root() -> Path:
    root = upload_root()
    root.mkdir(parents=True, exist_ok=True)
    probe = root / ".write-probe"
    try:
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
    except OSError as exc:
        raise RuntimeError(
            f"UPLOAD_DIR {root} 不可写。容器内进程用户是 uid 10001，"
            f"请在宿主机执行 sudo chown -R 10001:10001 {root}"
        ) from exc
    return root


def course_folder(user_id: int, course_id: int) -> Path:
    root = upload_root()
    folder = (root / str(int(user_id)) / str(int(course_id))).resolve()
    if folder != root and root not in folder.parents:
        raise NotFound()
    return folder


def _safe_suffix(original_name: str) -> str:
    suffix = Path(str(original_name or "")).suffix.lower()
    if suffix in ALLOWED_SUFFIXES:
        return suffix
    return ""


def allocate_path(user_id: int, course_id: int, original_name: str) -> tuple[Path, str]:
    folder = course_folder(user_id, course_id)
    folder.mkdir(parents=True, exist_ok=True)
    stored = f"{uuid.uuid4().hex}{_safe_suffix(original_name)}"
    path = (folder / stored).resolve()
    if path.parent != folder:
        raise NotFound()
    return path, stored


def resolve_stored(user_id: int, course_id: int, stored_name: str) -> Path:
    name = str(stored_name or "")
    if not name or name != Path(name).name or ".." in name:
        raise NotFound()
    folder = course_folder(user_id, course_id).resolve()
    path = (folder / name).resolve()
    if path.parent != folder:
        raise NotFound()
    return path


def remember_original_name(original_name: str) -> str:
    text = str(original_name or "").replace("\x00", "").strip()
    text = text.replace("\r", " ").replace("\n", " ")
    return (text or "download.bin")[:255]


def download_filename(original_name: str) -> str:
    base = Path(str(original_name or "").replace("\\", "/")).name
    base = base.replace("\r", "").replace("\n", "").replace('"', "")
    return base or "download.bin"


def save_blob(
    db, user_id: int, course_id: int, original_name: str, data: bytes, kind: str, term_id: int | None = None
) -> CourseFile:
    from sqlalchemy import func, select
    from web_app.resource_limits import setting
    from web_app.service import ServiceError
    used = int(db.scalar(select(func.coalesce(func.sum(CourseFile.size), 0)).where(CourseFile.user_id == int(user_id))) or 0)
    quota = setting("USER_STORAGE_MB", 2048, maximum=102400) * 1024 * 1024
    if used + len(data) > quota:
        raise ServiceError("已保存资料较多，请联系管理员整理旧资料后再试。", status=413)
    if shutil.disk_usage(upload_root()).free < len(data) + 200 * 1024 * 1024:
        raise ServiceError("服务器存储空间不足，请联系管理员处理。", status=503)
    path, stored = allocate_path(user_id, course_id, original_name)
    path.write_bytes(data)
    db.info.setdefault("new_blob_paths", []).append(path)
    row = CourseFile(
        user_id=int(user_id),
        course_id=int(course_id),
        term_id=int(term_id) if term_id else None,
        kind=kind,
        stored_name=stored,
        original_name=remember_original_name(original_name),
        size=len(data),
    )
    db.add(row)
    db.flush()
    if kind in {"grade", "previous"} and term_id:
        from web_app.db import Term
        from web_app.report_state import mark_legacy_changed
        mark_legacy_changed(db, db.get(Term, int(term_id)))
    return row


def read_blob(row: CourseFile) -> bytes:
    path = resolve_stored(row.user_id, row.course_id, row.stored_name)
    if not path.is_file():
        raise NotFound()
    return path.read_bytes()


def remove_tree(user_id: int, course_id: int) -> None:
    folder = course_folder(user_id, course_id)
    root = upload_root()
    if not folder.exists() or folder == root or root not in folder.parents:
        return
    shutil.rmtree(folder)
