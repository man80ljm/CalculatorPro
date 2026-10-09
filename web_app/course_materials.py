"""按学期查看资料和生成版本；下载使用生成时保存的同一套文件。"""
from __future__ import annotations

import json
import zipfile
from io import BytesIO

from sqlalchemy import select

from web_app.db import CourseFile, FileBatch, NotFound, ReportTask
from web_app.download_names import archive_filename
from web_app.storage import download_filename
from web_app.terms import list_terms, term_public


def file_public(row: CourseFile) -> dict:
    return {
        "id": row.id, "course_id": row.course_id, "term_id": row.term_id,
        "kind": row.kind, "original_name": row.original_name, "size": row.size,
        "created_at": row.created_at.isoformat(timespec="seconds") + "Z" if row.created_at else None,
    }


def make_archive(files: list[tuple[str, bytes]], excel: bytes | None = None,
                 previous: bytes | None = None) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        names = set()
        for name, content in files:
            name = download_filename(name)
            if name.lower().endswith(".zip") or name in names:
                continue
            archive.writestr(name, content)
            names.add(name)
        if excel is not None:
            archive.writestr("导入资料/本学期成绩表.xlsx", excel)
        if previous is not None:
            archive.writestr("导入资料/上一轮达成度.xlsx", previous)
    return buffer.getvalue()


def history_version_files(db, course, archive_id: int) -> list[CourseFile]:
    """在用户写入锁内核对历史包；只移除能明确归属且未被其他版本引用的文件。"""
    from web_app.service import ServiceError
    archive = db.scalar(select(CourseFile).where(
        CourseFile.id == archive_id, CourseFile.user_id == course.user_id, CourseFile.course_id == course.id,
    ))
    if archive is None or archive.kind not in {"report", "output"} or not archive.original_name.lower().endswith(".zip"):
        raise NotFound()
    packages = list(db.scalars(select(CourseFile).where(
        CourseFile.user_id == course.user_id, CourseFile.course_id == course.id,
        CourseFile.term_id == archive.term_id, CourseFile.kind.in_(("report", "output")),
    ).order_by(CourseFile.id.desc())))
    latest = next((row for row in packages if row.original_name.lower().endswith(".zip")), None)
    if latest is None or latest.id == archive.id:
        raise ServiceError("请保留本学期最新资料包，只能删除历史版本。", status=409)
    tasks = list(db.scalars(select(ReportTask).where(ReportTask.archive_file_id == archive.id)))
    if any(task.status in {"running", "queued"} for task in tasks):
        raise ServiceError("这一版仍在生成，请完成后再删除。", status=409)
    batches = list(db.scalars(select(FileBatch).where(
        FileBatch.user_id == course.user_id, FileBatch.course_id == course.id,
    )))
    batch = next((item for item in batches if item.archive_file_id == archive.id), None)
    members = []
    if batch is not None:
        protected = {item.archive_file_id for item in batches if item.id != batch.id}
        protected.update(file_id for item in batches if item.id != batch.id for file_id in json.loads(item.file_ids_json))
        member_ids = set(json.loads(batch.file_ids_json)) - protected
        if member_ids:
            members = list(db.scalars(select(CourseFile).where(
                CourseFile.id.in_(member_ids), CourseFile.user_id == course.user_id, CourseFile.course_id == course.id,
                CourseFile.term_id == archive.term_id, CourseFile.kind == archive.kind,
            )))
            members = [row for row in members if not row.original_name.lower().endswith(".zip")]
        db.delete(batch)
    # 保留任务统计，删除后的任务下载明确返回不存在。
    for task in tasks:
        task.archive_file_id = None
    return [archive, *members]


def materials_catalog(db, course) -> dict:
    rows = list(db.scalars(select(CourseFile).where(
        CourseFile.user_id == course.user_id, CourseFile.course_id == course.id,
    ).order_by(CourseFile.id.desc())).all())
    batches = list(db.scalars(select(FileBatch).where(
        FileBatch.user_id == course.user_id, FileBatch.course_id == course.id,
    )).all())
    batch_by_archive = {batch.archive_file_id: batch for batch in batches}
    by_id = {row.id: row for row in rows}
    items = []
    for term in list_terms(db, course.id):
        term_files = [row for row in rows if row.term_id == term.id]
        versions = []
        for row in term_files:
            if row.kind not in {"report", "output"} or not row.original_name.lower().endswith(".zip"):
                continue
            batch = batch_by_archive.get(row.id)
            member_ids = json.loads(batch.file_ids_json) if batch else []
            members = [by_id[file_id] for file_id in member_ids if file_id in by_id
                       and by_id[file_id].term_id == term.id]
            versions.append({
                "archive": file_public(row), "kind": row.kind,
                "files": [file_public(member) for member in members],
                "legacy": batch is None,
            })
        item = term_public(term, len(term_files))
        item["download_name"] = archive_filename(course.name, term.year_start, term.year_end,
                                                  term.semester, term.school_year_term)
        item["versions"] = versions
        item["source_files"] = [file_public(row) for row in term_files if row.kind in {"grade", "previous", "template"}]
        # 未归入版本的旧单文件继续保留，避免升级界面后找不到旧计算资料。
        grouped_ids = {member["id"] for version in versions for member in version["files"]}
        item["other_files"] = [file_public(row) for row in term_files
                               if row.kind in {"output", "report"} and row.id not in grouped_ids
                               and not row.original_name.lower().endswith(".zip")]
        items.append(item)
    items.sort(key=lambda item: (int(item["year_start"]) if str(item["year_start"]).isdigit() else 0,
                                 int(item["year_end"]) if str(item["year_end"]).isdigit() else 0,
                                 int(item["semester"]) if str(item["semester"]).isdigit() else 0, item["id"]), reverse=True)
    return {"course_id": course.id, "course_name": course.name, "terms": items}
