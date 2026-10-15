"""Previewed, transactional timetable deletion with a recoverable archive."""
import hashlib
import json
import re
import uuid
from datetime import datetime

from sqlalchemy import and_, or_, select
from sqlalchemy.exc import IntegrityError

from web_app.db import ScheduleDeletion, ScheduleDocument, ScheduleRecord, utcnow
from .api_settings import ApiError


def encode(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def identifier(value):
    if not isinstance(value, str) or not re.fullmatch(r"[a-f0-9]{32}", value):
        raise ApiError("请选择有效的课表记录。")
    return value


def document_data(row):
    return {"id": row.id, "filename": row.filename, "digest": row.digest,
            "term": row.term, "model": row.model, "created_at": row.created_at.isoformat()}


def record_data(row):
    return {"id": row.id, "document_id": row.document_id, "term_id": row.term_id,
            "course_name": row.course_name, "payload": row.payload}


def summary(scope, label, snapshot):
    documents = snapshot["documents"]
    return {"scope": scope, "label": label, "recordCount": len(snapshot["records"]),
            "documentCount": len(documents),
            "removedDocumentCount": len(snapshot["removed_document_ids"]),
            "filenames": [item["filename"] for item in documents],
            "terms": sorted({json.loads(item["term"])["label"] for item in documents})}


def trash_page(db, before=None):
    query = select(ScheduleDeletion).where(ScheduleDeletion.restored_at.is_(None))
    if before is not None:
        cursor = db.get(ScheduleDeletion, identifier(before))
        if cursor is None:
            raise ApiError("删除记录不存在，请刷新后重试。", 404)
        query = query.where(or_(ScheduleDeletion.created_at < cursor.created_at,
            and_(ScheduleDeletion.created_at == cursor.created_at, ScheduleDeletion.id < cursor.id)))
    rows = db.scalars(query.order_by(ScheduleDeletion.created_at.desc(),
                                     ScheduleDeletion.id.desc()).limit(51)).all()
    return {"items": [dict(summary(row.scope, row.label, json.loads(row.snapshot_json)),
                id=row.id, createdAt=row.created_at.isoformat(timespec="seconds") + "Z") for row in rows[:50]],
            "nextCursor": rows[49].id if len(rows) > 50 else None}


def list_trash(repository, before=None):
    with repository.connect() as db:
        return trash_page(db, before)

def prepare(db, data):
    scope = data.get("scope")
    allowed = {"records": "ids", "document": "documentId", "term": "termId"}
    if not isinstance(scope, str) or scope not in allowed or set(data) - {"scope", allowed[scope], "confirmation"}:
        raise ApiError("请选择删除记录、导入文件或指定学期。")
    # The repository serializes timetable writers. Include related surviving rows
    # in the fingerprint so empty-document cleanup and fallback notices cannot race.
    documents = db.scalars(select(ScheduleDocument).order_by(ScheduleDocument.id)).all()
    all_records = db.scalars(select(ScheduleRecord).order_by(ScheduleRecord.id)).all()
    document_map = {row.id: row for row in documents}
    if scope == "records":
        ids = data.get("ids")
        if not isinstance(ids, list) or not 1 <= len(ids) <= 10000:
            raise ApiError("请勾选 1–10000 条开课记录。")
        wanted = {identifier(item) for item in ids}
        records = [row for row in all_records if row.id in wanted]
        if len(records) != len(wanted):
            raise ApiError("部分记录已变更或删除，请刷新后重新选择。", 409)
        selected_docs = {row.document_id for row in records}
        label = f"已选 {len(records)} 条开课记录"
    elif scope == "document":
        doc_id = identifier(data.get("documentId"))
        if doc_id not in document_map:
            raise ApiError("这份课表已变更或删除，请刷新。", 409)
        selected_docs = {doc_id}
        records = [row for row in all_records if row.document_id == doc_id]
        label = document_map[doc_id].filename
    else:
        term_id = data.get("termId")
        if not isinstance(term_id, str) or not re.fullmatch(r"\d{4}-[12]", term_id):
            raise ApiError("请先选择一个具体学期，不能清空全部学期。")
        selected_docs = {row.id for row in documents if json.loads(row.term).get("id") == term_id}
        if not selected_docs:
            raise ApiError("此学期已没有课表，请刷新。", 409)
        records = [row for row in all_records if row.document_id in selected_docs]
        label = json.loads(document_map[sorted(selected_docs)[0]].term)["label"]
    selected_ids = {row.id for row in records}
    remaining = [row for row in all_records if row.id not in selected_ids]
    still_used = {row.document_id for row in remaining}
    affected_docs = [row for row in documents if row.id in selected_docs]
    snapshot = {"documents": [document_data(row) for row in affected_docs],
                "records": [record_data(row) for row in records],
                "removed_document_ids": sorted(selected_docs - still_used)}
    # Matching uses the latest PDF per normalized course and term. Warn if a
    # previous PDF can become active after removing all matches from that PDF.
    from .course_assignments import name_key
    groups = {}
    for row in all_records:
        groups.setdefault((row.term_id, name_key(row.course_name)), []).append(row)
    fallback = False
    related = []
    for rows in groups.values():
        if not any(row.id in selected_ids for row in rows):
            continue
        related.extend(record_data(row) for row in rows)
        latest = max(rows, key=lambda row: (document_map[row.document_id].created_at, row.document_id)).document_id
        survivors = [row for row in rows if row.id not in selected_ids]
        if survivors and not any(row.document_id == latest for row in survivors):
            fallback = True
    confirmation = hashlib.sha256(encode({"scope": scope, "snapshot": snapshot,
        "related": related, "document_inventory": [record_data(row) for row in all_records
                                                   if row.document_id in selected_docs]}).encode()).hexdigest()
    preview = summary(scope, label, snapshot)
    preview.update(confirmation=confirmation, fallbackWarning=(
        "删除后，部分课程将重新使用仍保留的旧版课表；请核对旧版授课安排。" if fallback else ""))
    return preview, snapshot, records, affected_docs


def preview_delete(repository, data):
    with repository.connect(True) as db:
        return prepare(db, data)[0]


def delete(repository, data):
    confirmation = data.get("confirmation")
    if not isinstance(confirmation, str) or not re.fullmatch(r"[a-f0-9]{64}", confirmation):
        raise ApiError("请先查看删除范围并确认。")
    with repository.connect(True) as db:
        preview, snapshot, records, documents = prepare(db, data)
        if confirmation != preview["confirmation"]:
            raise ApiError("课表已发生变化，请重新查看删除范围后确认。", 409)
        db.add(ScheduleDeletion(id=uuid.uuid4().hex, scope=preview["scope"], label=preview["label"],
                                snapshot_json=encode(snapshot)))
        for row in records:
            db.delete(row)
        db.flush()
        for row in documents:
            if row.id in snapshot["removed_document_ids"]:
                db.delete(row)
    return repository.snapshot()


def restore(repository, data):
    if set(data) != {"id"}:
        raise ApiError("请选择需要恢复的删除记录。")
    entry_id = identifier(data.get("id"))
    try:
        with repository.connect(True) as db:
            entry = db.scalar(select(ScheduleDeletion).where(ScheduleDeletion.id == entry_id).with_for_update())
            if entry is None:
                raise ApiError("删除记录不存在。", 404)
            if entry.restored_at is None:
                snapshot = json.loads(entry.snapshot_json)
                # Restore original document IDs/dates, including after earlier
                # partial deletions, without superseding newer imports.
                for value in snapshot["documents"]:
                    existing = db.get(ScheduleDocument, value["id"])
                    duplicate = db.scalar(select(ScheduleDocument).where(ScheduleDocument.digest == value["digest"]))
                    if (duplicate and duplicate.id != value["id"]) or (existing and existing.digest != value["digest"]):
                        raise ApiError("这份 PDF 已重新导入，不能覆盖新数据；请保留现有课表。", 409)
                    if existing is None:
                        db.add(ScheduleDocument(**dict(value, created_at=datetime.fromisoformat(value["created_at"]))))
                db.flush()
                for value in snapshot["records"]:
                    if db.get(ScheduleRecord, value["id"]) is not None:
                        raise ApiError("部分记录已恢复或变更，请刷新后核对。", 409)
                    db.add(ScheduleRecord(**value))
                entry.restored_at = utcnow()
    except IntegrityError as exc:
        raise ApiError("课表已发生变化，未恢复任何记录，请刷新后重试。", 409) from exc
    return repository.snapshot()