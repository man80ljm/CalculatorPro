"""历史版本删除：明确确认、账号隔离、最新版本保护和数据库/磁盘一致性。"""
import json
from concurrent.futures import ThreadPoolExecutor

from fastapi.testclient import TestClient
from sqlalchemy import select
from sqlalchemy.exc import SQLAlchemyError

from tests.test_course_materials import app, client, PASSWORD, _prepare, _catalog
from tests.test_web_flow import _course, _settings
from web_app.app import _persist_files
from web_app.db import CourseFile, FileBatch, ReportTask, Term, session_scope
from web_app.storage import read_blob, resolve_stored, save_blob, course_folder


def _versions(client):
    course_id, grades = _prepare(client)
    user_id = client.get("/api/me").json()["id"]
    term_id = client.get(f"/api/courses/{course_id}").json()["current_term_id"]
    for name in ("旧版合成报告.docx", "新版合成报告.docx"):
        _persist_files(user_id, course_id, [(name, name.encode())], "report", term_id, excel=grades)
    latest, old = _catalog(client, course_id)["terms"][0]["versions"]
    return course_id, user_id, term_id, latest, old


def _delete(client, cid, version, **kwargs):
    return client.delete(f"/api/courses/{cid}/materials/{version['archive']['id']}?confirm=1", **kwargs)


def _files(cid):
    with session_scope() as db:
        return list(db.scalars(select(CourseFile).where(CourseFile.course_id == cid)))


def test_delete_history_removes_only_its_package_and_documents(client):
    cid, uid, tid, latest, old = _versions(client)
    before = client.get(f"/api/courses/{cid}").json()
    latest_bytes = client.get(f"/api/courses/{cid}/files/{latest['archive']['id']}").content
    original = {row.id: read_blob(row) for row in _files(cid)}
    with session_scope() as db:
        term_settings = db.get(Term, tid).settings_json
        db.add(ReportTask(id="a" * 32, user_id=uid, course_id=cid, term_id=tid,
                          status="success", archive_file_id=old["archive"]["id"]))
    removed = {old["archive"]["id"], *(row["id"] for row in old["files"])}
    paths = {row.id: resolve_stored(uid, cid, row.stored_name) for row in _files(cid)}
    response = _delete(client, cid, old)
    assert response.status_code == 200, response.text
    assert response.json()["materials"]["terms"][0]["versions"] == [latest]
    after = response.json()["course"]
    for field in ("name", "settings", "edit_settings", "current_term_id", "report_state"):
        assert after[field] == before[field]
    assert after["edit_revision"] != before["edit_revision"]
    for file_id in removed:
        assert not paths[file_id].exists()
        assert client.get(f"/api/courses/{cid}/files/{file_id}").status_code == 404
    for row in _files(cid):
        assert row.id not in removed and read_blob(row) == original[row.id]
    assert client.get(f"/api/courses/{cid}/files/{latest['archive']['id']}").content == latest_bytes
    assert client.get("/api/report-jobs/" + "a" * 32 + "/download").status_code == 404
    with session_scope() as db:
        assert db.get(Term, tid).settings_json == term_settings
        assert db.get(ReportTask, "a" * 32).archive_file_id is None
        assert len(list(db.scalars(select(FileBatch).where(FileBatch.course_id == cid)))) == 1
    assert not list(course_folder(uid, cid).glob(".deleted-*"))


def test_confirmation_latest_and_source_protection(client):
    cid, _, _, latest, old = _versions(client)
    assert client.delete(f"/api/courses/{cid}/materials/{old['archive']['id']}").status_code == 400
    assert _delete(client, cid, latest).status_code == 409
    source = next(row for row in _files(cid) if row.kind == "grade")
    assert _delete(client, cid, {"archive": {"id": source.id}}).status_code == 404
    assert len(_catalog(client, cid)["terms"][0]["versions"]) == 2
    assert read_blob(source)


def test_history_delete_is_private_and_cannot_target_another_course(client, app):
    cid, _, _, latest, old = _versions(client)
    other_cid = _course(client, _settings(), name="另一门合成课程")
    assert _delete(client, other_cid, old).status_code == 404
    with TestClient(app) as stranger:
        stranger.post("/api/register", json={"username": "other-user", "password": PASSWORD})
        stranger.post("/api/login", json={"username": "other-user", "password": PASSWORD})
        assert _delete(stranger, cid, old).status_code == 404
    assert _catalog(client, cid)["terms"][0]["versions"] == [latest, old]


def test_revision_and_repeated_delete_are_safe(client):
    cid, _, _, latest, old = _versions(client)
    before = client.get(f"/api/courses/{cid}").json()
    changed = before["edit_settings"]
    changed["course_basic_info"]["hours"] = "64"
    assert client.patch(f"/api/courses/{cid}", json={"settings": changed}).status_code == 200
    assert _delete(client, cid, old, headers={"X-Course-Revision": ""}).status_code == 428
    assert _delete(client, cid, old, headers={"X-Course-Revision": before["edit_revision"]}).status_code == 409
    assert _delete(client, cid, old).status_code == 200
    assert _delete(client, cid, old).status_code == 404
    assert _catalog(client, cid)["terms"][0]["versions"] == [latest]


def test_deleting_previous_term_history_does_not_select_it(client):
    cid, _, old_tid, latest, old = _versions(client)
    created = client.post(f"/api/courses/{cid}/terms", json={"year_start": "2026", "year_end": "2027",
        "semester": "1", "class_name": "合成新班", "student_count": 2, "exam_count": 2})
    assert created.status_code == 200
    new_tid = created.json()["current_term_id"]
    assert _delete(client, cid, old).status_code == 200
    assert client.get(f"/api/courses/{cid}").json()["current_term_id"] == new_tid
    terms = _catalog(client, cid)["terms"]
    assert next(term for term in terms if term["id"] == old_tid)["versions"] == [latest]
    assert _delete(client, cid, latest).status_code == 409


def test_legacy_version_deletes_only_known_zip(client):
    cid, _ = _prepare(client)
    uid = client.get("/api/me").json()["id"]
    tid = client.get(f"/api/courses/{cid}").json()["current_term_id"]
    with session_scope() as db:
        loose = save_blob(db, uid, cid, "单独保存的旧文档.docx", b"loose", "report", tid)
        archive = save_blob(db, uid, cid, "旧报告.zip", b"synthetic old zip", "report", tid)
    _persist_files(uid, cid, [("最新合成报告.docx", b"latest")], "report", tid)
    assert _delete(client, cid, {"archive": {"id": archive.id}}).status_code == 200
    assert client.get(f"/api/courses/{cid}/files/{archive.id}").status_code == 404
    assert client.get(f"/api/courses/{cid}/files/{loose.id}").content == b"loose"


def test_shared_members_and_wrong_members_are_retained(client):
    cid, uid, tid, latest, old = _versions(client)
    with session_scope() as db:
        batches = list(db.scalars(select(FileBatch).where(FileBatch.course_id == cid).order_by(FileBatch.id)))
        shared_id = old["files"][0]["id"]
        source = next(row for row in _files(cid) if row.kind == "grade")
        wrong_term = save_blob(db, uid, cid, "另一期合成文档.docx", b"wrong term", "report", None)
        batches[0].file_ids_json = json.dumps([shared_id, source.id, wrong_term.id])
        batches[1].file_ids_json = json.dumps([*json.loads(batches[1].file_ids_json), shared_id])
    assert _delete(client, cid, old).status_code == 200
    for row in (source, wrong_term):
        assert client.get(f"/api/courses/{cid}/files/{row.id}").content == read_blob(row)
    assert client.get(f"/api/courses/{cid}/files/{shared_id}").status_code == 200
    assert any(file["id"] == shared_id for file in _catalog(client, cid)["terms"][0]["versions"][0]["files"])


def test_filesystem_failure_restores_files_and_database(client, monkeypatch):
    cid, uid, _, latest, old = _versions(client)
    original = {row.id: read_blob(row) for row in _files(cid)}
    import web_app.storage as storage
    replace = storage.os.replace
    calls = 0

    def fail_second(source, destination):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise PermissionError("synthetic file busy")
        return replace(source, destination)

    monkeypatch.setattr(storage.os, "replace", fail_second)
    assert _delete(client, cid, old).status_code == 503
    assert _catalog(client, cid)["terms"][0]["versions"] == [latest, old]
    assert {row.id: read_blob(row) for row in _files(cid)} == original
    assert not list(course_folder(uid, cid).glob(".deleted-*"))


def test_database_failure_restores_quarantined_files(client, monkeypatch):
    cid, uid, _, latest, old = _versions(client)
    original = {row.id: read_blob(row) for row in _files(cid)}
    with monkeypatch.context() as patch:
        def fail_catalog(*_args):
            raise SQLAlchemyError("synthetic transaction failure")
        patch.setattr("web_app.app.materials_catalog", fail_catalog)
        assert _delete(client, cid, old).status_code == 503
    assert _catalog(client, cid)["terms"][0]["versions"] == [latest, old]
    assert {row.id: read_blob(row) for row in _files(cid)} == original
    assert not list(course_folder(uid, cid).glob(".deleted-*"))


def test_missing_historical_file_can_be_cleaned_up(client):
    cid, uid, _, latest, old = _versions(client)
    missing = next(row for row in _files(cid) if row.id == old["archive"]["id"])
    resolve_stored(uid, cid, missing.stored_name).unlink()
    assert _delete(client, cid, old).status_code == 200
    assert _catalog(client, cid)["terms"][0]["versions"] == [latest]


def test_concurrent_deletes_keep_latest_downloadable(client):
    cid, _, _, latest, old = _versions(client)
    revision = client.get(f"/api/courses/{cid}").json()["edit_revision"]
    with ThreadPoolExecutor(max_workers=2) as pool:
        responses = list(pool.map(lambda _: _delete(client, cid, old, headers={"X-Course-Revision": revision}), range(2)))
    codes = [response.status_code for response in responses]
    assert codes.count(200) == 1 and all(code in {200, 409, 429} for code in codes), codes
    assert _catalog(client, cid)["terms"][0]["versions"] == [latest]
    assert client.get(f"/api/courses/{cid}/files/{latest['archive']['id']}").status_code == 200
