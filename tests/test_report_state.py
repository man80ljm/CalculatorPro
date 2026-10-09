"""报告状态随账号保存；有效输入改变才提示旧报告，计算结果保持原样。"""
import copy
import io

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from tests.test_report_job import app, client, _AiResponse, _prepare_course, _wait_job
from tests.test_web_flow import PASSWORD


def _report(client, monkeypatch):
    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-synthetic-test")
    monkeypatch.setattr("core_app.ai_report.requests.post", lambda *args, **kwargs: _AiResponse())
    course_id = _prepare_course(client)
    started = client.post(f"/api/courses/{course_id}/report-jobs")
    assert started.status_code == 200, started.text
    body, _ = _wait_job(client, started.json()["job_id"])
    assert body["done"] and not body["error"], body
    return course_id, body


def _current(client, course_id):
    response = client.get(f"/api/courses/{course_id}")
    assert response.status_code == 200, response.text
    return response.json()


def test_report_persists_summary_on_new_device_and_after_input_cleanup(client, app, monkeypatch):
    from web_app.db import ReportTask, session_scope
    course_id, job = _report(client, monkeypatch)
    state = _current(client, course_id)["report_state"]
    assert state["status"] == "ready" and not state["legacy"]
    assert state["summary"] == job["summary"]
    assert state["created_at"].endswith("Z")
    # 模拟定期清理任务临时输入，已保存状态和摘要仍然完整。
    with session_scope() as db:
        task = db.get(ReportTask, job["job_id"])
        task.settings_json = task.meta_json = "{}"
    with TestClient(app) as new_device:
        assert new_device.post("/api/login", json={"username": "teacher", "password": PASSWORD}).status_code == 200
        restored = _current(new_device, course_id)["report_state"]
        assert restored == state
        archive = new_device.get(restored["download_url"])
        assert archive.status_code == 200 and archive.content.startswith(b"PK")


def test_noop_save_and_derived_fields_stay_ready_edit_and_revert_are_detected(client, monkeypatch):
    course_id, _ = _report(client, monkeypatch)
    settings = _current(client, course_id)["edit_settings"]
    normalized = copy.deepcopy(settings)
    normalized["course_basic_info"]["unused_blank"] = ""
    normalized["grad_req_map"][0]["strength"] = ""
    normalized["grad_req_map"][0]["objective"] = "课程目标1"
    normalized["relation_payload"]["total_sum"] = 1.0
    for item in (settings, normalized):
        saved = client.patch(f"/api/courses/{course_id}", json={"settings": item})
        assert saved.status_code == 200, saved.text
        assert saved.json()["report_state"]["status"] == "ready"
    changed = copy.deepcopy(settings)
    changed["course_basic_info"]["hours"] = "64"
    saved = client.patch(f"/api/courses/{course_id}", json={"settings": changed})
    assert saved.json()["report_state"]["status"] == "stale"
    assert "旧报告仍可下载" in saved.json()["report_state"]["message"]
    restored = client.patch(f"/api/courses/{course_id}", json={"settings": settings})
    assert restored.json()["report_state"]["status"] == "ready"


def test_identical_grade_reupload_stays_ready_changed_grade_retains_old_download(client, monkeypatch):
    course_id, _ = _report(client, monkeypatch)
    current = _current(client, course_id)
    state = current["report_state"]
    old_archive = client.get(state["download_url"]).content
    grade = next(file for file in current["files"] if file["kind"] == "grade")
    grades = client.get(f"/api/courses/{course_id}/files/{grade['id']}").content

    def upload(content):
        response = client.post(f"/api/courses/{course_id}/files", data={"kind": "grade"},
                               files={"file": ("synthetic.xlsx", content)})
        assert response.status_code == 200, response.text

    upload(grades)
    assert _current(client, course_id)["report_state"]["status"] == "ready"
    workbook = load_workbook(io.BytesIO(grades))
    score_cell = next(cell for cell in workbook.active[3] if isinstance(cell.value, (int, float)))
    score_cell.value += 1
    stream = io.BytesIO()
    workbook.save(stream)
    upload(stream.getvalue())
    assert _current(client, course_id)["report_state"]["status"] == "stale"
    assert client.get(state["download_url"]).content == old_archive
    upload(grades)
    assert _current(client, course_id)["report_state"]["status"] == "ready"


def test_failed_new_report_preserves_successful_report(client, monkeypatch):
    course_id, _ = _report(client, monkeypatch)
    original = _current(client, course_id)["report_state"]
    monkeypatch.setattr("web_app.service.generate_answers", lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("合成失败")))
    started = client.post(f"/api/courses/{course_id}/report-jobs")
    assert started.status_code == 200, started.text
    failed, _ = _wait_job(client, started.json()["job_id"])
    assert failed["error"]
    assert _current(client, course_id)["report_state"] == original


def test_legacy_report_edit_tracking_and_regeneration(client):
    from web_app.app import _persist_files
    from web_app.db import Course, session_scope
    course_id = _prepare_course(client)
    term_id = _current(client, course_id)["current_term_id"]
    with session_scope() as db:
        user_id = db.get(Course, course_id).user_id
    _persist_files(user_id, course_id, [("合成旧报告.docx", b"synthetic")], "report", term_id)
    current = _current(client, course_id)
    assert current["report_state"]["status"] == "ready" and current["report_state"]["legacy"]
    settings = current["edit_settings"]
    noop = client.patch(f"/api/courses/{course_id}", json={"settings": settings})
    assert noop.json()["report_state"]["status"] == "ready"
    settings["course_description"] = "已修改的合成简介"
    changed = client.patch(f"/api/courses/{course_id}", json={"settings": settings})
    assert changed.json()["report_state"]["status"] == "stale"
    _persist_files(user_id, course_id, [("合成新报告.docx", b"synthetic-new")], "report", term_id)
    assert _current(client, course_id)["report_state"]["status"] == "ready"


def test_report_state_isolated_by_term(client, monkeypatch):
    from web_app.db import Course, session_scope
    from web_app.terms import create_term
    course_id, _ = _report(client, monkeypatch)
    original = _current(client, course_id)
    with session_scope() as db:
        settings = copy.deepcopy(original["edit_settings"])
        settings["course_open_info"].update(year_start="2028", year_end="2029")
        settings["course_basic_info"]["school_year_term"] = "2028-2029学年第1学期"
        create_term(db, db.get(Course, course_id), settings, make_current=True)
    assert _current(client, course_id)["report_state"] is None
    selected = client.post(f"/api/courses/{course_id}/terms/{original['current_term_id']}/select")
    assert selected.status_code == 200, selected.text
    assert selected.json()["report_state"] == original["report_state"]


@pytest.mark.parametrize("settings", [{}, {"relation_payload": "未填", "grad_req_map": None},
                                     {"relation_payload": {"links": [{"methods": None}]}, "word_limit": "未填", "spread_mode": []}])
def test_state_signature_accepts_incomplete_settings(settings):
    from web_app.report_state import settings_signature
    assert len(settings_signature(settings)) == 64
