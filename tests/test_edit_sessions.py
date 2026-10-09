"""会话切换和旧版本保存，使用真实HTTP与独立数据库验证。"""
from copy import deepcopy
from fastapi.testclient import TestClient
from tests.test_web_flow import app, client, _course, _settings, PASSWORD


def test_new_login_rejects_old_writes_and_keeps_account_data(client, app):
    course_id = _course(client, _settings())
    with TestClient(app) as other:
        assert other.post("/api/login", json={"username": "teacher", "password": PASSWORD}).status_code == 200
        lost = client.patch(f"/api/courses/{course_id}", json={"name": "不应被写入"})
        assert lost.status_code == 401
        assert lost.json()["code"] == "session_replaced"
        assert other.get(f"/api/courses/{course_id}").json()["name"] == "测试课程"


def test_old_page_without_revision_cannot_save(client):
    course_id = _course(client, _settings())
    reply = client.patch(f"/api/courses/{course_id}", json={"name": "旧页"}, headers={"X-Course-Revision": ""})
    assert reply.status_code == 428
    assert client.get(f"/api/courses/{course_id}").json()["name"] == "测试课程"


def test_nonconflicting_changes_merge_without_overwriting(client):
    course_id = _course(client, _settings())
    path = f"/api/courses/{course_id}"
    before = client.get(path).json()
    first, second = deepcopy(before["settings"]), deepcopy(before["settings"])
    first["course_basic_info"]["hours"] = "64"
    second["course_basic_info"]["credits"] = "4"
    assert client.patch(path, json={"settings": first}).status_code == 200
    merged = client.patch(path, headers={"X-Course-Revision": before["edit_revision"]}, json={"name": before["name"], "settings": second,
        "base": {"name": before["name"], "settings": before["settings"], "term_id": before["current_term_id"]}})
    assert merged.status_code == 200, merged.text
    basic = merged.json()["settings"]["course_basic_info"]
    assert basic["hours"] == "64" and basic["credits"] == "4"


def test_same_field_conflict_preserves_both_choices(client):
    course_id = _course(client, _settings())
    path = f"/api/courses/{course_id}"
    before = client.get(path).json()
    first, second = deepcopy(before["settings"]), deepcopy(before["settings"])
    first["course_basic_info"]["hours"] = "64"
    second["course_basic_info"]["hours"] = "80"
    client.patch(path, json={"settings": first})
    reply = client.patch(path, headers={"X-Course-Revision": before["edit_revision"]}, json={"settings": second,
        "base": {"name": before["name"], "settings": before["settings"], "term_id": before["current_term_id"]}})
    assert reply.status_code == 409
    assert reply.json()["local_choice"]["settings"]["course_basic_info"]["hours"] == "80"
    assert reply.json()["saved_choice"]["settings"]["course_basic_info"]["hours"] == "64"
    assert client.get(path).json()["settings"]["course_basic_info"]["hours"] == "64"


def test_save_response_contains_current_term_identity(client):
    from scripts.local_load_test import grade_register, settings
    course_id = _course(client, settings(), name="并发合成课程")
    imported = client.post(f"/api/courses/{course_id}/grade-register", files={"file": ("合成成绩.xlsx", grade_register(1,35))})
    assert imported.status_code == 200
    path = f"/api/courses/{course_id}"
    before = client.get(path).json()
    values = deepcopy(before["edit_settings"])
    values["course_basic_info"]["hours"] = "64"
    first = client.patch(path, json={"settings": values}).json()
    assert first["edit_settings"]["course_open_info"]["year_start"] == "2026"
    assert first["edit_settings"]["course_open_info"]["semester"] == "1"
    assert "year_start" not in first["settings"]["course_open_info"]
    second = client.patch(path, json={"settings": first["edit_settings"]}).json()
    assert second["current_term"]["year_start"] == "2026"
    assert second["current_term"]["year_end"] == "2027"
    assert second["current_term"]["semester"] == "1"
