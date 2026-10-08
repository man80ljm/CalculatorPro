"""新大纲用于新学期，旧学期计算、表5和文件不受影响。"""
import copy
import json

import pytest
from tests.test_grade_register import client, _craft_settings, _register_xlsx, _FORWARD_HEADERS, _FORWARD_ROWS, _post_register
from tests.test_word_register import _table5


def new_syllabus(name="创意手作", code="BJ2230005"):
    return {
        "fields": {"course_name": name, "course_code": code, "course_type": "选修", "college": "合成学院"},
        "description": "新大纲的合成简介",
        "objectives": ["新目标甲", "新目标乙"],
        "grad_req_map": [{"requirement": "要求1", "indicator": "指标1", "strength": "H"}, {"requirement": "要求2", "indicator": "指标2", "strength": "M"}],
        "relation_grid": [["考核环节", "占比", "考核方式", "课程目标1", "课程目标2", "小计"],
                          ["平时考核", "30%", "课堂参与", "10%", "30%", "40%"],
                          ["", "", "作业", "40%", "20%", "60%"],
                          ["期末考核", "70%", "考查（设计作品）", "20%", "30%", "50%"],
                          ["", "", "课程考核情况表", "30%", "20%", "50%"]],
    }


def setup_course(client):
    created = client.post("/api/courses", json={"name": "创意手作", "settings": _craft_settings()}).json()
    grades = _register_xlsx(_FORWARD_HEADERS, _FORWARD_ROWS)
    imported = _post_register(client, created["id"], grades)
    assert imported.status_code == 200, imported.text
    return created["id"], imported.json()["term_id"], grades


@pytest.mark.parametrize("mode", ["forward", "reverse"])
def test_new_version_preserves_old_calculation_and_table5(client, mode):
    course_id, old_term, grades = setup_course(client)
    stored = client.get(f"/api/courses/{course_id}").json()["settings"]
    stored["mode"] = mode
    assert client.patch(f"/api/courses/{course_id}", json={"settings": stored}).status_code == 200
    before = client.post(f"/api/courses/{course_id}/calculate").json()
    table = _table5(client, course_id)
    old_files = client.get(f"/api/courses/{course_id}").json()["files"]
    old_grade = next(item for item in old_files if item["kind"] == "grade")
    original_bytes = client.get(f"/api/courses/{course_id}/files/{old_grade['id']}").content
    updated = client.post(f"/api/courses/{course_id}/syllabus-version", json={**new_syllabus(), "expected_version": 1})
    assert updated.status_code == 200, updated.text
    assert updated.json()["latest_syllabus_version"] == 2
    assert updated.json()["current_term"]["syllabus_version"] == 1
    assert updated.json()["settings"]["relation_payload"] == stored["relation_payload"]
    after = client.post(f"/api/courses/{course_id}/calculate").json()
    for key in ("term_id", "student_count", "average_score", "achievement"):
        assert after[key] == before[key]
    assert _table5(client, course_id) == table
    assert client.get(f"/api/courses/{course_id}/files/{old_grade['id']}").content == original_bytes
    later = _post_register(client, course_id, _register_xlsx(_FORWARD_HEADERS, _FORWARD_ROWS, term="2026-2027学年第1学期"))
    assert later.status_code == 200, later.text
    assert any("大纲已更新" in warning for warning in later.json()["warnings"])
    assert later.json()["consistent"] is False
    current = client.get(f"/api/courses/{course_id}").json()
    assert current["current_term"]["syllabus_version"] == 2
    assert current["settings"]["objective_requirements"] == ["新目标甲", "新目标乙"]
    back = client.post(f"/api/courses/{course_id}/terms/{old_term}/select").json()
    assert back["settings"]["relation_payload"] == stored["relation_payload"]
    assert back["settings"]["objective_requirements"] == stored["objective_requirements"]


def test_editing_old_term_does_not_replace_new_course_defaults(client):
    course_id, old_term, _ = setup_course(client)
    assert client.post(f"/api/courses/{course_id}/syllabus-version", json=new_syllabus()).status_code == 200
    old = client.get(f"/api/courses/{course_id}").json()["settings"]
    old["course_description"] = "修正旧学期简介"
    assert client.patch(f"/api/courses/{course_id}", json={"settings": old}).status_code == 200
    assert client.get(f"/api/courses/{course_id}").json()["settings"]["course_description"] == "修正旧学期简介"
    later = _post_register(client, course_id, _register_xlsx(_FORWARD_HEADERS, _FORWARD_ROWS, term="2026-2027学年第1学期"))
    assert later.status_code == 200, later.text
    assert client.get(f"/api/courses/{course_id}").json()["settings"]["course_description"] == "新大纲的合成简介"


def test_legacy_terms_are_frozen_before_new_version(client):
    from web_app.db import Course, CourseVersion, Term, session_scope
    from web_app.terms import parse_settings
    from sqlalchemy import delete
    course_id, term_id, _ = setup_course(client)
    with session_scope() as db:
        course = db.get(Course, course_id)
        stored = parse_settings(course.settings_json)
        for key in ("syllabus_version", "syllabus_version_id"):
            stored.pop(key, None)
        course.settings_json = json.dumps(stored)
        term = db.get(Term, term_id)
        blob = parse_settings(term.settings_json)
        for key in ("syllabus_settings", "syllabus_version", "syllabus_version_id"):
            blob.pop(key, None)
        term.settings_json = json.dumps(blob)
        db.execute(delete(CourseVersion).where(CourseVersion.course_id == course_id))
    updated = client.post(f"/api/courses/{course_id}/syllabus-version", json=new_syllabus())
    assert updated.status_code == 200, updated.text
    assert updated.json()["current_term"]["syllabus_version"] == 1
    assert updated.json()["settings"]["relation_payload"] == stored["relation_payload"]


def test_new_version_before_any_grade_creates_no_empty_term(client):
    first = client.post("/api/courses/from-syllabus", json=new_syllabus()).json()
    updated = client.post(f"/api/courses/{first['id']}/syllabus-version", json=new_syllabus())
    assert updated.status_code == 200, updated.text
    assert updated.json()["terms"] == []
    assert updated.json()["current_term"] is None
    assert updated.json()["latest_syllabus_version"] == 2


def test_stale_version_and_wrong_owner_are_rejected(client):
    course_id, _, _ = setup_course(client)
    assert client.post(f"/api/courses/{course_id}/syllabus-version", json={**new_syllabus(), "expected_version": 1}).status_code == 200
    stale = client.post(f"/api/courses/{course_id}/syllabus-version", json={**new_syllabus(), "expected_version": 1})
    assert stale.status_code == 409
    assert stale.json()["code"] == "stale_version"
    from tests.test_web_flow import PASSWORD
    assert client.post("/api/register", json={"username": "other", "password": PASSWORD}).status_code == 200
    assert client.post("/api/login", json={"username": "other", "password": PASSWORD}).status_code == 200
    assert client.post(f"/api/courses/{course_id}/syllabus-version", json=new_syllabus()).status_code == 404
