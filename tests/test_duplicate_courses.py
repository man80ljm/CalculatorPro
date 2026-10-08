"""同名建课提示、课程代码区分和账号隔离。"""
import copy
from tests.test_grade_register import client
from tests.test_course_versions import new_syllabus
from tests.test_web_flow import PASSWORD


def test_blank_course_repeated_name_returns_existing_without_overwriting(client):
    first = client.post("/api/courses", json={"name": "合成课程"}).json()
    blocked = client.post("/api/courses", json={"name": "  合成课程  "})
    assert blocked.status_code == 409
    assert blocked.json()["code"] == "existing_course"
    assert blocked.json()["courses"][0]["id"] == first["id"]
    assert len(client.get("/api/courses").json()["courses"]) == 1
    assert client.get(f"/api/courses/{first['id']}").json()["settings"] == first["settings"]


def test_syllabus_duplicate_can_be_updated_in_place(client):
    payload = new_syllabus()
    first = client.post("/api/courses/from-syllabus", json=payload).json()
    newer = copy.deepcopy(payload)
    newer["description"] = "修订后的合成大纲"
    blocked = client.post("/api/courses/from-syllabus", json=newer)
    assert blocked.status_code == 409
    candidate = blocked.json()["courses"][0]
    assert candidate["id"] == first["id"]
    assert candidate["syllabus_version"] == 1
    updated = client.post(f"/api/courses/{first['id']}/syllabus-version", json={**newer, "expected_version": candidate["syllabus_version"]})
    assert updated.status_code == 200, updated.text
    assert updated.json()["id"] == first["id"]
    assert updated.json()["latest_syllabus_version"] == 2
    assert len(client.get("/api/courses").json()["courses"]) == 1


def test_same_name_with_different_explicit_codes_can_coexist(client):
    first = client.post("/api/courses/from-syllabus", json=new_syllabus(code="TEST1"))
    second = client.post("/api/courses/from-syllabus", json=new_syllabus(code="TEST2"))
    assert first.status_code == second.status_code == 200
    courses = client.get("/api/courses").json()["courses"]
    assert {course["course_code"] for course in courses} == {"TEST1", "TEST2"}
    blank = client.post("/api/courses", json={"name": "创意手作"})
    assert blank.status_code == 409
    assert len(blank.json()["courses"]) == 2


def test_same_code_with_changed_name_prompts_for_existing_course(client):
    first = client.post("/api/courses/from-syllabus", json=new_syllabus(name="旧课程名", code="TEST1")).json()
    blocked = client.post("/api/courses/from-syllabus", json=new_syllabus(name="新课程名", code=" test1 "))
    assert blocked.status_code == 409
    assert blocked.json()["courses"][0]["id"] == first["id"]


def test_duplicate_detection_only_checks_current_user(client):
    first = client.post("/api/courses/from-syllabus", json=new_syllabus()).json()
    assert client.post("/api/register", json={"username": "other", "password": PASSWORD}).status_code == 200
    assert client.post("/api/login", json={"username": "other", "password": PASSWORD}).status_code == 200
    second = client.post("/api/courses/from-syllabus", json=new_syllabus())
    assert second.status_code == 200, second.text
    assert second.json()["id"] != first["id"]
    blocked = client.post("/api/courses/from-syllabus", json=new_syllabus())
    assert blocked.status_code == 409
    assert all(course["id"] != first["id"] for course in blocked.json()["courses"])
