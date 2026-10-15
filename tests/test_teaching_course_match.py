"""A renamed timetable course is never used until the teacher confirms its mapping."""
import copy
import json

from sqlalchemy import select
from sqlalchemy.orm import Session

from tests.test_web_flow import app, client
from tests.test_course_calendar import create_course, calendar_data, cached_template, converted_template, workbook_from_download
from tests.test_teaching_assignments import add_schedule, make_plan, decoy_grades
from web_app.db import Course, TeachingMetadataOverride, Term, get_engine
from web_app.teaching import course_assignments as assignments


def test_similar_course_name_requires_confirmation_then_reuses_teacher_and_class(client):
    course = create_course(client, calendar_data(), name="3D软件基础")
    term_id = decoy_grades(course)
    add_schedule("3D设计软件基础", [
        ("黄楚峰", ["25数字媒体A"]),
        ("王阳子", ["25数字媒体B"]),
    ])
    url = f"/api/courses/{course['id']}/teaching-assignment"
    missing = client.get(url).json()
    assert missing["status"] == "missing"
    assert missing["teacher"] == "" and missing["classes"] == []
    assert "3D设计软件基础" in missing["course_matches"]
    assert missing["matched_course_name"] == ""

    plan = make_plan(client, course)
    assert plan["teaching_assignment"]["status"] == "missing"
    update_url = f"/api/lesson-plans/{plan['id']}"
    invalid = client.patch(update_url, json={
        "revision": plan["revision"], "context": plan["context"],
        "teaching_metadata": {"mode": "course_match", "revision": 0, "course_name": "另一门课"},
    })
    assert invalid.status_code == 409
    assert client.get(url).json()["matched_course_name"] == ""

    matched = client.patch(update_url, json={
        "revision": plan["revision"], "context": plan["context"],
        "teaching_metadata": {"mode": "course_match", "revision": 0,
                              "course_name": "3D设计软件基础"},
    })
    assert matched.status_code == 200, matched.text
    matched = matched.json()
    assert matched["teaching_assignment"]["matched_course_name"] == "3D设计软件基础"
    assert matched["teaching_assignment"]["status"] == "selection_required"
    assert {r["teacher"] for r in matched["teaching_assignment"]["options"]} == {"黄楚峰", "王阳子"}
    assert matched["teaching_assignment"]["has_manual"] is False

    selected = client.patch(update_url, json={
        "revision": matched["revision"], "context": matched["context"], "teacher_choice": "黄楚峰",
    })
    assert selected.status_code == 200, selected.text
    selected = selected.json()
    assert selected["teaching_assignment"]["status"] == "ready"
    assert selected["context"]["授课教师"] == "黄楚峰"
    assert selected["context"]["授课对象"] == "25数字媒体A"
    assert client.get(url).json()["teacher"] == "黄楚峰"
    assert client.get(url).json()["classes"] == ["25数字媒体A"]
    assert client.get(update_url).json()["context"]["授课教师"] == "黄楚峰"

    with Session(get_engine()) as db:
        grade = db.get(Term, term_id)
        assert grade.teacher == "历史成绩单教师"
        assert grade.class_name == "历史成绩单班级"
        assert json.loads(db.get(Course, course["id"]).settings_json)["course_basic_info"]["teacher"] == "成绩单教师"
        rows = list(db.scalars(select(TeachingMetadataOverride).where(
            TeachingMetadataOverride.course_id == course["id"])))
        assert len(rows) == 1
        assert json.loads(rows[0].fields_json) == {"_course_match": "3D设计软件基础"}


def test_course_match_is_scoped_by_semester_and_can_be_cleared(client):
    course = create_course(client, calendar_data(), name="3D软件基础")
    add_schedule("3D设计软件基础", [("黄楚峰", ["25数字媒体A"])])
    plan = make_plan(client, course)
    url = f"/api/lesson-plans/{plan['id']}"
    matched = client.patch(url, json={
        "revision": plan["revision"], "context": plan["context"],
        "teaching_metadata": {"mode": "course_match", "revision": 0,
                              "course_name": "3D设计软件基础"},
    }).json()
    assert matched["teaching_assignment"]["status"] == "ready"
    other = copy.deepcopy(matched["context"])
    other["学期"] = "2"
    changed = client.patch(url, json={"revision": matched["revision"], "context": other})
    assert changed.status_code == 200, changed.text
    assert changed.json()["teaching_assignment"]["matched_course_name"] == ""
    assert changed.json()["context"]["授课教师"] == ""

    original = copy.deepcopy(changed.json()["context"])
    original["学期"] = "1"
    back = client.patch(url, json={"revision": changed.json()["revision"], "context": original}).json()
    assert back["teaching_assignment"]["matched_course_name"] == "3D设计软件基础"
    assert back["context"]["授课教师"] == "黄楚峰"
    cleared = client.patch(url, json={
        "revision": back["revision"], "context": back["context"],
        "teaching_metadata": {"mode": "course_match",
                              "revision": back["teaching_assignment"]["override_revision"],
                              "course_name": ""},
    })
    assert cleared.status_code == 200, cleared.text
    assert cleared.json()["teaching_assignment"]["status"] == "missing"
    assert cleared.json()["teaching_assignment"]["matched_course_name"] == ""



def test_calendar_confirms_similar_course_before_using_teacher(client, cached_template, tmp_path):
    course = create_course(client, calendar_data(), name="3D软件基础")
    add_schedule("3D设计软件基础", [
        ("黄楚峰", ["25数字媒体A"]), ("王阳子", ["25数字媒体B"]),
    ])
    calendar_url = f"/api/courses/{course['id']}/teaching-calendar"
    period = assignments.academic_period()
    blocked = client.post(calendar_url + "?progress=1")
    assert blocked.status_code == 409
    assert blocked.json()["code"] == "teaching_course_match_required"
    assert blocked.json()["teaching_assignment"]["course_matches"] == ["3D设计软件基础"]

    match = client.patch(f"/api/courses/{course['id']}/teaching-metadata", json={
        "mode": "course_match", "course_name": "3D设计软件基础", "revision": 0,
        "year_start": period["year_start"], "year_end": period["year_end"],
        "semester": period["semester"],
    })
    assert match.status_code == 200, match.text
    assert match.json()["status"] == "selection_required"
    selected = client.post(f"/api/courses/{course['id']}/teaching-assignment", json={
        "teacher": "王阳子", "year_start": period["year_start"],
        "year_end": period["year_end"], "semester": period["semester"],
    })
    assert selected.status_code == 200, selected.text
    assert selected.json()["classes"] == ["25数字媒体B"]
    calendar = client.post(calendar_url)
    assert calendar.status_code == 200, calendar.text[:500]
    workbook = workbook_from_download(calendar, tmp_path)
    sheet = workbook.active
    assert sheet["M3"].value == "王阳子"
    assert sheet["G3"].value == "25数字媒体B"
    workbook.close()
