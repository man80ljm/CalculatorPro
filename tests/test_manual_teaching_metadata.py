"""Manual teaching metadata survives persistence, period changes and Word/calendar export."""
import copy
import uuid
from unittest.mock import Mock

from sqlalchemy.orm import Session
from tests.test_web_flow import app, client
from tests.test_course_calendar import create_course, calendar_data, cached_template, converted_template, workbook_from_download
from tests.test_lesson_plans import wait_job
from tests.test_teaching_assignments import make_plan, add_schedule, text_in_docx
from web_app.db import TeacherProfile, get_engine
from web_app.lesson_plans import engine, service
from web_app.teaching import course_assignments as assignments
from web_app.teaching.teacher_directory import key


def test_manual_metadata_is_saved_and_not_overwritten_by_new_timetables(client,monkeypatch,cached_template,tmp_path):
    course=create_course(client,calendar_data(),name="合成教案课程")
    plan=make_plan(client,course)
    assert plan["teaching_assignment"]["status"]=="missing"
    assert all(plan["teaching_assignment"]["editable"].values())
    with Session(get_engine()) as db:
        db.add(TeacherProfile(id=uuid.uuid4().hex,name="目录教师",name_key=key("目录教师"),college="合成学院",
            department="设计系",mobile="13800138000",notes="内部资料"))
        db.commit()
    names=client.get("/api/teaching/teachers",params={"q":"目录"}).json()["teachers"]
    assert names[0]["teaching_unit"]=="设计"
    assert "mobile" not in names[0] and "notes" not in names[0]
    context=copy.deepcopy(plan["context"]);context["授课教师"]="目录教师";context["授课对象"]="新学期A班"
    url="/api/lesson-plans/"+plan["id"]
    response=client.patch(url,json={"revision":plan["revision"],"context":context,
        "teaching_metadata":{"mode":"manual","revision":0,"fields":{"teacher":"目录教师","audience":"新学期A班"}}})
    assert response.status_code==200,response.text
    saved=response.json()
    assert saved["context"]["教研室"]=="设计" and saved["context"]["授课教师"]=="目录教师"
    add_schedule("合成教案课程",[("新课表教师",["课表班级"])])
    reopened=client.get(url).json()
    assert reopened["context"]["授课教师"]=="目录教师" and reopened["context"]["授课对象"]=="新学期A班"
    assert reopened["teaching_assignment"]["status"]=="manual"
    original=service.make_document
    monkeypatch.setattr(service,"make_document",lambda context,check=True:original(context,check=False))
    monkeypatch.setattr(engine,"deepseek",Mock(side_effect=AssertionError("export must not call API")))
    export=client.post(url+"/export",json={"revision":reopened["revision"]})
    assert export.status_code==202,export.text
    assert wait_job(client,export.json())["status"]=="success"
    word=client.get(url+"/download"); assert word.status_code==200,word.text[:300]
    text=text_in_docx(word.content)
    assert "目录教师" in text and "新学期A班" in text and "设计" in text
    calendar=client.post(f"/api/courses/{course['id']}/teaching-calendar")
    assert calendar.status_code==200,calendar.text[:300]
    workbook=workbook_from_download(calendar,tmp_path)
    assert workbook.active["M3"].value=="目录教师" and workbook.active["G3"].value=="新学期A班"
    workbook.close()
    # An explicit restore uses the current timetable and invalidates the old Word.
    restored=client.patch(url,json={"revision":reopened["revision"],"context":reopened["context"],
        "teaching_metadata":{"mode":"timetable","revision":reopened["teaching_assignment"]["override_revision"]}})
    assert restored.status_code==200,restored.text
    assert restored.json()["context"]["授课教师"]=="新课表教师"
    assert not restored.json()["download_url"]


def test_manual_metadata_is_period_scoped_and_revisions_are_checked(client):
    course=create_course(client,calendar_data(),name="合成教案课程"); plan=make_plan(client,course)
    period=assignments.academic_period()
    url=f"/api/courses/{course['id']}/teaching-metadata"
    body={**period,"mode":"manual","revision":0,"fields":{"teacher":"手填教师","audience":"手填班级","teaching_unit":"传达与媒体设计系"}}
    response=client.patch(url,json=body); assert response.status_code==200,response.text
    assert response.json()["teacher_info"]["teaching_unit"]=="传达与媒体设计"
    assert client.patch(url,json=body).status_code==409
    assert client.patch(url,json={**body,"revision":1},headers={"Origin":"https://other.example"}).status_code==403
    context=client.get("/api/lesson-plans/"+plan["id"]).json()["context"]
    context["学期"]="2" if period["semester"]=="1" else "1"
    changed=client.patch("/api/lesson-plans/"+plan["id"],json={"revision":plan["revision"],"context":context})
    assert changed.status_code==200,changed.text
    assert changed.json()["context"]["授课教师"]=="" and changed.json()["context"]["授课对象"]==""
    previous=client.get(f"/api/courses/{course['id']}/teaching-assignment",params={key:period[key] for key in ("year_start","year_end","semester")}).json()
    assert previous["teacher"]=="手填教师"
