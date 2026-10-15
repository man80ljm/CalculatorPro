"""Administrator timetable removal is scoped, recoverable and isolated from courses."""
import copy
import uuid
from datetime import timedelta

import pytest
from fastapi.testclient import TestClient
from sqlalchemy import select
from sqlalchemy.orm import Session

from tests.test_teaching_admin import app, admin
from tests.test_course_calendar import calendar_data, create_course
from tests.test_teaching_assignments import decoy_grades, make_plan
from web_app.db import Course, ScheduleDocument, Term, get_engine, utcnow
from web_app.teaching import course_assignments as assignments
from web_app.teaching.pdf_schedule import normalize_result
from web_app.teaching.schedule_db import ScheduleDatabase

BASE = "/admin/api/schedule"


def seed_schedule(*, names=("Drawing",), year="2026-2027", semester=1,
                  filename="schedule.pdf", digest=None, newer=0, teacher="Timetable teacher"):
    repository = ScheduleDatabase(get_engine())
    identifier, digest = uuid.uuid4().hex, digest or uuid.uuid4().hex * 2
    repository.create_draft(identifier, filename, digest, "offline")
    result = normalize_result({
        "term": {"academicYear": year, "semester": semester,
                 "startDate": f"{year[:4]}-09-07", "weeks": 19},
        "records": [{"courseName": name, "teachers": [teacher], "classes": ["Class A"],
                     "weeks": [1, 2], "credits": 3, "totalHours": 48,
                     "timePattern": "Tuesday", "sourcePage": 1,
                     "sourceText": name, "issues": []} for name in names], "notes": [],
    })
    repository.update_draft(identifier, "ready", dict(result, pages=1))
    repository.save_draft(identifier, result)
    with Session(get_engine()) as db:
        document = db.scalar(select(ScheduleDocument).where(ScheduleDocument.digest == digest))
        document.created_at = utcnow() + timedelta(minutes=newer)
        document_id = document.id
        db.commit()
    snapshot = repository.snapshot()
    return (next(item for item in snapshot["documents"] if item["id"] == document_id),
            [item for item in snapshot["records"] if item["documentId"] == document_id], digest)


def preview(admin, selection):
    response = admin.post(BASE + "/delete-preview", json=selection)
    assert response.status_code == 200, response.text
    result = response.json()
    assert result["confirmation"]
    return result


def delete(admin, selection):
    proposed = preview(admin, selection)
    response = admin.post(BASE + "/delete", json={**selection, "confirmation": proposed["confirmation"]})
    assert response.status_code == 200, response.text
    snapshot = response.json()
    assert snapshot["trash"]
    return snapshot


def restore(admin, identifier):
    response = admin.post(BASE + "/restore", json={"id": identifier})
    assert response.status_code == 200, response.text
    return response.json()


def record_ids(snapshot):
    return {item["id"] for item in snapshot["records"]}


@pytest.fixture
def teacher(app, admin):
    code = admin.post("/admin/api/invites", json={}).json()["invitation"]["code"]
    with TestClient(app) as client:
        credentials = {"username": "schedule-deletion-teacher", "password": "schedule-deletion-test-pass"}
        assert client.post("/api/register", json={**credentials, "invite_code": code}).status_code == 200
        assert client.post("/api/login", json=credentials).status_code == 200
        yield client


@pytest.mark.parametrize("operation", ["delete-preview", "delete", "restore"])
def test_deletion_endpoints_require_admin_and_csrf(app, admin, teacher, operation):
    url = BASE + "/" + operation
    with TestClient(app) as anonymous:
        assert anonymous.post(url, json={}).status_code == 401
    assert teacher.post(url, json={}).status_code == 401
    assert admin.post(url, json={}, headers={"X-Admin-CSRF": ""}).status_code == 403
    assert admin.post(url, json={}, headers={"Origin": "https://untrusted.example"}).status_code == 403


def test_selected_rows_are_archived_and_restore_their_original_ids(admin):
    document, rows, _ = seed_schedule(names=("Drawing", "Sculpture"))
    original = ScheduleDatabase(get_engine()).snapshot()
    selected = {"scope": "records", "ids": [rows[0]["id"]]}
    proposed = preview(admin, selected)
    assert proposed["recordCount"] == 1 and proposed["documentCount"] == 1
    assert proposed["removedDocumentCount"] == 0
    assert proposed["filenames"] == [document["filename"]]
    snapshot = delete(admin, selected)
    assert record_ids(snapshot) == {rows[1]["id"]}
    assert [item["id"] for item in snapshot["documents"]] == [document["id"]]
    assert snapshot["trash"][0]["recordCount"] == 1
    restored = restore(admin, snapshot["trash"][0]["id"])
    assert record_ids(restored) == record_ids(original)
    assert restored["documents"] == original["documents"]
    assert {item["id"]: item for item in restored["records"]} == {item["id"]: item for item in original["records"]}
    assert not restored["trash"]


def test_deleting_last_record_removes_empty_document_and_releases_pdf_digest(admin):
    document, rows, digest = seed_schedule()
    selection = {"scope": "records", "ids": [rows[0]["id"]]}
    assert preview(admin, selection)["removedDocumentCount"] == 1
    snapshot = delete(admin, selection)
    assert not snapshot["records"] and not snapshot["documents"]
    assert ScheduleDatabase(get_engine()).find_source(digest) is None
    restored = restore(admin, snapshot["trash"][0]["id"])
    assert record_ids(restored) == {rows[0]["id"]}
    assert restored["documents"][0]["id"] == document["id"]
    assert ScheduleDatabase(get_engine()).find_source(digest)["documentId"] == document["id"]


def test_document_delete_preserves_other_imports_in_same_semester(admin):
    document, rows, _ = seed_schedule(names=("Drawing", "Sculpture"), filename="mistaken.pdf")
    other, other_rows, _ = seed_schedule(filename="correct.pdf")
    proposed = preview(admin, {"scope": "document", "documentId": document["id"]})
    assert proposed["recordCount"] == 2 and proposed["removedDocumentCount"] == 1
    snapshot = delete(admin, {"scope": "document", "documentId": document["id"]})
    assert record_ids(snapshot) == {item["id"] for item in other_rows}
    assert [item["id"] for item in snapshot["documents"]] == [other["id"]]
    restored = restore(admin, snapshot["trash"][0]["id"])
    assert record_ids(restored) == {item["id"] for item in rows + other_rows}


def test_term_delete_removes_only_selected_semester_across_documents(admin):
    _, first_rows, _ = seed_schedule(filename="one.pdf")
    _, second_rows, _ = seed_schedule(filename="two.pdf")
    next_term, next_rows, _ = seed_schedule(year="2027-2028", filename="future.pdf")
    other_term, other_rows, _ = seed_schedule(semester=2, filename="spring.pdf")
    selection = {"scope": "term", "termId": first_rows[0]["termId"]}
    proposed = preview(admin, selection)
    assert proposed["recordCount"] == 2 and proposed["removedDocumentCount"] == 2
    assert set(proposed["filenames"]) == {"one.pdf", "two.pdf"}
    snapshot = delete(admin, selection)
    assert record_ids(snapshot) == {item["id"] for item in next_rows + other_rows}
    assert {item["id"] for item in snapshot["documents"]} == {next_term["id"], other_term["id"]}
    restored = restore(admin, snapshot["trash"][0]["id"])
    assert record_ids(restored) == {item["id"] for item in first_rows + second_rows + next_rows + other_rows}


def test_stale_preview_cannot_delete_updated_records(admin):
    _, rows, _ = seed_schedule()
    selection = {"scope": "records", "ids": [rows[0]["id"]]}
    proposed = preview(admin, selection)
    edited = copy.deepcopy(rows[0]); edited["teachers"] = ["Corrected teacher"]
    response = admin.post(BASE + "/record", json={"id": edited["id"], "record": edited})
    assert response.status_code == 200, response.text
    before = response.json()
    response = admin.post(BASE + "/delete", json={**selection, "confirmation": proposed["confirmation"]})
    assert response.status_code == 409, response.text
    after = admin.get(BASE).json()
    assert after["records"] == before["records"] and not after["trash"]


def test_stale_term_preview_cannot_delete_newly_imported_records(admin):
    _, rows, _ = seed_schedule()
    selection = {"scope": "term", "termId": rows[0]["termId"]}
    proposed = preview(admin, selection)
    seed_schedule(filename="new-import.pdf")
    response = admin.post(BASE + "/delete", json={**selection, "confirmation": proposed["confirmation"]})
    assert response.status_code == 409, response.text
    assert len(admin.get(BASE).json()["records"]) == 2


def test_confirmation_cannot_be_reused_for_another_scope(admin):
    document, rows, _ = seed_schedule(names=("Drawing", "Sculpture"))
    proposed = preview(admin, {"scope": "records", "ids": [rows[0]["id"]]})
    response = admin.post(BASE + "/delete", json={"scope": "document", "documentId": document["id"],
                                                "confirmation": proposed["confirmation"]})
    assert response.status_code == 409, response.text
    assert len(admin.get(BASE).json()["records"]) == 2


def test_restore_conflict_after_same_pdf_reimport_does_not_overwrite_new_data(admin):
    document, _, digest = seed_schedule()
    deleted = delete(admin, {"scope": "document", "documentId": document["id"]})
    new_document, new_rows, _ = seed_schedule(digest=digest, teacher="Newly corrected teacher")
    response = admin.post(BASE + "/restore", json={"id": deleted["trash"][0]["id"]})
    assert response.status_code == 409, response.text
    snapshot = admin.get(BASE).json()
    assert record_ids(snapshot) == {item["id"] for item in new_rows}
    assert [item["id"] for item in snapshot["documents"]] == [new_document["id"]]
    assert snapshot["records"][0]["teachers"] == ["Newly corrected teacher"]
    assert snapshot["trash"][0]["id"] == deleted["trash"][0]["id"]


def test_restore_is_idempotent(admin):
    document, rows, _ = seed_schedule()
    deleted = delete(admin, {"scope": "document", "documentId": document["id"]})
    identifier = deleted["trash"][0]["id"]
    first = restore(admin, identifier)
    second = restore(admin, identifier)
    assert second == first
    assert record_ids(second) == {rows[0]["id"]}


@pytest.mark.parametrize("reverse", [False, True])
def test_separate_deletions_of_same_document_can_restore_in_either_order(admin, reverse):
    document, rows, _ = seed_schedule(names=("Drawing", "Sculpture"))
    first = delete(admin, {"scope": "records", "ids": [rows[0]["id"]]})["trash"][0]["id"]
    second_snapshot = delete(admin, {"scope": "records", "ids": [rows[1]["id"]]})
    second = next(item["id"] for item in second_snapshot["trash"] if item["id"] != first)
    assert not second_snapshot["documents"]
    for identifier in ([second, first] if reverse else [first, second]):
        snapshot = restore(admin, identifier)
    assert record_ids(snapshot) == {item["id"] for item in rows}
    assert snapshot["documents"] == [document]
    assert not snapshot["trash"]


def test_old_timetable_fallback_is_disclosed_and_restore_preserves_import_priority(admin):
    seed_schedule(teacher="Old teacher", newer=-10)
    latest, _, _ = seed_schedule(teacher="Latest teacher", newer=10)
    period = assignments.academic_period("2026", "2027", "1")
    def lookup():
        with Session(get_engine()) as db:
            return assignments.lookup(db, 1, "Drawing", period)
    assert lookup()["teacher"] == "Latest teacher"
    selection = {"scope": "document", "documentId": latest["id"]}
    assert preview(admin, selection)["fallbackWarning"]
    deleted = delete(admin, selection)
    assert lookup()["teacher"] == "Old teacher"
    restored = restore(admin, deleted["trash"][0]["id"])
    assert lookup()["teacher"] == "Latest teacher"
    assert next(item for item in restored["documents"] if item["id"] == latest["id"])["createdAt"] == latest["createdAt"]


def test_deletion_invalidates_timetable_lookup_but_preserves_manual_metadata_courses_and_grades(admin, teacher):
    name = "合成教案课程"
    course = create_course(teacher, calendar_data(), name=name)
    grade_id = decoy_grades(course)
    period = assignments.academic_period()
    document, _, _ = seed_schedule(names=(name,), year=f"{period['year_start']}-{period['year_end']}",
                                   semester=int(period["semester"]))
    plan = make_plan(teacher, course)
    response = teacher.patch(f"/api/courses/{course['id']}/teaching-metadata", json={
        **period, "mode": "manual", "revision": 0,
        "fields": {"teacher": "手填教师", "audience": "手填班级", "teaching_unit": "设计系"}})
    assert response.status_code == 200, response.text
    with Session(get_engine()) as db:
        before_settings = db.get(Course, course["id"]).settings_json
        before_grade = db.get(Term, grade_id).settings_json
        user_id = db.get(Course, course["id"]).user_id
    deleted = delete(admin, {"scope": "document", "documentId": document["id"]})
    assert not deleted["records"]
    with Session(get_engine()) as db:
        assert assignments.lookup(db, user_id, name, period)["status"] == "missing"
        assert db.get(Course, course["id"]).settings_json == before_settings
        grade = db.get(Term, grade_id)
        assert grade.settings_json == before_grade
        assert grade.teacher == "历史成绩单教师" and grade.class_name == "历史成绩单班级"
    reopened = teacher.get("/api/lesson-plans/" + plan["id"])
    assert reopened.status_code == 200, reopened.text
    context = reopened.json()["context"]
    assert context["授课教师"] == "手填教师" and context["授课对象"] == "手填班级"
    assert context["教研室"] == "设计"
    assert context["lessons"] == plan["context"]["lessons"]


def test_delete_requires_confirmation_and_rejects_unbounded_scope(admin):
    document, rows, _ = seed_schedule()
    original = admin.get(BASE).json()
    response = admin.post(BASE + "/delete", json={"scope": "document", "documentId": document["id"]})
    assert response.status_code == 400
    for selection in ({"scope": "term", "termId": ""}, {"scope": "term", "termId": "all"},
                      {"scope": "records", "ids": []}, {"scope": "all"}, {"scope": []}):
        assert admin.post(BASE + "/delete-preview", json=selection).status_code == 400
    assert admin.get(BASE).json() == original


def test_multi_document_restore_conflict_rolls_back_every_document(admin):
    first, first_rows, first_digest = seed_schedule(filename="first.pdf")
    second, second_rows, second_digest = seed_schedule(filename="second.pdf")
    deleted = delete(admin, {"scope": "term", "termId": first_rows[0]["termId"]})
    # Documents are restored by ID order; conflict on the later one must also
    # roll back reconstruction of the earlier document in the same transaction.
    later = max([(first["id"], first_digest), (second["id"], second_digest)])
    current, current_rows, _ = seed_schedule(digest=later[1], filename="reimported.pdf")
    response = admin.post(BASE + "/restore", json={"id": deleted["trash"][0]["id"]})
    assert response.status_code == 409, response.text
    snapshot = admin.get(BASE).json()
    assert snapshot["documents"] == [current]
    assert record_ids(snapshot) == {item["id"] for item in current_rows}
    assert len(snapshot["trash"]) == 1 and snapshot["trash"][0]["recordCount"] == 2


def test_trash_pagination_has_no_gaps_or_duplicates_and_checks_access(app, admin):
    import json
    from web_app.db import ScheduleDeletion

    created = utcnow()
    archived_ids = [f"{index:032x}" for index in range(1, 56)]
    snapshot_json = json.dumps({"documents": [], "records": [], "removed_document_ids": []})
    with Session(get_engine()) as db:
        for index, identifier in enumerate(archived_ids):
            db.add(ScheduleDeletion(id=identifier, scope="records", label=f"Deleted batch {index}",
                                    snapshot_json=snapshot_json, created_at=created))
        db.add(ScheduleDeletion(id="f" * 32, scope="records", label="Already restored",
                                snapshot_json=snapshot_json, created_at=created, restored_at=created))
        db.commit()
    with TestClient(app) as anonymous:
        assert anonymous.get(BASE + "/trash").status_code == 401
    snapshot = admin.get(BASE).json()
    assert len(snapshot["trash"]) == 50
    first = admin.get(BASE + "/trash")
    assert first.status_code == 200, first.text
    first = first.json()
    assert first["items"] == snapshot["trash"]
    assert first["nextCursor"] == snapshot["trashNextCursor"] == first["items"][-1]["id"]
    second = admin.get(BASE + "/trash", params={"before": first["nextCursor"]})
    assert second.status_code == 200, second.text
    second = second.json()
    assert len(second["items"]) == 5 and not second["nextCursor"]
    combined = [item["id"] for item in first["items"] + second["items"]]
    assert combined == sorted(archived_ids, reverse=True)
    assert len(set(combined)) == 55
    assert admin.get(BASE + "/trash", params={"before": "invalid"}).status_code == 400
    assert admin.get(BASE + "/trash", params={"before": "a" * 32}).status_code == 404
