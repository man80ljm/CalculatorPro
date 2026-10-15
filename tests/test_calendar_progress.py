"""流式进度、文件完整性、权限、失败与转换清理；不调用外部 API。"""
import asyncio
import base64
import json
import threading
from unittest.mock import Mock

import pytest
import xlrd
from fastapi.testclient import TestClient

from tests.test_web_flow import app, client, PASSWORD
from tests.test_course_calendar import calendar_data, create_course, cached_template, converted_template
from web_app.teaching_calendar import progress, service
from web_app.teaching_calendar.errors import CalendarError


def events(response):
    return [json.loads(line) for line in response.content.splitlines() if line]


def test_progress_reports_real_stages_and_transfers_valid_school_xls(client, monkeypatch, cached_template):
    calendar = calendar_data()
    calendar["mode"] = "ai"
    course = create_course(client, calendar)
    forbidden = Mock(side_effect=AssertionError("Download must not call AI"))
    monkeypatch.setattr("web_app.syllabus.service.complete", forbidden)
    monkeypatch.setattr("web_app.teaching_calendar.engine.call_deepseek", forbidden)
    temporary = service.TemporaryDirectory
    created = []
    def tracked(*args, **kwargs):
        directory = temporary(*args, **kwargs)
        created.append(directory.name)
        return directory
    monkeypatch.setattr(service, "TemporaryDirectory", tracked)
    response = client.post(f"/api/courses/{course['id']}/teaching-calendar?progress=1")
    assert response.status_code == 200, response.text[:1000]
    assert response.headers["content-type"].startswith("application/x-ndjson")
    assert response.headers["X-Accel-Buffering"] == "no" and response.headers["Cache-Control"] == "no-store"
    assert response.headers["X-Course-Revision"] == course["edit_revision"]
    data = events(response)
    stages = [event for event in data if event["type"] == "progress"]
    assert [event["percent"] for event in stages] == [10,15,35,50,65,80,95]
    assert "AI 整理内容" in stages[0]["message"]
    meta = next(event for event in data if event["type"] == "file")
    assert meta["filename"].endswith(".xls") and meta["summary"]["hours_match"] is True
    content = b"".join(base64.b64decode(event["data"]) for event in data if event["type"] == "chunk")
    assert len(content) == meta["size"] and content.startswith(service.OLE_MAGIC)
    assert data[-1] == {"type":"complete"}
    workbook = xlrd.open_workbook(file_contents=content, formatting_info=True)
    ws = workbook.sheet_by_index(0)
    for index, row in enumerate(calendar["lessons"],7):
        assert ws.cell_type(index,0) == xlrd.XL_CELL_TEXT
        assert ws.cell_type(index,1) == xlrd.XL_CELL_NUMBER and ws.cell_value(index,1) == float(row["学时"])
        assert workbook.format_map[workbook.xf_list[ws.cell_xf_index(index,1)].format_key].format_str == "General"
        assert ws.cell_value(index,3) == row["教学内容"]
        assert (index,index+1,3,14) in ws.merged_cells and (index,index+1,14,20) in ws.merged_cells
    footer = 7 + len(calendar["lessons"])
    assert (footer,footer+4,0,11) in ws.merged_cells and (footer,footer+4,11,21) in ws.merged_cells
    workbook.release_resources()
    forbidden.assert_not_called()
    from pathlib import Path
    assert created and all(not Path(directory).exists() for directory in created)


def test_progress_business_error_is_readable_and_never_sends_file(client):
    course = create_course(client)
    response = client.post(f"/api/courses/{course['id']}/teaching-calendar?progress=1")
    data = events(response)
    assert len(data) == 1 and data[0]["type"] == "error"
    assert data[0]["status"] == 400 and data[0]["code"] == "calendar_source_missing"
    assert "从大纲建课" in data[0]["message"]
    assert client.get("/healthz").status_code == 200


def test_progress_keeps_owner_origin_revision_and_format_protection(client, app):
    course = create_course(client, calendar_data())
    url = f"/api/courses/{course['id']}/teaching-calendar?progress=1"
    with TestClient(app) as anonymous:
        assert anonymous.post(url).status_code == 401
    assert client.post(url, headers={"X-Course-Revision":"stale"}).status_code == 409
    assert client.post(url, headers={"Origin":"https://other.example"}).status_code == 403
    assert client.post(url + "&format=xlsx").status_code == 400
    assert client.post("/api/register", json={"username":"progress-other", "password":PASSWORD, "invite_code":"TST123"}).status_code == 200
    assert client.post("/api/login", json={"username":"progress-other", "password":PASSWORD}).status_code == 200
    assert client.post(url).status_code == 404


@pytest.mark.parametrize("unexpected", [False,True])
def test_stream_conversion_failure_is_sanitized_and_releases_request(client, monkeypatch, unexpected):
    course = create_course(client, calendar_data())
    def broken(name, settings, *, progress, teaching_assignment=None):
        progress(15,"正在准备模板")
        if unexpected: raise RuntimeError("private-internal-detail")
        raise CalendarError("转换超时，请重试。",504,code="conversion_timeout")
    monkeypatch.setattr(progress,"generate_for_course",broken)
    response = client.post(f"/api/courses/{course['id']}/teaching-calendar?progress=1")
    data = events(response)
    assert data[0]["type"] == "progress" and data[-1]["type"] == "error"
    assert all(event["type"] not in {"complete","file","chunk"} for event in data)
    assert "private-internal-detail" not in response.text
    from web_app.resource_limits import admit_request, release_request
    assert admit_request(1)
    release_request(1)


def test_initial_progress_and_heartbeat_arrive_before_conversion_finishes(monkeypatch):
    gate = threading.Event()
    cleaned = threading.Event()
    def waiting(name, settings, *, progress, teaching_assignment=None):
        progress(15,"等待模板转换")
        try:
            assert gate.wait(3)
            return "合成.xls", b"synthetic", {"hours_match":True}
        finally:
            cleaned.set()
    monkeypatch.setattr(progress,"generate_for_course",waiting)
    monkeypatch.setattr(progress,"HEARTBEAT_SECONDS",0.02)
    async def scenario():
        stream = progress.calendar_events("合成",{})
        first = json.loads(await anext(stream))
        assert first["percent"] == 15 and not cleaned.is_set()
        heartbeat = json.loads(await anext(stream))
        assert heartbeat["type"] == "heartbeat" and not cleaned.is_set()
        gate.set()
        remaining = [json.loads(line) async for line in stream]
        assert remaining[-1]["type"] == "complete"
    asyncio.run(scenario())
    assert cleaned.is_set()


def test_closed_stream_waits_for_worker_cleanup(monkeypatch):
    gate = threading.Event()
    cleaned = threading.Event()
    def waiting(name, settings, *, progress, teaching_assignment=None):
        progress(15,"处理中")
        try:
            assert gate.wait(3)
            return "合成.xls", b"synthetic", {}
        finally:
            cleaned.set()
    monkeypatch.setattr(progress,"generate_for_course",waiting)
    async def scenario():
        stream = progress.calendar_events("合成",{})
        assert json.loads(await anext(stream))["type"] == "progress"
        closing = asyncio.create_task(stream.aclose())
        await asyncio.sleep(0.01)
        assert not closing.done() and not cleaned.is_set()
        gate.set()
        await asyncio.wait_for(closing,1)
        assert cleaned.is_set()
    asyncio.run(scenario())
