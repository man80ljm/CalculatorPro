"""单次建课请求中的日历整理、后台配置复用与规则回退。网络请求全部模拟。"""
import copy
import io
import json
import urllib.error
from contextlib import contextmanager
from unittest.mock import Mock

import docx
import pytest
import xlrd
from sqlalchemy.orm import Session

from tests.test_web_flow import app, client
from tests.test_syllabus_rules import _ai
from tests.test_teaching_calendar import outline_bytes
from tests.test_course_calendar import (calendar_data, create_course, download,
                                       cached_template, converted_template)
from web_app.db import AdminApiConfiguration, get_engine
from web_app.syllabus import client as ai_client, service as syllabus_service
from web_app.teaching.api_settings import ApiSettings
from web_app.teaching_calendar import engine
from web_app.teaching_calendar.source import apply_calendar_ai, normalize_calendar_data

KEY = "sk-private-calendar-offline-fixture"


def ai_calendar():
    return {"教材": {"教材名称":"程序设计基础", "版次":"第3版", "编者":"张三", "出版社":"高等教育出版社"},
            "lessons": [{"课次": n, "教学内容": f"1.学习知识{n}；2.完成实践。",
                         "教学要求": f"1.理解知识{n}；2.能够完成实践。",
                         "学时":999, "教材章节":"AI 不得改写章节", "备注":"AI 不得改写备注"}
                        for n in (1, 2)]}


def result(calendar=None):
    data = _ai()
    if calendar is not None:
        data["teaching_calendar"] = calendar
    return {"content":json.dumps(data, ensure_ascii=False), "model":"offline-model", "prompt_tokens":12, "completion_tokens":8}


def extract(client):
    return client.post("/api/syllabus/extract", files={"syllabus":("大纲.docx", outline_bytes())})


def test_one_ai_call_refines_and_persists_calendar_without_changing_source_fields(client, monkeypatch, cached_template):
    monkeypatch.setenv("DEEPSEEK_API_KEY", KEY)
    complete = Mock(return_value=result(ai_calendar()))
    monkeypatch.setattr(syllabus_service, "complete", complete)
    response = extract(client)
    assert response.status_code == 200, response.text
    assert complete.call_count == 1
    messages = complete.call_args.args[0]
    assert engine.SYSTEM_PROMPT in messages[0]["content"]
    sent = json.loads(messages[1]["content"].split("附加教学日历原文（仅供 teaching_calendar 整理）：\n")[1])
    assert "学时" not in sent["lessons"][0] and "备注" not in sent["lessons"][0]
    assert "掌握知识1" in sent["lessons"][0]["教学目标及教学内容"]
    calendar = response.json()["teaching_calendar"]
    assert calendar["mode"] == "ai" and calendar["warnings"] == []
    original = calendar_data()
    for old, new in zip(original["lessons"], calendar["lessons"]):
        for key in ("课次","学时","教材章节","备注"):
            assert new[key] == old[key]
    course = create_course(client, calendar)
    assert course["settings"]["teaching_calendar"] == calendar
    complete.side_effect = AssertionError("Download must reuse saved calendar, not call AI")
    file = download(client, course)
    assert file.status_code == 200, file.text[:1000]
    workbook = xlrd.open_workbook(file_contents=file.content)
    ws = workbook.sheet_by_index(0)
    assert ws.cell_value(7,3) == "1.学习知识1；2.完成实践。"
    assert ws.cell(7,1).ctype == xlrd.XL_CELL_NUMBER and ws.cell_value(7,1) == 2
    assert complete.call_count == 1
    workbook.release_resources()


def test_missing_calendar_response_keeps_original_rules_and_allows_creation(client, monkeypatch):
    monkeypatch.setenv("DEEPSEEK_API_KEY", KEY)
    complete = Mock(return_value=result())
    monkeypatch.setattr(syllabus_service, "complete", complete)
    response = extract(client)
    assert response.status_code == 200
    calendar = response.json()["teaching_calendar"]
    assert calendar["mode"] == "rule" and "按规则保存" in calendar["warnings"][0]
    assert calendar["lessons"] == calendar_data()["lessons"]
    assert create_course(client, calendar)["settings"]["teaching_calendar"] == calendar
    assert complete.call_count == 1


def test_partial_calendar_only_falls_back_for_missing_lesson():
    original = calendar_data()
    data = ai_calendar()
    data["lessons"] = data["lessons"][:1]
    refined = apply_calendar_ai(original, data)
    assert refined["mode"] == "ai_partial" and refined["warnings"]
    assert refined["lessons"][0]["教学内容"] == "1.学习知识1；2.完成实践。"
    assert refined["lessons"][1] == original["lessons"][1]


@pytest.mark.parametrize("issue", ["unknown_id", "duplicate_id", "boolean_id", "bad_text", "bad_book", "control_char"])
def test_invalid_ai_calendar_is_safely_replaced_with_rules(issue):
    original = calendar_data()
    data = ai_calendar()
    if issue == "unknown_id": data["lessons"][0]["课次"] = 99
    elif issue == "duplicate_id": data["lessons"][1]["课次"] = 1
    elif issue == "boolean_id": data["lessons"][0]["课次"] = True
    elif issue == "bad_text": data["lessons"][0]["教学内容"] = {"nested":"invalid"}
    elif issue == "bad_book": data["教材"]["教材名称"] = {"nested":"invalid"}
    elif issue == "control_char": data["lessons"][0]["教学内容"] = "invalid\x00content"
    refined = apply_calendar_ai(original, data)
    assert refined["mode"] == "rule" and refined["warnings"]
    assert refined["lessons"] == original["lessons"]


def test_overlong_ai_text_preserves_rule_lesson():
    original = calendar_data()
    data = ai_calendar()
    data["lessons"][0]["教学内容"] = "长" * 81
    refined = apply_calendar_ai(original, data)
    assert refined["mode"] == "ai_partial"
    assert refined["lessons"][0] == original["lessons"][0]
    assert refined["lessons"][1]["教学内容"] == "1.学习知识2；2.完成实践。"


def test_absent_textbook_source_cannot_be_invented_by_ai():
    original = calendar_data()
    original["textbook_source"] = ""
    original["教材"] = {key:"" for key in engine.BOOK_CELLS}
    assert apply_calendar_ai(original, ai_calendar())["教材"] == original["教材"]


def test_invalid_mode_is_uniform_calendar_error():
    original = calendar_data()
    original["mode"] = []
    with pytest.raises(engine.CalendarError):
        normalize_calendar_data(original)


class OfflineResponse:
    def __init__(self, body): self.body = body
    def read(self, size): return self.body[:size]
    def close(self): pass


def configured_settings():
    ApiSettings(get_engine()).save({"source":"custom", "key":KEY, "model":"configured-calendar-model"})


def test_saved_admin_key_and_model_are_used_without_environment_key(client, monkeypatch):
    configured_settings()
    calls = []
    class OfflineOpener:
        def __init__(self, config, *, wait_timeout):
            assert config["key"] == KEY and config["source"] == "custom"
            assert wait_timeout <= 10
        @contextmanager
        def open(self, request, timeout):
            calls.append(json.loads(request.data))
            assert request.full_url == "https://api.deepseek.com/chat/completions"
            assert timeout == 120
            yield OfflineResponse(json.dumps({"choices":[{"finish_reason":"stop","message":result(ai_calendar())}],
                "usage":{"prompt_tokens":12,"completion_tokens":8}}).encode())
    monkeypatch.setattr(ai_client, "ProviderOpener", OfflineOpener)
    response = extract(client)
    assert response.status_code == 200, response.text
    assert len(calls) == 1 and calls[0]["model"] == "configured-calendar-model"
    assert response.json()["teaching_calendar"]["mode"] == "ai"
    assert KEY not in response.text and KEY not in json.dumps(calls)
    with Session(get_engine()) as db:
        assert KEY not in db.get(AdminApiConfiguration,1).encrypted_config


@pytest.mark.parametrize("failure", ["auth", "timeout", "invalid_json"])
def test_saved_config_call_errors_never_leak_key(client, monkeypatch, failure):
    configured_settings()
    class BrokenOpener:
        def __init__(self, config, *, wait_timeout): pass
        @contextmanager
        def open(self, request, timeout):
            if failure == "auth":
                raise urllib.error.HTTPError(request.full_url,401,KEY,None,io.BytesIO(KEY.encode()))
            if failure == "timeout":
                raise OSError(KEY)
            yield OfflineResponse(KEY.encode())
    monkeypatch.setattr(ai_client, "ProviderOpener", BrokenOpener)
    response = extract(client)
    assert response.status_code == 502 and KEY not in response.text
    assert client.get("/healthz").status_code == 200


def test_saved_project_model_uses_existing_pool_and_releases_slot(client, monkeypatch):
    monkeypatch.setenv("DEEPSEEK_API_KEY", KEY)
    ApiSettings(get_engine()).save({"source":"project", "model":"saved-project-model"})
    seen = []
    class OfflineHttp:
        def open(self, request, timeout):
            seen.append((request.get_header("Authorization"), json.loads(request.data)))
            return OfflineResponse(json.dumps({"choices":[{"finish_reason":"stop", "message":result(ai_calendar())}]}).encode())
    monkeypatch.setattr("web_app.teaching.provider.urllib.request.build_opener", lambda *args:OfflineHttp())
    response = extract(client)
    assert response.status_code == 200, response.text
    assert seen[0][0] == "Bearer " + KEY and seen[0][1]["model"] == "saved-project-model"
    assert KEY not in response.text
    from web_app.deepseek_pool import get_pool
    pool = get_pool()
    assert pool.external_inflight == 0 and all(account.inflight == 0 for account in pool.accounts)


def test_project_configuration_retries_next_key_without_exposing_either(client, monkeypatch):
    second = "sk-calendar-second-offline-fixture"
    monkeypatch.setenv("DEEPSEEK_KEYS_A", KEY + "," + second)
    ApiSettings(get_engine()).save({"source":"project", "model":"saved-project-model"})
    seen = []
    class OfflineHttp:
        def open(self, request, timeout):
            authorization = request.get_header("Authorization")
            seen.append(authorization)
            if authorization == "Bearer " + KEY:
                raise urllib.error.HTTPError(request.full_url,401,KEY,None,io.BytesIO(KEY.encode()))
            return OfflineResponse(json.dumps({"choices":[{"finish_reason":"stop", "message":result(ai_calendar())}]}).encode())
    monkeypatch.setattr("web_app.teaching.provider.urllib.request.build_opener",lambda *args:OfflineHttp())
    response = extract(client)
    assert response.status_code == 200, response.text
    assert seen == ["Bearer " + KEY, "Bearer " + second]
    assert KEY not in response.text and second not in response.text


def test_waiting_for_syllabus_ai_does_not_block_other_requests(client, monkeypatch):
    import threading
    import time
    from concurrent.futures import ThreadPoolExecutor
    monkeypatch.setenv("DEEPSEEK_API_KEY",KEY)
    started = threading.Event()
    release = threading.Event()
    def delayed(*args, **kwargs):
        started.set()
        assert release.wait(5)
        return result(ai_calendar())
    monkeypatch.setattr(syllabus_service,"complete",delayed)
    with ThreadPoolExecutor(max_workers=1) as worker:
        pending = worker.submit(extract,client)
        assert started.wait(5)
        begin = time.monotonic()
        try:
            assert client.get("/healthz").status_code == 200
            assert time.monotonic() - begin < 1
        finally:
            release.set()
        assert pending.result(timeout=5).status_code == 200


def test_ai_explicitly_missing_book_field_clears_rule_guess():
    original = calendar_data()
    original["教材"]["编者"] = "规则误识别的片段"
    data = ai_calendar()
    data["教材"]["编者"] = ""
    assert apply_calendar_ai(original,data)["教材"]["编者"] == ""
