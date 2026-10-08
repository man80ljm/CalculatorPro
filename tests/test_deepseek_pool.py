"""DeepSeek 账号池：假密钥，不访问真实接口。"""
import threading
import time

import pytest
import requests

from web_app.deepseek_pool import (
    PoolError,
    QueueTimeout,
    account_loads,
    acquire,
    chat_completions,
    queue_label,
    redact_text,
    report_text,
    reset_pool,
)


def _clear_key_env(monkeypatch):
    for name in (
        "DEEPSEEK_API_KEY",
        "DEEPSEEK_KEYS_A",
        "DEEPSEEK_KEYS_B",
        "DEEPSEEK_KEYS_A_FILE",
        "DEEPSEEK_KEYS_B_FILE",
        "DEEPSEEK_ACCOUNT_CONCURRENCY",
        "DEEPSEEK_QUEUE_REPORT_SECONDS",
        "DEEPSEEK_QUEUE_SYLLABUS_SECONDS",
    ):
        monkeypatch.delenv(name, raising=False)


def _two_accounts(monkeypatch, concurrency="5"):
    _clear_key_env(monkeypatch)
    monkeypatch.setenv("DEEPSEEK_KEYS_A", "sk-test-a1,sk-test-a2")
    monkeypatch.setenv("DEEPSEEK_KEYS_B", "sk-test-b1")
    monkeypatch.setenv("DEEPSEEK_ACCOUNT_CONCURRENCY", concurrency)
    reset_pool()


def _wait_until(predicate, timeout=2):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if predicate():
            return
        time.sleep(0.01)
    raise AssertionError("等待超时")


class _Resp:
    def __init__(self, status, content="ok", body_text=""):
        self.status_code = status
        self.text = body_text
        self._content = content

    def json(self):
        return {
            "choices": [{"message": {"content": self._content}}],
            "usage": {"prompt_tokens": 2, "completion_tokens": 1, "total_tokens": 3},
        }


def test_equal_idle_accounts_round_robin(monkeypatch):
    _two_accounts(monkeypatch)
    names = []
    for _ in range(4):
        lease = acquire("report")
        names.append(lease.account_name)
        lease.release()
    assert names == ["A", "B", "A", "B"]
    assert account_loads() == {"A": 0, "B": 0}


def test_prefers_less_busy_account(monkeypatch):
    _two_accounts(monkeypatch)
    first = acquire("report")
    assert first.account_name == "A"
    second = acquire("syllabus")
    assert second.account_name == "B"
    third = acquire("report")
    assert third.account_name == "B"
    fourth = acquire("syllabus")
    assert fourth.account_name == "A"
    assert account_loads() == {"A": 2, "B": 2}
    for lease in (first, second, third, fourth):
        lease.release()
    assert account_loads() == {"A": 0, "B": 0}


def test_switches_key_on_429_401_403_timeout_and_5xx(monkeypatch):
    _clear_key_env(monkeypatch)
    monkeypatch.setenv("DEEPSEEK_KEYS_A", "sk-test-a1,sk-test-a2")
    monkeypatch.setenv("DEEPSEEK_ACCOUNT_CONCURRENCY", "2")
    reset_pool()
    def raise_timeout():
        raise requests.Timeout("slow")

    cases = [
        ("429", lambda: _Resp(429, body_text="nope")),
        ("401", lambda: _Resp(401, body_text="nope")),
        ("403", lambda: _Resp(403, body_text="nope")),
        ("500", lambda: _Resp(500, body_text="nope")),
        ("timeout", raise_timeout),
    ]
    for name, fail in cases:
        calls = []

        def fake_post(url, headers=None, json=None, timeout=None, _fail=fail, _calls=calls):
            _calls.append(headers["Authorization"])
            if len(_calls) == 1:
                return _fail()
            return _Resp(200, content=f"done-{name}")

        monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
        result = chat_completions("syllabus", [{"role": "user", "content": name}], max_tokens=20, timeout=5)
        assert result["content"] == f"done-{name}"
        assert calls == ["Bearer sk-test-a1", "Bearer sk-test-a2"]
        assert result["prompt_tokens"] == 2
        assert "sk-test" not in result["content"]
    assert account_loads()["A"] == 0


def test_stops_after_one_round_and_does_not_try_the_other_account(monkeypatch):
    _two_accounts(monkeypatch, concurrency="2")
    calls = []

    def fake_post(url, headers=None, json=None, timeout=None):
        calls.append(headers["Authorization"])
        return _Resp(500, body_text="sk-test-a1-secret")

    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    with pytest.raises(PoolError) as exc:
        chat_completions("report", [{"role": "user", "content": "x"}], timeout=5, temperature=0.7)
    assert calls == ["Bearer sk-test-a1", "Bearer sk-test-a2"]
    text = str(exc.value)
    assert "sk-test-a1" not in text
    assert "sk-test-a2" not in text
    assert "sk-test-b1" not in text
    assert "暂时不可用" in text
    assert account_loads() == {"A": 0, "B": 0}


def test_client_error_does_not_switch_key(monkeypatch):
    _two_accounts(monkeypatch)
    calls = []

    def fake_post(url, headers=None, json=None, timeout=None):
        calls.append(headers["Authorization"])
        return _Resp(400, body_text="sk-test-a1")

    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    with pytest.raises(PoolError) as exc:
        chat_completions("syllabus", [{"role": "user", "content": "x"}], timeout=5)
    assert calls == ["Bearer sk-test-a1"]
    assert "sk-test-a1" not in str(exc.value)
    assert account_loads()["A"] == 0


def test_error_text_hides_key_fragments(monkeypatch):
    _clear_key_env(monkeypatch)
    secret = "sk-test-a1-secret"
    monkeypatch.setenv("DEEPSEEK_KEYS_A", secret)
    reset_pool()

    def fake_post(url, headers=None, json=None, timeout=None):
        raise RuntimeError(f"rejected {secret} fragment {secret[8:]}")

    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    with pytest.raises(PoolError) as exc:
        chat_completions("syllabus", [{"role": "user", "content": "x"}], timeout=5)
    text = str(exc.value)
    assert secret not in text
    assert secret[8:] not in text
    assert "Bearer" not in text
    assert redact_text(f"see {secret}") != f"see {secret}"
    assert secret not in redact_text(f"see {secret}")


def test_full_pool_queues_and_reports_go_first(monkeypatch):
    _two_accounts(monkeypatch, concurrency="1")
    monkeypatch.setenv("DEEPSEEK_QUEUE_REPORT_SECONDS", "30")
    monkeypatch.setenv("DEEPSEEK_QUEUE_SYLLABUS_SECONDS", "15")
    held = [acquire("report"), acquire("syllabus")]
    assert account_loads() == {"A": 1, "B": 1}

    order = []
    seen = []
    started = threading.Event()
    got_report = threading.Event()
    release_report = threading.Event()

    def on_syllabus(position, eta):
        seen.append(("syllabus", position, eta))
        started.set()

    def on_report(position, eta):
        seen.append(("report", position, eta))

    def wait_syllabus():
        lease = acquire("syllabus", on_queue=on_syllabus, wait_timeout=3)
        order.append("syllabus")
        lease.release()

    def wait_report():
        lease = acquire("report", on_queue=on_report, wait_timeout=3)
        order.append("report")
        got_report.set()
        assert release_report.wait(3)
        lease.release()

    syllabus_thread = threading.Thread(target=wait_syllabus)
    report_thread = threading.Thread(target=wait_report)
    syllabus_thread.start()
    assert started.wait(2)
    report_thread.start()
    _wait_until(lambda: any(item[0] == "report" for item in seen))
    assert ("report", 1, 30) in seen
    assert any(item[0] == "syllabus" and item[1] == 2 and item[2] == 30 for item in seen)
    assert queue_label(2, 30) == "前面还有 2 人，大约再等 30 秒"

    held[0].release()
    assert got_report.wait(2)
    assert order == ["report"]
    release_report.set()
    held[1].release()
    syllabus_thread.join(3)
    report_thread.join(3)
    assert order == ["report", "syllabus"]
    assert account_loads() == {"A": 0, "B": 0}


def test_queue_eta_is_position_times_configured_seconds(monkeypatch):
    _clear_key_env(monkeypatch)
    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-test-a1")
    monkeypatch.setenv("DEEPSEEK_ACCOUNT_CONCURRENCY", "1")
    monkeypatch.setenv("DEEPSEEK_QUEUE_REPORT_SECONDS", "7")
    reset_pool()
    held = acquire("report")
    positions = []

    def on_queue(position, eta):
        positions.append((position, eta))

    outcome = {}

    def blocked():
        try:
            acquire("report", on_queue=on_queue, wait_timeout=0.2)
        except QueueTimeout as exc:
            outcome["text"] = str(exc)
            outcome["position"] = exc.position
            outcome["eta"] = exc.eta_seconds

    thread = threading.Thread(target=blocked)
    thread.start()
    _wait_until(lambda: positions)
    thread.join(2)
    assert positions[0] == (1, 7)
    assert outcome["position"] == 1
    assert outcome["eta"] == 7
    assert outcome["text"] == "前面还有 1 人，大约再等 7 秒"
    assert "sk-test-a1" not in outcome["text"]
    held.release()


def test_single_api_key_stays_compatible(monkeypatch):
    _clear_key_env(monkeypatch)
    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-test-a1")
    reset_pool()
    calls = []

    def fake_post(url, headers=None, json=None, timeout=None):
        calls.append({"auth": headers["Authorization"], "temperature": json["temperature"], "model": json["model"]})
        return _Resp(200, content="报告段落")

    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    monkeypatch.setenv("DEEPSEEK_MODEL", "deepseek-flash")
    lease = acquire("report")
    try:
        assert lease.account_name == "default"
        text = report_text(lease.account, [{"role": "user", "content": "写一段"}], 30, 5)
    finally:
        lease.release()
    assert text == "报告段落"
    assert calls == [{"auth": "Bearer sk-test-a1", "temperature": 0.7, "model": "deepseek-flash"}]
    assert account_loads() == {"default": 0}

    monkeypatch.setenv("DEEPSEEK_KEYS_A", "sk-test-a1")
    reset_pool()
    lease = acquire("syllabus")
    assert lease.account_name == "A"
    lease.release()


def test_keys_file_skips_blanks_and_comments(monkeypatch, tmp_path):
    _clear_key_env(monkeypatch)
    path = tmp_path / "keys-a.txt"
    path.write_text("# comment\n\nsk-test-a1\n  sk-test-a2  \n", encoding="utf-8")
    monkeypatch.setenv("DEEPSEEK_KEYS_A_FILE", str(path))
    monkeypatch.setenv("DEEPSEEK_KEYS_B", "sk-test-b1")
    reset_pool()
    calls = []

    def fake_post(url, headers=None, json=None, timeout=None):
        calls.append((headers["Authorization"], json["temperature"]))
        if headers["Authorization"].endswith("sk-test-a1"):
            return _Resp(429)
        return _Resp(200, content="大纲")

    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    result = chat_completions("syllabus", [{"role": "user", "content": "大纲"}], timeout=5)
    assert result["content"] == "大纲"
    assert calls[0][0] == "Bearer sk-test-a1"
    assert calls[1] == ("Bearer sk-test-a2", 0.1)
    assert "sk-test" not in str(result)

    inline = tmp_path / "inline-keys.txt"
    inline.write_text("# note\nsk-test-a1\n", encoding="utf-8")
    monkeypatch.delenv("DEEPSEEK_KEYS_A_FILE", raising=False)
    monkeypatch.setenv("DEEPSEEK_KEYS_A", str(inline))
    reset_pool()
    calls.clear()

    def inline_post(url, headers=None, json=None, timeout=None):
        calls.append(headers["Authorization"])
        return _Resp(200, content="ok")

    monkeypatch.setattr("core_app.ai_report.requests.post", inline_post)
    again = chat_completions("syllabus", [{"role": "user", "content": "x"}], timeout=5)
    assert again["content"] == "ok"
    assert calls == ["Bearer sk-test-a1"]


def test_syllabus_queue_timeout_returns_friendly_text(monkeypatch):
    from io import BytesIO

    from docx import Document

    from web_app.service import ServiceError
    from web_app.syllabus.client import complete as real_complete
    from web_app.syllabus.service import extract_draft

    _clear_key_env(monkeypatch)
    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-test-a1")
    monkeypatch.setenv("DEEPSEEK_ACCOUNT_CONCURRENCY", "1")
    monkeypatch.setenv("DEEPSEEK_QUEUE_SYLLABUS_SECONDS", "15")
    reset_pool()
    held = acquire("report")

    def fast(messages, max_tokens=8000, timeout=120):
        return real_complete(messages, max_tokens=max_tokens, timeout=0.3)

    monkeypatch.setattr("web_app.syllabus.service.complete", fast)
    document = Document()
    document.add_paragraph("课程目标1：能完成作品")
    buffer = BytesIO()
    document.save(buffer)
    try:
        with pytest.raises(ServiceError) as caught:
            extract_draft("syllabus.docx", buffer.getvalue())
    finally:
        held.release()
    assert caught.value.status == 503
    assert str(caught.value) == "前面还有 1 人，大约再等 15 秒"
    assert "sk-test" not in str(caught.value)
