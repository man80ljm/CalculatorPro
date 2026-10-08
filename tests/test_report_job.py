"""报告任务：真实阶段顺序、压缩包只生成一次，并且只给发起人下载。"""
import io
import threading
import time
import zipfile

import pytest
from fastapi.testclient import TestClient

from tests.test_web_flow import PASSWORD, _course, _fill_forward_template, _settings
from web_app.limiter import login_limiter, login_user_limiter, register_limiter, reset_limiters
from web_app.report_jobs import reset_report_jobs


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setenv("SECRET_KEY", "test-secret-key")
    db_path = (tmp_path / "app.db").resolve()
    monkeypatch.setenv("DATABASE_URL", "sqlite:///" + db_path.as_posix())
    monkeypatch.setenv("UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("COOKIE_SECURE", "false")
    monkeypatch.setenv("DEEPSEEK_API_KEY", "")
    reset_limiters()
    reset_report_jobs()
    login_limiter.max_calls = 8
    login_user_limiter.max_calls = 8
    register_limiter.max_calls = 8
    from web_app.app import create_app

    return create_app()


@pytest.fixture
def client(app):
    with TestClient(app) as test_client:
        created = test_client.post("/api/register", json={"username": "teacher", "password": PASSWORD})
        assert created.status_code == 200, created.text
        logged = test_client.post("/api/login", json={"username": "teacher", "password": PASSWORD})
        assert logged.status_code == 200, logged.text
        yield test_client


def _prepare_course(client, name="测试课程") -> int:
    course_id = _course(client, _settings(), name=name)
    template = client.post(f"/api/courses/{course_id}/template")
    assert template.status_code == 200, template.text
    excel_bytes = _fill_forward_template(template.content)
    uploaded = client.post(
        f"/api/courses/{course_id}/files",
        data={"kind": "grade"},
        files={"file": ("grades.xlsx", excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert uploaded.status_code == 200, uploaded.text
    return course_id


def _wait_job(client, job_id: str, timeout: float = 90):
    deadline = time.time() + timeout
    percents = []
    last = None
    while time.time() < deadline:
        response = client.get(f"/api/report-jobs/{job_id}")
        assert response.status_code == 200, response.text
        last = response.json()
        percents.append(last["percent"])
        if last["done"] or last["error"]:
            return last, percents
        time.sleep(0.05)
    raise AssertionError(f"报告任务超时：{last}")


def _assert_monotonic(values) -> None:
    assert values == sorted(values)
    assert all(values[index] <= values[index + 1] for index in range(len(values) - 1))


class _AiResponse:
    def raise_for_status(self):
        return None

    def json(self):
        return {
            "choices": [
                {
                    "message": {
                        "content": "本轮课程目标达成情况总体平稳，考核覆盖了主要能力点，后续将补充针对性练习。"
                    }
                }
            ]
        }


def test_report_job_stage_order_zip_privacy_and_single_flight(client, app, monkeypatch):
    from core_app.excel_calc import ExcelCalcMixin

    course_id = _prepare_course(client)
    grade_runs = []
    original = ExcelCalcMixin.process_forward_grades

    def wrapped(self, *args, **kwargs):
        grade_runs.append(self.input_file)
        return original(self, *args, **kwargs)

    prompts = []
    entered = threading.Event()
    release = threading.Event()

    def pause(stage):
        if stage == "ai":
            entered.set()
            assert release.wait(60), "测试没有放开 AI 阶段"

    def fake_post(url, headers=None, json=None, timeout=None):
        prompts.append((json or {}).get("messages") or [])
        return _AiResponse()

    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-test-secret")
    monkeypatch.setattr(ExcelCalcMixin, "process_forward_grades", wrapped)
    monkeypatch.setattr("web_app.report_jobs.pause_before_stage", pause)
    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)

    job_id = ""
    try:
        started = client.post(f"/api/courses/{course_id}/report-jobs")
        assert started.status_code == 200, started.text
        job_id = started.json()["job_id"]
        assert entered.wait(60), started.text
        mid = client.get(f"/api/report-jobs/{job_id}")
        assert mid.status_code == 200, mid.text
        mid_body = mid.json()
        assert mid_body["done"] is False
        assert mid_body["stage"] == "ai"
        assert [item["stage"] for item in mid_body["stages"]] == ["calculate", "tables", "ai"]
        _assert_monotonic([item["percent"] for item in mid_body["stages"]])

        again = client.post(f"/api/courses/{course_id}/report-jobs")
        assert again.status_code == 200, again.text
        assert again.json()["job_id"] == job_id

        with TestClient(app) as other:
            assert other.post("/api/register", json={"username": "other-teacher", "password": PASSWORD}).status_code == 200
            assert other.post("/api/login", json={"username": "other-teacher", "password": PASSWORD}).status_code == 200
            assert other.get(f"/api/report-jobs/{job_id}").status_code == 404
            assert other.get(f"/api/report-jobs/{job_id}/download").status_code == 404

        release.set()
        body, percents = _wait_job(client, job_id)
        assert body["error"] == ""
        assert body["done"] is True
        assert body["failed_stage"] == ""
        assert [item["stage"] for item in body["stages"]] == ["calculate", "tables", "ai", "package"]
        _assert_monotonic([item["percent"] for item in body["stages"]])
        _assert_monotonic(percents)
        assert body["summary"]["student_count"] == 2
        assert "总达成度" in body["summary"]["achievement"]
        from web_app.db import Course, session_scope
        from web_app.terms import current_term, parse_settings

        with session_scope() as db:
            term = current_term(db, db.get(Course, course_id))
            stored = parse_settings(term.settings_json).get("last_achievement") or {}
        achievement = body["summary"]["achievement"]
        assert stored["课程总目标"] == pytest.approx(achievement["总达成度"])
        for key, value in achievement.items():
            if key == "总达成度":
                continue
            assert stored[key] == pytest.approx(value)
        assert len(grade_runs) == 1

        downloaded = client.get(f"/api/report-jobs/{job_id}/download")
        assert downloaded.status_code == 200, downloaded.text
        assert downloaded.headers["content-type"].startswith("application/zip")
        archive = zipfile.ZipFile(io.BytesIO(downloaded.content))
        names = archive.namelist()
        assert names
        assert len(names) == len(set(names))
        assert not any(name.lower().endswith(".zip") for name in names)
        assert sum(name.endswith("5.基于考核结果的课程目标达成情况评价结果表.docx") for name in names) == 1
        assert sum("课程目标达成情况分析" in name and name.endswith(".docx") for name in names) == 1
        assert prompts, "AI 接口没有被调用"
        blob = "\n".join(str(item.get("content") or "") for message in prompts for item in message)
        assert "张三" not in blob
        assert "李四" not in blob
        assert "sk-test-secret" not in downloaded.content.decode("utf-8", "ignore")

        with TestClient(app) as other:
            assert other.post("/api/login", json={"username": "other-teacher", "password": PASSWORD}).status_code == 200
            assert other.get(f"/api/report-jobs/{job_id}").status_code == 404
            assert other.get(f"/api/report-jobs/{job_id}/download").status_code == 404
        saved = client.get(f"/api/courses/{course_id}/files").json()["files"]
        assert any(item["kind"] == "report" and item["original_name"].endswith(".zip") for item in saved)
    finally:
        release.set()
        if job_id:
            _wait_job(client, job_id)


def test_report_job_ai_failure_sets_chinese_failed_stage(client, monkeypatch):
    course_id = _prepare_course(client)

    def fake_post(url, headers=None, json=None, timeout=None):
        raise RuntimeError("connection reset")

    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-test-secret")
    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    started = client.post(f"/api/courses/{course_id}/report-jobs")
    assert started.status_code == 200, started.text
    job_id = started.json()["job_id"]
    body, percents = _wait_job(client, job_id)
    assert body["done"] is False
    assert body["failed_stage"] == "ai"
    assert body["error"]
    assert "失败" in body["error"]
    assert any("\u4e00" <= char <= "\u9fff" for char in body["error"])
    assert [item["stage"] for item in body["stages"]] == ["calculate", "tables", "ai"]
    _assert_monotonic([item["percent"] for item in body["stages"]])
    _assert_monotonic(percents)
    failed = client.get(f"/api/report-jobs/{job_id}/download")
    assert failed.status_code == 400
    assert "失败" in failed.json()["detail"]


def test_report_job_queue_exposes_position_and_single_flight(client, monkeypatch):
    from web_app.deepseek_pool import reset_pool

    for name in ("DEEPSEEK_KEYS_A", "DEEPSEEK_KEYS_B", "DEEPSEEK_KEYS_A_FILE", "DEEPSEEK_KEYS_B_FILE"):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-test-secret")
    monkeypatch.setenv("DEEPSEEK_ACCOUNT_CONCURRENCY", "1")
    monkeypatch.setenv("DEEPSEEK_QUEUE_REPORT_SECONDS", "30")
    reset_pool()

    course_a = _prepare_course(client, "课程甲")
    course_b = _prepare_course(client, "课程乙")
    entered = threading.Event()
    release = threading.Event()

    def pause(stage):
        if stage == "calculate":
            entered.set()
            assert release.wait(60), "测试没有放开计算阶段"

    def fake_post(url, headers=None, json=None, timeout=None):
        return _AiResponse()

    monkeypatch.setattr("web_app.report_jobs.pause_before_stage", pause)
    monkeypatch.setattr("core_app.ai_report.requests.post", fake_post)
    job_a = ""
    job_b = ""
    try:
        started = client.post(f"/api/courses/{course_a}/report-jobs")
        assert started.status_code == 200, started.text
        job_a = started.json()["job_id"]
        assert entered.wait(60), started.text
        holding = client.get(f"/api/report-jobs/{job_a}")
        assert holding.json()["stage"] == "calculate"
        assert holding.json()["queue_position"] is None

        queued = client.post(f"/api/courses/{course_b}/report-jobs")
        assert queued.status_code == 200, queued.text
        body = queued.json()
        job_b = body["job_id"]
        assert body["stage"] == "queued"
        assert body["queue_position"] == 1
        assert body["queue_eta_seconds"] == 30
        assert body["stage_label"] == "前面还有 1 人，大约再等 30 秒"
        assert "sk-test-secret" not in queued.text
        again = client.post(f"/api/courses/{course_b}/report-jobs")
        assert again.status_code == 200, again.text
        assert again.json()["job_id"] == job_b
        assert again.json()["stage"] == "queued"
    finally:
        release.set()
        monkeypatch.setattr("web_app.report_jobs.pause_before_stage", None)
        if job_a:
            _wait_job(client, job_a)
        if job_b:
            finished, _percents = _wait_job(client, job_b)
            assert finished["done"] is True
            assert finished["queue_position"] is None
            assert finished["queue_eta_seconds"] is None
            assert [item["stage"] for item in finished["stages"]] == ["calculate", "tables", "ai", "package"]
            assert "sk-test-secret" not in str(finished)
