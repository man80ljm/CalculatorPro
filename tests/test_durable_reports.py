"""排队上限、数据库恢复和 AI 检查点复用。所有输入为合成数据。"""
import json
from datetime import timedelta
from sqlalchemy import select
from tests.test_report_job import app, client, _prepare_course, _wait_job
from web_app.db import ReportTask, session_scope, utcnow
from web_app.report_jobs import JobRuntime, stop_job_runtime


def test_queue_limit_and_repeated_click_reuses_task(client, monkeypatch):
    stop_job_runtime(force=True)
    monkeypatch.setenv("REPORT_WORKER_MODE", "external")
    courses = [_prepare_course(client, f"合成课程{i}") for i in range(5)]
    ids = []
    for course in courses[:4]:
        response = client.post(f"/api/courses/{course}/report-jobs")
        assert response.status_code == 200, response.text
        ids.append(response.json()["job_id"])
    repeated = client.post(f"/api/courses/{courses[0]}/report-jobs")
    assert repeated.json()["job_id"] == ids[0]
    assert client.post(f"/api/courses/{courses[4]}/report-jobs").status_code == 429
    with session_scope() as db:
        rows = list(db.scalars(select(ReportTask)))
        assert len(rows) == 4
        assert all(row.status == "queued" for row in rows)
    current = client.get(f"/api/courses/{courses[0]}/report-jobs/current").json()
    assert current["job"]["job_id"] == ids[0]


def test_expired_running_task_recovers_and_reuses_cached_ai(client, monkeypatch):
    stop_job_runtime(force=True)
    monkeypatch.setenv("REPORT_WORKER_MODE", "external")
    monkeypatch.setenv("DEEPSEEK_API_KEY", "sk-test-only")
    course = _prepare_course(client)
    started = client.post(f"/api/courses/{course}/report-jobs").json()
    job_id = started["job_id"]
    with session_scope() as db:
        task = db.get(ReportTask, job_id)
        task.status = "running"
        task.claim_token = "dead-worker"
        task.lease_until = utcnow() - timedelta(seconds=1)
    from web_app.storage import upload_root
    folder = upload_root() / ".jobs" / job_id
    (folder / "answers.json").write_text(json.dumps(["合成总体分析", "合成目标一分析", "合成目标二分析"]), encoding="utf-8")
    def forbidden(*args, **kwargs):
        raise AssertionError("已缓存分析不能再次调用付费 API")
    monkeypatch.setattr("web_app.service.generate_answers", forbidden)
    runtime = JobRuntime()
    try:
        done, _ = _wait_job(client, job_id)
        assert done["done"], done
        assert client.get(f"/api/report-jobs/{job_id}/download").status_code == 200
        from web_app.report_jobs import _JOBS
        assert _JOBS[job_id].content == b""
        with session_scope() as db:
            task = db.get(ReportTask, job_id)
            assert task.archive_file_id and task.active_key is None
        # 调度器再启动时不能生成第二份同任务资料。
        runtime.tick()
    finally:
        runtime.stop()


def test_unexpired_lease_is_not_stolen(client, monkeypatch):
    stop_job_runtime(force=True)
    monkeypatch.setenv("REPORT_WORKER_MODE", "external")
    course = _prepare_course(client)
    started = client.post(f"/api/courses/{course}/report-jobs").json()
    with session_scope() as db:
        task = db.get(ReportTask, started["job_id"])
        task.status = "running"
        task.claim_token = "another-live-worker"
        task.lease_until = utcnow() + timedelta(minutes=5)
    runtime = JobRuntime()
    try:
        runtime.tick()
        with session_scope() as db:
            assert db.get(ReportTask, started["job_id"]).claim_token == "another-live-worker"
        assert not runtime.active
    finally:
        runtime.stop()
