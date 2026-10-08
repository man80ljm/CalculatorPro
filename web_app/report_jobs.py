"""内存里的报告任务。单进程 uvicorn 下按用户和学期隔离，完成后保留半小时。

成功或失败时另写一条 report_job_events，供只读运维后台统计。进行中的任务只留在内存里。
"""
from __future__ import annotations

import threading
import time
import uuid
from datetime import datetime, timezone

from web_app.service import ReportBundle, ServiceError, run_report_pipeline

KEEP_SECONDS = 30 * 60
STAGE_INFO = {
    "calculate": ("正在计算达成度…", 15),
    "tables": ("正在生成统计表…", 45),
    "ai": ("AI 正在撰写分析（约 30 秒）…", 75),
    "package": ("打包完成", 100),
}

# 测试可以在进入某个阶段时拦住后台线程，生产环境保持 None。
pause_before_stage = None

_LOCK = threading.Lock()
_JOBS: dict[str, "ReportJob"] = {}


class ReportJob:
    def __init__(self, user_id: int, course_id: int, term_id: int):
        self.id = uuid.uuid4().hex
        self.user_id = int(user_id)
        self.course_id = int(course_id)
        self.term_id = int(term_id)
        self.stage = "calculate"
        self.stage_label = STAGE_INFO["calculate"][0]
        self.percent = 0
        self.done = False
        self.running = True
        self.error = ""
        self.failed_stage = ""
        self.summary = None
        self.filename = ""
        self.content = b""
        self.stages: list[dict] = []
        self.queue_position = None
        self.queue_eta_seconds = None
        now = time.time()
        self.created_at = now
        self.updated_at = now

    def public(self) -> dict:
        return {
            "job_id": self.id,
            "stage": self.stage,
            "stage_label": self.stage_label,
            "percent": self.percent,
            "done": self.done,
            "error": self.error,
            "failed_stage": self.failed_stage,
            "summary": self.summary,
            "stages": [dict(item) for item in self.stages],
            "queue_position": self.queue_position,
            "queue_eta_seconds": self.queue_eta_seconds,
        }


def reset_report_jobs() -> None:
    from web_app.deepseek_pool import reset_pool

    reset_pool()
    with _LOCK:
        _JOBS.clear()


def _sweep_locked(now: float | None = None) -> None:
    now = time.time() if now is None else now
    stale = [
        job_id
        for job_id, job in _JOBS.items()
        if not job.running and now - job.updated_at > KEEP_SECONDS
    ]
    for job_id in stale:
        _JOBS.pop(job_id, None)


def _running_locked(user_id: int, course_id: int, term_id: int) -> ReportJob | None:
    for job in _JOBS.values():
        if job.running and job.user_id == user_id and job.course_id == course_id and job.term_id == term_id:
            return job
    return None


def _error_text(exc: Exception) -> str:
    from web_app.deepseek_pool import redact_text

    text = str(exc).strip() or exc.__class__.__name__
    return redact_text(text)


def _set_queued(job: ReportJob, position: int, eta: int) -> None:
    from web_app.deepseek_pool import queue_label

    with _LOCK:
        if not job.running or job.done or job.stages:
            return
        job.stage = "queued"
        job.stage_label = queue_label(position, eta)
        job.queue_position = int(position)
        job.queue_eta_seconds = int(eta)
        job.percent = 0
        job.updated_at = time.time()


def _set_stage(job: ReportJob, stage: str) -> None:
    label, percent = STAGE_INFO[stage]
    with _LOCK:
        if not job.running:
            return
        if percent < job.percent:
            percent = job.percent
        job.stage = stage
        job.stage_label = label
        job.percent = percent
        job.queue_position = None
        job.queue_eta_seconds = None
        job.updated_at = time.time()
        if not job.stages or job.stages[-1]["stage"] != stage:
            job.stages.append({"stage": stage, "percent": percent})
    hook = pause_before_stage
    if hook is not None:
        hook(stage)


def _fail(job: ReportJob, stage: str, exc: Exception) -> None:
    with _LOCK:
        job.running = False
        job.done = False
        job.failed_stage = stage or job.stage or "calculate"
        job.stage = job.failed_stage
        info = STAGE_INFO.get(job.stage)
        if info:
            job.stage_label = info[0]
        job.error = _error_text(exc)
        job.updated_at = time.time()
    _record_terminal(job, "fail")


def _succeed(job: ReportJob, summary: dict, filename: str, content: bytes) -> None:
    with _LOCK:
        job.running = False
        job.done = True
        job.error = ""
        job.failed_stage = ""
        job.summary = summary
        job.filename = filename
        job.content = content
        job.stage = "package"
        job.stage_label = STAGE_INFO["package"][0]
        job.queue_position = None
        job.queue_eta_seconds = None
        if job.percent < 100:
            job.percent = 100
        job.updated_at = time.time()
    _record_terminal(job, "success")


def _clip_error(text: str, limit: int = 500) -> str:
    cleaned = " ".join((text or "").split())
    if len(cleaned) <= limit:
        return cleaned
    return cleaned[: limit - 1] + "…"


def _record_terminal(job: ReportJob, status: str) -> None:
    """只在成功或失败时写库。写失败不能反过来弄坏已经结束的任务。"""
    try:
        from web_app.db import ReportJobEvent, session_scope

        duration_ms = int(max(0.0, job.updated_at - job.created_at) * 1000)
        with session_scope() as db:
            db.add(
                ReportJobEvent(
                    user_id=job.user_id,
                    course_id=job.course_id,
                    term_id=job.term_id,
                    status=status,
                    error=_clip_error(job.error),
                    duration_ms=duration_ms,
                )
            )
    except Exception:
        return


def _admin_status(job: ReportJob) -> str:
    if job.running and job.stage == "queued":
        return "queued"
    if job.running:
        return "running"
    if job.done:
        return "success"
    return "fail"


def _epoch_iso(timestamp: float) -> str:
    return datetime.fromtimestamp(timestamp, timezone.utc).replace(tzinfo=None).isoformat(timespec="seconds")


def list_jobs_for_admin() -> list[dict]:
    """内存任务的只读快照。不含压缩包、成绩和密钥。"""
    with _LOCK:
        _sweep_locked()
        rows = []
        for job in _JOBS.values():
            rows.append(
                {
                    "source": "memory",
                    "job_id": job.id,
                    "event_id": None,
                    "user_id": job.user_id,
                    "course_id": job.course_id,
                    "term_id": job.term_id,
                    "status": _admin_status(job),
                    "error": _clip_error(job.error, 500),
                    "duration_ms": int(max(0.0, job.updated_at - job.created_at) * 1000),
                    "created_at_epoch": job.created_at,
                }
            )
    rows.sort(key=lambda item: item["created_at_epoch"], reverse=True)
    for item in rows:
        item["created_at"] = _epoch_iso(item.pop("created_at_epoch"))
    return rows


def _worker(job: ReportJob, excel: bytes, previous: bytes | None, settings: dict, meta: dict, persist, ticket) -> None:
    current = {"stage": "calculate"}

    def on_stage(stage: str) -> None:
        current["stage"] = stage
        _set_stage(job, stage)

    lease = None
    try:
        if ticket is not None:
            lease = ticket.wait()
            lease.bind()
        bundle: ReportBundle = run_report_pipeline(excel, previous, settings, on_stage=on_stage)
        summary = dict(bundle.summary)
        summary["term_id"] = meta.get("term_id")
        summary["term_label"] = meta.get("term_label") or ""
        summary["class_name"] = meta.get("class_name") or ""
        try:
            from web_app.previous_attainment import record_last_achievement

            record_last_achievement(job.user_id, job.course_id, job.term_id, summary.get("achievement") or {})
            saved_archive = persist(bundle.files + [(bundle.filename, bundle.content)])
        except Exception as exc:
            _fail(job, "package", exc)
            return
        filename, content = saved_archive or (bundle.filename, bundle.content)
        _succeed(job, summary, filename, content)
    except Exception as exc:
        stage = current["stage"]
        with _LOCK:
            if job.stage == "queued":
                stage = "queued"
        _fail(job, stage, exc)
    finally:
        if lease is not None:
            lease.unbind()
            lease.release()
        elif ticket is not None:
            ticket.cancel()


def launch_report_job(
    *,
    user_id: int,
    course_id: int,
    term_id: int,
    excel: bytes,
    previous: bytes | None,
    settings: dict,
    meta: dict,
    persist,
    allow_ai,
) -> dict:
    with _LOCK:
        _sweep_locked()
        existing = _running_locked(user_id, course_id, term_id)
        if existing is not None:
            return existing.public()
        if allow_ai is not None and not allow_ai():
            raise ServiceError("AI 报告请求过于频繁，请稍后再试", status=429)
        job = ReportJob(user_id, course_id, term_id)
        _JOBS[job.id] = job
    ticket = _reserve_report_ticket(job)
    with _LOCK:
        if not job.running:
            return job.public()
    threading.Thread(
        target=_worker,
        args=(job, excel, previous, settings, meta, persist, ticket),
        name=f"report-{job.id[:8]}",
        daemon=True,
    ).start()
    with _LOCK:
        return job.public()


def _reserve_report_ticket(job: ReportJob):
    from web_app.deepseek_pool import PoolError, configured, reserve

    try:
        ready = configured()
    except PoolError as exc:
        _fail(job, "calculate", exc)
        return None
    if not ready:
        return None
    try:
        return reserve("report", on_queue=lambda position, eta: _set_queued(job, position, eta))
    except PoolError as exc:
        _fail(job, "calculate", exc)
        return None


def report_job_snapshot(user_id: int, job_id: str) -> dict | None:
    with _LOCK:
        _sweep_locked()
        job = _JOBS.get(job_id)
        if job is None or job.user_id != int(user_id):
            return None
        return job.public()


def report_job_download(user_id: int, job_id: str) -> tuple[str, bytes]:
    with _LOCK:
        _sweep_locked()
        job = _JOBS.get(job_id)
        if job is None or job.user_id != int(user_id):
            from web_app.db import NotFound

            raise NotFound()
        if job.error:
            raise ServiceError("报告生成失败，没有可下载的压缩包")
        if not job.done or not job.content:
            raise ServiceError("报告还在生成", status=409)
        return job.filename or "AI分析报告.zip", job.content
