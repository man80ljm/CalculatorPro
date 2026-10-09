"""有上限的持久化报告队列。输入存磁盘，固定工作线程执行，过期租约自动接续。"""
from __future__ import annotations

import threading
import time
import uuid
import json
import os
import logging
import shutil
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from sqlalchemy import select, text

from web_app.service import ReportBundle, ServiceError, run_report_pipeline

KEEP_SECONDS = 30 * 60
STAGE_INFO = {
    "calculate": ("正在计算达成度…", 15),
    "tables": ("正在生成统计表…", 45),
    "ai": ("AI 正在撰写分析（约 30 秒）…", 75),
    "package": ("正在整理报告…", 90),
}

# 测试可以在进入某个阶段时拦住后台线程，生产环境保持 None。
pause_before_stage = None

_LOCK = threading.Lock()
_JOBS: dict[str, "ReportJob"] = {}
_RUNTIME_LOCK = threading.Lock()
_RUNTIME = None
_USERS = 0


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
        self.source_signature = ""
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
            "source_signature": self.source_signature,
            "stages": [dict(item) for item in self.stages],
            "queue_position": self.queue_position,
            "queue_eta_seconds": self.queue_eta_seconds,
        }


def reset_report_jobs() -> None:
    from web_app.deepseek_pool import reset_pool

    stop_job_runtime(force=True)
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
    completed = sorted((job for job in _JOBS.values() if not job.running), key=lambda job: job.updated_at)
    for job in completed[:-200]:
        _JOBS.pop(job.id, None)


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
    if getattr(job, "claim_token", None):
        _store_progress(job)
    hook = pause_before_stage
    if hook is not None:
        hook(stage)


def _fail(job: ReportJob, stage: str, exc: Exception, record=True) -> None:
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
    if record:
        _record_terminal(job, "fail")


def _succeed(job: ReportJob, summary: dict, filename: str, content: bytes, record=True) -> None:
    with _LOCK:
        job.running = False
        job.done = True
        job.error = ""
        job.failed_stage = ""
        job.summary = summary
        job.filename = filename
        job.content = content
        job.stage = "package"
        job.stage_label = "打包完成"
        job.queue_position = None
        job.queue_eta_seconds = None
        if job.percent < 100:
            job.percent = 100
        job.updated_at = time.time()
    if record:
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
    """数据库任务和兼容的内存调试记录。不含成绩内容和密钥。"""
    from web_app.db import ReportTask, session_scope
    with _LOCK:
        _sweep_locked()
        rows = []
        for job in _JOBS.values():
            if getattr(job, "claim_token", None):
                continue
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
    with session_scope() as db:
        tasks = list(db.scalars(select(ReportTask).order_by(ReportTask.created_at.desc()).limit(100)))
        for task in tasks:
            view = json.loads(task.public_json)
            rows.append({"source": "database", "job_id": task.id, "event_id": None, "user_id": task.user_id,
                "course_id": task.course_id, "term_id": task.term_id,
                "status": "queued" if task.status == "running" and view.get("stage") == "queued" else task.status,
                "error": _clip_error(view.get("error", "")), "duration_ms": int(max(0, (task.updated_at-task.created_at).total_seconds())*1000),
                "created_at": task.created_at.isoformat(timespec="seconds")})
    return rows


def _worker(job: ReportJob, folder: Path, settings: dict, meta: dict) -> None:
    current = {"stage": "calculate"}

    def on_stage(stage: str) -> None:
        current["stage"] = stage
        _set_stage(job, stage)

    lease = None
    ticket = None
    try:
        from web_app.deepseek_pool import configured, reserve
        if configured():
            ticket = reserve("report", on_queue=lambda position, eta: _set_queued(job, position, eta))
        if ticket is not None:
            lease = ticket.wait()
            lease.bind()
        excel = (folder / "grade.xlsx").read_bytes()
        previous_path = folder / "previous.xlsx"
        previous = previous_path.read_bytes() if previous_path.exists() else None
        bundle: ReportBundle = run_report_pipeline(excel, previous, settings, on_stage=on_stage, checkpoint_dir=folder)
        summary = dict(bundle.summary)
        summary["term_id"] = meta.get("term_id")
        summary["term_label"] = meta.get("term_label") or ""
        summary["class_name"] = meta.get("class_name") or ""
        try:
            from web_app.app import _persist_files
            saved_archive = _persist_files(job.user_id, job.course_id, bundle.files + [(bundle.filename, bundle.content)],
                "report", job.term_id, excel=excel, previous=previous,
                task_id=job.id, claim_token=job.claim_token, summary=summary)
        except Exception as exc:
            _fail_task(job, "package", exc)
            return
        filename, content = saved_archive or (bundle.filename, bundle.content)
        # 下载从课程资料读取。完成任务只保留很小的状态，释放 ZIP/Word/Excel 字节。
        _succeed(job, summary, filename, b"", record=False)
    except Exception as exc:
        stage = current["stage"]
        with _LOCK:
            if job.stage == "queued":
                stage = "queued"
        _fail_task(job, stage, exc)
    finally:
        with _LOCK:
            job.running = False
            job.updated_at = time.time()
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
    from web_app.db import ReportTask, session_scope
    from web_app.resource_limits import setting
    from web_app.storage import upload_root
    active_key = f"{int(user_id)}:{int(course_id)}:{int(term_id)}"
    folder = None
    try:
        with session_scope() as db:
            _queue_lock(db)
            existing = db.scalar(select(ReportTask).where(ReportTask.active_key == active_key))
            if existing is not None:
                return _task_public(db, existing)
            active = list(db.scalars(select(ReportTask).where(ReportTask.status.in_(["queued", "running"]))))
            if sum(item.user_id == user_id for item in active) >= 4:
                raise ServiceError("你已有 4 个报告任务，完成后即可继续提交。", status=429)
            if sum(item.status == "queued" for item in active) >= setting("REPORT_QUEUE_LIMIT", 100, maximum=500):
                raise ServiceError("报告排队人数较多，请稍后再试。", status=429)
            if allow_ai is not None and not allow_ai():
                raise ServiceError("AI 报告请求过于频繁，请稍后再试", status=429)
            job = ReportJob(user_id, course_id, term_id)
            from web_app.report_state import submitted_signature
            job.source_signature = submitted_signature(settings, excel, previous, meta.get("course_name") or "")
            _set_queued(job, max(1, len(active)), max(1, len(active)) * _estimate())
            folder = upload_root() / ".jobs" / job.id
            folder.mkdir(parents=True, exist_ok=False)
            (folder / "grade.xlsx").write_bytes(excel)
            if previous is not None:
                (folder / "previous.xlsx").write_bytes(previous)
            task = ReportTask(id=job.id, user_id=user_id, course_id=course_id, term_id=term_id,
                active_key=active_key, settings_json=json.dumps(settings, ensure_ascii=False),
                meta_json=json.dumps(meta, ensure_ascii=False), public_json=json.dumps(job.public(), ensure_ascii=False))
            db.add(task)
            db.flush()
            view = _task_public(db, task)
    except Exception:
        if folder is not None:
            shutil.rmtree(folder, ignore_errors=True)
        raise
    # 内存只记实际执行中的任务；排队记录不保留输入，也不创建等待线程。
    if os.environ.get("REPORT_WORKER_MODE", "embedded") != "external":
        _ensure_runtime()
    return view


def _estimate():
    from web_app.resource_limits import setting
    return setting("DEEPSEEK_QUEUE_REPORT_SECONDS", 30)


def _capacity():
    from web_app.deepseek_pool import load_specs
    from web_app.resource_limits import setting
    slots = setting("DEEPSEEK_GLOBAL_CONCURRENCY", 5, maximum=5)
    return max(1, min(slots, len(load_specs()) * setting("DEEPSEEK_ACCOUNT_CONCURRENCY", 5)))


def _queue_lock(db):
    from web_app.db import JobQueueLock, write_context
    if db.bind.dialect.name == "sqlite" and write_context.get() is None:
        db.execute(text("BEGIN IMMEDIATE"))
    db.scalar(select(JobQueueLock).where(JobQueueLock.id == 1).with_for_update())


def completed_public(view, summary):
    return {**view, "stage": "package", "stage_label": "打包完成", "percent": 100, "done": True,
            "summary": summary, "error": "", "failed_stage": "", "queue_position": None, "queue_eta_seconds": None}


def _task_public(db, task):
    from web_app.db import ReportTask
    view = json.loads(task.public_json)
    if task.status == "queued":
        from web_app.deepseek_pool import queue_label
        waiting = list(db.scalars(select(ReportTask).where(ReportTask.status.in_(["queued", "running"]))))
        ahead = sum(item.id != task.id and (item.status == "running" or (item.created_at, item.id) < (task.created_at, task.id)) for item in waiting)
        position = max(1, ahead)
        eta = max(1, (position + _capacity() - 1) // _capacity()) * _estimate()
        view.update(stage="queued", stage_label=queue_label(position, eta), queue_position=position, queue_eta_seconds=eta)
    return view


def _store_progress(job):
    from web_app.db import ReportTask, session_scope
    with session_scope() as db:
        _queue_lock(db)
        task = db.get(ReportTask, job.id)
        if task is None or task.status != "running" or task.claim_token != job.claim_token:
            raise ServiceError("任务已由系统重新接续，本次处理已停止。")
        task.public_json = json.dumps(job.public(), ensure_ascii=False)


def _fail_task(job, stage, exc):
    from web_app.db import ReportTask, ReportJobEvent, session_scope
    with session_scope() as db:
        _queue_lock(db)
        task = db.get(ReportTask, job.id)
        if task is None or task.status != "running" or task.claim_token != job.claim_token:
            return
        task.status = "fail"
        task.active_key = None
        task.lease_until = None
        _fail(job, stage, exc, record=False)
        task.public_json = json.dumps(job.public(), ensure_ascii=False)
        db.add(ReportJobEvent(user_id=job.user_id, course_id=job.course_id, term_id=job.term_id,
            status="fail", error=_clip_error(job.error), duration_ms=int(max(0,job.updated_at-job.created_at)*1000)))


def report_job_snapshot(user_id: int, job_id: str) -> dict | None:
    from web_app.db import ReportTask, session_scope
    with session_scope() as db:
        task = db.get(ReportTask, str(job_id))
        if task is not None and task.user_id == int(user_id):
            return _task_public(db, task)
    return None


def current_report_job(user_id, course_id, term_id):
    from web_app.db import ReportTask, session_scope
    with session_scope() as db:
        task = db.scalar(select(ReportTask).where(ReportTask.user_id == user_id, ReportTask.course_id == course_id,
            ReportTask.term_id == term_id).order_by(ReportTask.created_at.desc()).limit(1))
        return _task_public(db, task) if task else None


def report_job_download(user_id: int, job_id: str):
    from web_app.db import CourseFile, NotFound, ReportTask, session_scope
    from web_app.storage import resolve_stored
    with session_scope() as db:
        task = db.get(ReportTask, str(job_id))
        if task is None or task.user_id != int(user_id):
            raise NotFound()
        if task.status == "fail":
            raise ServiceError("报告生成失败，没有可下载的压缩包")
        if task.status != "success":
            raise ServiceError("报告还在生成", status=409)
        archive = db.get(CourseFile, task.archive_file_id) if task.archive_file_id else None
        if archive is None or archive.user_id != user_id:
            raise NotFound()
        path = resolve_stored(archive.user_id, archive.course_id, archive.stored_name)
        if not path.is_file():
            raise NotFound()
        return archive.original_name, path


class JobRuntime:
    def __init__(self):
        from web_app.db import get_engine
        from web_app.storage import upload_root
        self.engine = get_engine()
        self.root = upload_root()
        self.stop_event = threading.Event()
        self.pool = ThreadPoolExecutor(max_workers=5, thread_name_prefix="report-worker")
        self.active = {}
        self.last_cleanup = 0
        self.thread = threading.Thread(target=self.loop, name="report-dispatcher", daemon=True)
        self.thread.start()

    def context(self):
        from contextlib import contextmanager
        from web_app.db import job_engine, write_context
        from web_app.storage import job_upload_root
        @contextmanager
        def scope():
            tokens = [job_engine.set(self.engine), job_upload_root.set(self.root), write_context.set(None)]
            try:
                yield
            finally:
                write_context.reset(tokens[2]); job_upload_root.reset(tokens[1]); job_engine.reset(tokens[0])
        return scope()

    def run(self, task):
        with self.context():
            job = ReportJob(task.user_id, task.course_id, task.term_id)
            job.id = task.id
            job.claim_token = task.claim_token
            job.created_at = task.created_at.replace(tzinfo=timezone.utc).timestamp()
            # 恢复后进度不倒退；同一阶段仍可安全重新计算，AI 已完成的答案会复用。
            old = json.loads(task.public_json)
            job.source_signature = old.get("source_signature") or ""
            job.percent = int(old.get("percent") or 0)
            job.stages = old.get("stages") or []
            with _LOCK:
                _JOBS[job.id] = job
            _worker(job, self.root / ".jobs" / task.id, json.loads(task.settings_json), json.loads(task.meta_json))

    def tick(self):
        from web_app.db import ReportTask, session_scope, utcnow
        from web_app.resource_limits import setting
        now = utcnow()
        lease = now + timedelta(seconds=setting("REPORT_LEASE_SECONDS", 90, minimum=5))
        with session_scope() as db:
            _queue_lock(db)
            rows = list(db.scalars(select(ReportTask).where(ReportTask.status.in_(["queued", "running"])).order_by(ReportTask.created_at, ReportTask.id).with_for_update()))
            for task in rows:
                if task.status == "running":
                    if self.active.get(task.id, (None, None))[1] == task.claim_token:
                        task.lease_until = lease
                    elif task.lease_until is None or task.lease_until < now:
                        task.status = "queued"
                        task.claim_token = ""
                        task.lease_until = None
            running = [task for task in rows if task.status == "running"]
            users = {task.user_id for task in running}
            chosen = []
            available = min(5 - len(self.active), _capacity() - len(running))
            for task in rows:
                if available <= 0:
                    break
                if task.status != "queued" or task.user_id in users:
                    continue
                task.status = "running"
                task.claim_token = uuid.uuid4().hex
                task.lease_until = lease
                users.add(task.user_id)
                available -= 1
                chosen.append(task)
        for task in chosen:
            self.active[task.id] = (self.pool.submit(self.run, task), task.claim_token)
        finished = [key for key, (future, _) in self.active.items() if future.done()]
        for key in finished:
            future, _ = self.active.pop(key)
            try:
                future.result()
            except Exception:
                logging.getLogger(__name__).exception("报告工作线程异常，租约到期后会自动恢复")
        with _LOCK:
            _sweep_locked()
        if time.monotonic() - self.last_cleanup > 3600:
            self.last_cleanup = time.monotonic()
            self.cleanup_inputs()

    def cleanup_inputs(self):
        from web_app.db import ReportTask, session_scope, utcnow
        jobs_root = self.root / ".jobs"
        if not jobs_root.is_dir():
            return
        with session_scope() as db:
            tasks = {row.id: row for row in db.scalars(select(ReportTask))}
            for folder in jobs_root.iterdir():
                if folder.is_symlink() or not folder.is_dir() or len(folder.name) != 32:
                    continue
                task = tasks.get(folder.name)
                stale = task is not None and task.status in {"success", "fail"} and task.updated_at < utcnow() - timedelta(days=1)
                orphan = task is None and time.time() - folder.stat().st_mtime > 86400
                if stale or orphan:
                    shutil.rmtree(folder)
                    if task is not None:
                        task.settings_json = "{}"
                        task.meta_json = "{}"

    def loop(self):
        with self.context():
            while not self.stop_event.is_set():
                try:
                    self.tick()
                except Exception:
                    logging.getLogger(__name__).exception("报告调度暂时失败，将自动重试")
                self.stop_event.wait(0.3)

    def stop(self):
        self.stop_event.set()
        self.thread.join(timeout=5)
        self.pool.shutdown(wait=True)


def _ensure_runtime():
    global _RUNTIME
    with _RUNTIME_LOCK:
        if _RUNTIME is None:
            _RUNTIME = JobRuntime()


def start_job_runtime():
    global _USERS
    if os.environ.get("REPORT_WORKER_MODE", "embedded") == "external":
        return
    _ensure_runtime()
    with _RUNTIME_LOCK:
        _USERS += 1


def stop_job_runtime(force=False):
    global _RUNTIME, _USERS
    runtime = None
    with _RUNTIME_LOCK:
        _USERS = 0 if force else max(0, _USERS - 1)
        if _USERS == 0:
            runtime, _RUNTIME = _RUNTIME, None
    if runtime is not None:
        runtime.stop()


def worker_main():
    from web_app.db import init_db
    from web_app.storage import ensure_upload_root
    init_db(); ensure_upload_root()
    runtime = JobRuntime()
    try:
        while runtime.thread.is_alive():
            runtime.thread.join(timeout=1)
    except KeyboardInterrupt:
        runtime.stop()


if __name__ == "__main__":
    worker_main()
