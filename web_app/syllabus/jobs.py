"""Independent persistent syllabus queue, using the existing queue lock/lease convention."""
import hashlib
import json
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import timedelta
from pathlib import Path

from sqlalchemy import event, select
from sqlalchemy.orm import Session
from web_app.db import Course, SyllabusImportTask, LessonPlanSource, LessonPlan, get_engine, job_engine, utcnow
from web_app.report_jobs import _queue_lock
from web_app.storage import upload_root
from web_app.service import ServiceError
from web_app.lesson_plans import source
from . import compatibility, completion, service

_runtimes = {}
_refs = {}
_guard = threading.RLock()


def view(task):
    public = json.loads(task.public_json)
    result = {**public, "job_id": task.id, "status": task.status,
              "done": task.status in {"success","fail"}, "error": public.get("error","")}
    if task.status == "success":
        result["draft"] = json.loads(task.result_json)["draft"]
    return result


def launch(db, user_id, filename, data, allow_completion=True, course_id=None):
    if len(data) > compatibility.MAX_BYTES:
        raise ServiceError("教学大纲不能超过 5MB。", status=413)
    if Path(filename).suffix.lower() not in {".docx",".pdf"}:
        raise ServiceError("请上传 .docx 大纲或带文字的 PDF。")
    _queue_lock(db)
    digest = hashlib.sha256(data).hexdigest()
    active = db.scalar(select(SyllabusImportTask).where(SyllabusImportTask.active_key == str(user_id)))
    if active:
        if active.digest == digest and bool(active.allow_completion) == bool(allow_completion):
            return view(active)
        raise ServiceError("已有大纲正在读取，请完成后再上传。", status=409)
    count = len(list(db.scalars(select(SyllabusImportTask.id).where(SyllabusImportTask.status.in_(["queued","running"])))))
    if count >= 40:
        raise ServiceError("大纲读取任务正在排队，请稍后再试。", status=429)
    identity = uuid.uuid4().hex
    source_id = None
    if Path(filename).suffix.lower() == ".docx":
        row = source.store_source(db, user_id, None, filename, data)
        source_id, path = row.id, source.source_path(row)
    else:
        folder = upload_root()/".syllabus-jobs"/str(user_id)
        folder.mkdir(parents=True, exist_ok=True)
        path = folder/(identity+".pdf")
        path.write_bytes(data)
        db.info.setdefault("new_blob_paths",[]).append(path)
    existing_course = db.get(Course,course_id) if course_id else None
    expected_version = json.loads(existing_course.settings_json).get("syllabus_version",1) if existing_course else None
    row = SyllabusImportTask(id=identity,user_id=user_id,course_id=course_id,source_id=source_id,expected_version=expected_version,
        filename=Path(filename.replace("\\","/")).name[:255],digest=digest,input_path=path.relative_to(upload_root()).as_posix(),
        allow_completion=int(allow_completion),active_key=str(user_id),
        public_json=json.dumps({"percent":0,"stage_label":"等待后台读取…","api_calls":0},ensure_ascii=False))
    db.add(row)
    db.flush()
    return view(row)


def bind_to_course(db, user_id, course, body, settings):
    identity = body.get("completion_job_id")
    if not identity:
        return
    row = db.get(SyllabusImportTask, str(identity))
    if row is None or row.user_id != user_id:
        raise ServiceError("大纲补全记录不存在。", status=404)
    if row.status != "success" or (row.course_id is not None and row.course_id != course.id):
        raise ServiceError("此大纲任务尚未完成或不属于当前课程。", status=409)
    result = json.loads(row.result_json)
    outline = result["outline"]
    if not outline.get("lessons"):
        return  # Text PDF can create a course without a downloadable teaching outline.
    outline = compatibility.normalize_outline(outline)
    outline["课程名称"] = settings["course_basic_info"]["course_name"]
    outline["课程简介"] = settings.get("course_description","")
    outline["课程目标"] = settings.get("objective_requirements",[])
    outline["source_digest"] = row.digest
    outline["completion_id"] = row.id
    final_draft = {"fields": {"course_name": outline["课程名称"], "hours": settings["course_basic_info"].get("hours","")},
        "objectives":outline["课程目标"],"grad_req_map":settings["grad_req_map"],
        "relation":compatibility.validate_relation_grid(settings["relation_grid"])}
    errors = completion.draft_errors(final_draft, outline)
    if errors:
        raise ServiceError("；".join(errors))
    # Evaluation percentages follow the user's final confirmed calculation table.
    relation = final_draft["relation"]
    outline["考核"]["平时比例"] = str(relation["ratios"]["usual"]*100)
    outline["考核"]["期末比例"] = str(relation["ratios"]["final"]*100)
    outline["考核"]["平时项"] = [method["name"] for link in relation["payload"]["links"] if "平时" in link["name"] for method in link["methods"]]
    outline["考核"]["期末项"] = [method["name"] for link in relation["payload"]["links"] if "期末" in link["name"] for method in link["methods"]]
    outline["考核"]["原表"] = "\n".join(" | ".join(row) for row in settings["relation_grid"])
    settings["teaching_outline"] = outline
    settings["syllabus_completion"] = result["draft"].get("completion",{})
    original = db.get(LessonPlanSource, row.source_id) if row.source_id else None
    if original:
        staging = original
        if original.course_id != course.id:
            original = source.store_source(db, user_id, course.id, original.filename, source.source_path(original).read_bytes())
        row.source_id = original.id
        original.created_at = utcnow()
        if staging.course_id is None and staging.id != original.id:
            other_job = db.scalar(select(SyllabusImportTask.id).where(SyllabusImportTask.source_id == staging.id,SyllabusImportTask.id != row.id).limit(1))
            used = db.scalar(select(LessonPlan.id).where(LessonPlan.source_id == staging.id).limit(1))
            if not other_job and not used:
                old_path = source.source_path(staging)
                db.delete(staging)
                def remove_staging(_session):
                    try: old_path.unlink(missing_ok=True)
                    except OSError: pass
                event.listen(db,"after_commit",remove_staging,once=True)
    row.course_id = course.id
    db.flush()


def preserve_confirmed_outline(course, settings, term=None):
    # Ordinary saves (including older browser tabs) cannot erase validated teaching data.
    from web_app.course_versions import term_curriculum
    previous = term_curriculum(json.loads(course.settings_json),term)
    for key in ("teaching_outline","syllabus_completion"):
        if key in previous:
            settings[key] = previous[key]
        else:
            settings.pop(key,None)


class Runtime:
    def __init__(self, engine, root):
        self.engine, self.root = engine, root
        self.stopping = threading.Event()
        self.active = {}
        self.pool = ThreadPoolExecutor(max_workers=2,thread_name_prefix="syllabus-worker")
        self.thread = threading.Thread(target=self.loop,name="syllabus-dispatcher",daemon=True)
        self.thread.start()

    def update(self, identity, token, **public):
        with Session(self.engine) as db:
            _queue_lock(db)
            row = db.get(SyllabusImportTask, identity)
            if row is None or row.claim_token != token or row.status != "running":
                raise ServiceError("大纲任务已重新接续。", status=409)
            previous = json.loads(row.public_json)
            previous.update(public)
            row.public_json = json.dumps(previous,ensure_ascii=False)
            row.updated_at = utcnow()
            db.commit()

    def run(self, identity, token):
        fixed = job_engine.set(self.engine)
        cleanup = None
        try:
            with Session(self.engine) as db:
                row = db.get(SyllabusImportTask,identity)
                filename, allow_completion = row.filename, bool(row.allow_completion)
                call_offset = json.loads(row.public_json).get("api_calls",0)
                path = (self.root/row.input_path).resolve()
                if self.root not in path.parents or not path.is_file():
                    raise ServiceError("原始大纲不可用，请重新上传。")
                if path.suffix == ".pdf":
                    cleanup = path
            self.update(identity,token,percent=10,stage_label="正在读取原文并识别旧版表格…")
            data = path.read_bytes()
            analysis = compatibility.load(filename,data)
            self.update(identity,token,percent=30,stage_label="已识别课程目标、课次及原文成绩占比，正在检查缺项…")
            if allow_completion:
                draft, outline = completion.run(analysis,call_offset=call_offset,progress=lambda **values:self.update(identity,token,**values))
            else:
                draft = service.extract_draft(filename,data)
                # Deterministically recover legacy matrices; no invented pedagogical content.
                if any(not item.get("requirement",{}).get("value") for item in draft["grad_req_map"]):
                    draft["grad_req_map"] = analysis["draft"]["grad_req_map"]
                if len(draft["objectives"]) < len(analysis["draft"]["objectives"]):
                    draft["objectives"] = analysis["draft"]["objectives"]
                outline = analysis["outline"]
                errors = completion.draft_errors(draft,outline)
                draft["completion"] = {"ok":not errors,"errors":errors,"changes":[],"warnings":analysis["issues"],"inferred_fields":[]}
            draft["completion_job_id"] = identity
            with Session(self.engine) as db:
                saved = db.get(SyllabusImportTask,identity)
                draft["source_course_id"] = saved.course_id
                draft["expected_version"] = saved.expected_version
            draft.setdefault("meta",{})["source_digest"] = analysis["source_digest"]
            with Session(self.engine) as db:
                _queue_lock(db)
                row = db.get(SyllabusImportTask,identity)
                if row is None or row.claim_token != token or row.status != "running":
                    cleanup = None
                    return
                row.result_json = json.dumps({"draft":draft,"outline":outline},ensure_ascii=False)
                row.public_json = json.dumps({"percent":100,"stage_label":"读取完成，请检查补全内容并确认建课",
                                             "api_calls":draft.get("meta",{}).get("calls",0)},ensure_ascii=False)
                row.status = "success"
                row.active_key = None
                row.lease_until = None
                row.input_path = ""
                db.commit()
        except Exception as exc:
            message = str(exc) if isinstance(exc,ServiceError) else "读取大纲失败，请检查文件或稍后重试。"
            with Session(self.engine) as db:
                _queue_lock(db)
                row = db.get(SyllabusImportTask,identity)
                if row is not None and row.claim_token == token:
                    public = json.loads(row.public_json)
                    public.update(error=message,stage_label="读取未完成")
                    row.public_json = json.dumps(public,ensure_ascii=False)
                    row.status, row.active_key, row.lease_until = "fail", None, None
                    row.input_path = ""
                    db.commit()
                else:
                    cleanup = None
        finally:
            if cleanup:
                cleanup.unlink(missing_ok=True)
            job_engine.reset(fixed)

    def tick(self):
        now = utcnow()
        with Session(self.engine) as db:
            _queue_lock(db)
            tasks = list(db.scalars(select(SyllabusImportTask).where(SyllabusImportTask.status.in_(["queued","running"]))
                .order_by(SyllabusImportTask.created_at).with_for_update()))
            for row in tasks:
                if row.status == "running":
                    if self.active.get(row.id,(None,None))[1] == row.claim_token:
                        row.lease_until = now + timedelta(seconds=90)
                    elif row.lease_until is None or row.lease_until < now:
                        row.status, row.claim_token, row.lease_until = "queued", "", None
            available = min(2-len(self.active),2-sum(row.status=="running" for row in tasks))
            selected = []
            for row in tasks:
                if available <= 0:
                    break
                if row.status == "queued":
                    row.status, row.claim_token = "running", uuid.uuid4().hex
                    row.lease_until = now + timedelta(seconds=90)
                    selected.append((row.id,row.claim_token))
                    available -= 1
            db.commit()
        for identity, token in selected:
            self.active[identity] = (self.pool.submit(self.run,identity,token),token)
        for identity, (future, _) in list(self.active.items()):
            if future.done():
                self.active.pop(identity)
                future.result()

    def loop(self):
        while not self.stopping.is_set():
            try:
                self.tick()
            except Exception:
                pass
            self.stopping.wait(0.5)

    def stop(self):
        self.stopping.set()
        self.thread.join(timeout=5)
        self.pool.shutdown(wait=True)


def start_runtime(*, managed=False):
    engine = get_engine()
    identity = id(engine)
    with _guard:
        if managed:
            _refs[identity] = _refs.get(identity,0)+1
        if identity not in _runtimes:
            if managed:
                with Session(engine) as db:
                    pending = db.scalar(select(SyllabusImportTask.id).where(SyllabusImportTask.status.in_(["queued","running"])).limit(1))
                if pending is None:
                    return None
            _runtimes[identity] = Runtime(engine,upload_root())
        return _runtimes[identity]


def stop_runtime(*, managed=False):
    identity = id(get_engine())
    with _guard:
        if managed:
            count = max(0,_refs.get(identity,0)-1)
            if count:
                _refs[identity] = count
                return
            _refs.pop(identity,None)
        runtime = _runtimes.pop(identity,None)
    if runtime:
        runtime.stop()
