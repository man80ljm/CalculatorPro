"""教案专用持久化队列，复用报告队列锁/租约做法，不修改报告任务。"""
import copy
import json
import os
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import timedelta, datetime
from pathlib import Path
from sqlalchemy import select
from sqlalchemy.orm import Session
from web_app.db import LessonPlan, LessonPlanTask, LessonPlanSource, get_engine, utcnow
from web_app.report_jobs import _queue_lock
from web_app.storage import upload_root
from web_app.teaching_calendar.academic_term import CHINA_TIMEZONE
from . import engine, service, client
from .source import read_source, source_path
from .errors import LessonPlanError

_runtimes={}
_refs={}
_guard=threading.RLock()


def task_view(task):
    view=json.loads(task.public_json)
    return {**view,"job_id":task.id,"plan_id":task.plan_id,"action":task.action,"status":task.status,
            "done":task.status=="success","error":view.get("error","")}


def launch(db,user_id,plan,action,lesson=None):
    if action not in {"generate","compress","check","export"}: raise LessonPlanError("未知教案任务。")
    _queue_lock(db)
    active=db.scalar(select(LessonPlanTask).where(LessonPlanTask.active_key==plan.id))
    if active:
        if active.action!=action or active.lesson!=lesson: raise LessonPlanError("此教案正在处理其他任务，请完成后再试。",409)
        return task_view(active)
    tasks=list(db.scalars(select(LessonPlanTask).where(LessonPlanTask.status.in_(["queued","running"]))))
    if sum(t.user_id==user_id for t in tasks)>=2 or len(tasks)>=40: raise LessonPlanError("教案任务正在排队，请稍后再试。",429)
    if action in {"generate","compress"} and not client.api_key(): raise LessonPlanError("尚未配置服务器 DeepSeek API。",503,code="ai_not_configured")
    context=json.loads(plan.context_json)
    if action=="compress" and lesson not in {row["课次"] for row in context["lessons"]}: raise LessonPlanError("没有找到该课次。",404)
    row=LessonPlanTask(id=uuid.uuid4().hex,user_id=user_id,plan_id=plan.id,action=action,lesson=lesson,active_key=plan.id,
        input_revision=plan.revision,public_json=json.dumps({"percent":0,"stage_label":"等待后台生成…","completed":0,"total":len(context["lessons"]),"api_calls":0},ensure_ascii=False))
    db.add(row); db.flush(); return task_view(row)


class Runtime:
    def __init__(self,engine,root):
        self.engine=engine; self.root=root; self.stopping=threading.Event(); self.active={}
        self.pool=ThreadPoolExecutor(max_workers=2,thread_name_prefix="lesson-plan-worker")
        self.thread=threading.Thread(target=self.loop,name="lesson-plan-dispatcher",daemon=True); self.thread.start()

    def update(self,identity,token,*,public=None,checkpoint=None,context=None):
        with Session(self.engine) as db:
            _queue_lock(db); task=db.get(LessonPlanTask,identity)
            if task is None or task.claim_token!=token or task.status!="running": raise LessonPlanError("任务已重新接续，本次处理停止。",409)
            if public is not None:
                previous=json.loads(task.public_json); previous.update(public); task.public_json=json.dumps(previous,ensure_ascii=False)
            if checkpoint is not None: task.checkpoint_json=json.dumps(checkpoint,ensure_ascii=False)
            if context is not None:
                plan=db.get(LessonPlan,task.plan_id)
                if plan is None or plan.revision!=task.input_revision: raise LessonPlanError("教案版本已更新，请重新生成。",409)
                plan.context_json=json.dumps(context,ensure_ascii=False)
            task.updated_at=utcnow(); db.commit()

    def run(self,identity,token):
        output=None; committed=False
        try:
            with Session(self.engine) as db:
                task=db.get(LessonPlanTask,identity); plan=db.get(LessonPlan,task.plan_id); source=db.get(LessonPlanSource,plan.source_id)
                context=json.loads(plan.context_json); checkpoint=json.loads(task.checkpoint_json)
                snapshot={"action":task.action,"lesson":task.lesson,"plan_id":plan.id,"user_id":plan.user_id,"revision":task.input_revision}
                path=(self.root/source.stored_path).resolve()
                if self.root not in path.parents or not path.is_file(): raise LessonPlanError("原始大纲已不可用，请重新上传。")
                filename=source.filename
            from .source import parse_bytes
            outline=parse_bytes(filename,path.read_bytes(),snapshot=context.get("_outline"))
            notices=checkpoint.get("warnings",[]); done=checkpoint.setdefault("lessons",{}); calls=checkpoint.get("api_calls",0)
            mutex=threading.RLock()
            def progress(label,percent,**extra): self.update(identity,token,public={"stage_label":label,"percent":percent,**extra})
            def call_notice(number=None):
                nonlocal calls
                with mutex:
                    calls+=1; checkpoint["api_calls"]=calls
                    self.update(identity,token,public={"api_calls":calls},checkpoint=checkpoint)
            items={int(engine.pick(item,"课次")):item for item in outline["lessons"]}
            if snapshot["action"] in {"generate","compress"}:
                targets=[row for row in context["lessons"] if snapshot["action"]=="generate" or row["课次"]==snapshot["lesson"]]
                def one(row):
                    number=row["课次"]; attempts=0
                    def observe():
                        nonlocal attempts
                        attempts+=1; call_notice(number)
                        with mutex:
                            completed=len(done)
                            progress(f"正在生成第 {number} 课次（已完成 {completed}/{len(targets)}）" if attempts==1 else f"正在调整第 {number} 课次字数（仅重写一次）",
                                     10+int(65*completed/len(targets)),completed=completed,total=len(targets))
                    extra=""
                    if snapshot["action"]=="compress":
                        total=sum(engine.hanzi_len(row[k]) for k in engine.LIMITS)
                        extra=f"上一版共 {total} 字，排版后超过一页。请把四栏总字数压缩到约 {int(total*0.8)} 字，保留结构。\n当前教案："+json.dumps({k:row[k] for k in engine.LIMITS},ensure_ascii=False)
                    try:
                        with client.observe_calls(observe): fields=engine.gen_lesson(outline,items[number],client.api_key(),extra=extra)
                        candidate=copy.deepcopy(context); next(r for r in candidate["lessons"] if r["课次"]==number).update(fields)
                        service.normalize_context(candidate)
                        return number,fields,None
                    except Exception as exc:
                        reason=str(exc) if isinstance(exc,LessonPlanError) else "AI 返回内容无法使用"
                        if snapshot["action"]=="compress":
                            return number,{k:row[k] for k in engine.LIMITS},f"第 {number} 课次压缩失败：{reason}；已保留当前编辑内容。"
                        return number,engine.rule_lesson(items[number]),f"第 {number} 课次生成失败：{reason}；已使用大纲原文，请编辑后核对。"
                remaining=[row for row in targets if str(row["课次"]) not in done]
                with ThreadPoolExecutor(max_workers=4,thread_name_prefix="lesson") as pool:
                    futures=[pool.submit(one,copy.deepcopy(row)) for row in remaining]
                    for future in as_completed(futures):
                        number,fields,warning=future.result()
                        with mutex:
                            done[str(number)]=fields
                            if warning: notices.append(warning)
                            checkpoint["warnings"]=notices
                            for row in context["lessons"]:
                                if str(row["课次"]) in done: row.update(done[str(row["课次"])])
                            self.update(identity,token,checkpoint=checkpoint,context=context,
                                public={"completed":len(done),"total":len(targets),"percent":10+int(65*len(done)/len(targets)),"stage_label":f"已完成 {len(done)}/{len(targets)} 课次"})
                for row in context["lessons"]:
                    if str(row["课次"]) in done: row.update(done[str(row["课次"])])
                if snapshot["action"]=="generate":
                    progress("正在生成课程评价方式…",80)
                    if "evaluation" not in checkpoint:
                        with client.observe_calls(call_notice): checkpoint["evaluation"]=engine.build_eval(outline,client.api_key(),notices=notices)
                        self.update(identity,token,checkpoint=checkpoint)
                    context["评价"]=checkpoint["evaluation"]
            progress("正在渲染 Word 并检查各课次页数…",90)
            rendered_template = service.template_digest()
            content,pages,page_notices=service.make_document(context,check=True)
            checkpoint["template_digest"] = rendered_template
            notices.extend(page_notices)
            with Session(self.engine) as db:
                _queue_lock(db); task=db.get(LessonPlanTask,identity); plan=db.get(LessonPlan,snapshot["plan_id"])
                if task is None or task.claim_token!=token or task.status!="running" or plan is None or plan.revision!=snapshot["revision"]:
                    return
                if snapshot["action"] in {"generate","compress"}:
                    plan.context_json=json.dumps(service.normalize_context(context),ensure_ascii=False); plan.revision+=1
                task.checkpoint_json=json.dumps(checkpoint,ensure_ascii=False)
                plan.pages_json=json.dumps(pages); plan.checked_revision=plan.revision
                plan.warnings_json=json.dumps(list(dict.fromkeys(notices)),ensure_ascii=False)
                if snapshot["action"]=="export":
                    folder=path.parent if plan.course_id is not None else self.root/".lesson-plans"/str(plan.user_id)/plan.id
                    folder.mkdir(parents=True,exist_ok=True)
                    output=folder/(identity+".docx"); output.write_bytes(content); db.info.setdefault("new_blob_paths",[]).append(output)
                    previous=plan.output_path
                    plan.output_path=output.relative_to(self.root).as_posix(); plan.output_revision=plan.revision
                    plan.output_date=datetime.now(CHINA_TIMEZONE).date().isoformat()
                task.status="success"; task.active_key=None; task.lease_until=None
                label="Word 已生成，可下载" if snapshot["action"]=="export" else "教案已生成，请检查和编辑"
                if snapshot["action"]=="generate":
                    failed=sum(any(note.startswith(f"第 {row['课次']} 课次生成失败") for note in notices) for row in context["lessons"])
                    if failed==len(context["lessons"]): label="AI 生成未完成，已保留大纲原文，请查看提示"
                    elif failed: label=f"教案已生成，{failed} 个课次采用大纲原文，请核对"
                elif snapshot["action"]=="compress" and any("压缩失败" in note for note in notices):
                    label="压缩未完成，已保留当前编辑内容"
                view=json.loads(task.public_json); view.update(percent=100,stage_label=label,error="")
                task.public_json=json.dumps(view,ensure_ascii=False); db.commit(); committed=True
                if snapshot["action"]=="export" and previous:
                    old=(self.root/previous).resolve()
                    if self.root in old.parents and old!=output:
                        try: old.unlink(missing_ok=True)
                        except OSError: pass
        except Exception as exc:
            if output is not None and not committed: output.unlink(missing_ok=True)
            message=str(exc) if isinstance(exc,LessonPlanError) else "教案任务失败，请稍后重试。"
            with Session(self.engine) as db:
                _queue_lock(db); task=db.get(LessonPlanTask,identity)
                if task is not None and task.claim_token==token:
                    view=json.loads(task.public_json); view.update(error=message,stage_label="生成未完成")
                    task.public_json=json.dumps(view,ensure_ascii=False); task.status="fail"; task.active_key=None; task.lease_until=None; db.commit()

    def tick(self):
        now=utcnow()
        with Session(self.engine) as db:
            _queue_lock(db)
            tasks=list(db.scalars(select(LessonPlanTask).where(LessonPlanTask.status.in_(["queued","running"])).order_by(LessonPlanTask.created_at,LessonPlanTask.id).with_for_update()))
            for task in tasks:
                if task.status=="running":
                    if self.active.get(task.id,(None,None))[1]==task.claim_token: task.lease_until=now+timedelta(seconds=90)
                    elif task.lease_until is None or task.lease_until<now: task.status="queued"; task.claim_token=""; task.lease_until=None
            running=[task for task in tasks if task.status=="running"]; users={task.user_id for task in running}; selected=[]
            available=min(2-len(self.active),2-len(running))
            for task in tasks:
                if available<=0: break
                if task.status!="queued" or task.user_id in users: continue
                task.status="running"; task.claim_token=uuid.uuid4().hex; task.lease_until=now+timedelta(seconds=90)
                selected.append((task.id,task.claim_token)); users.add(task.user_id); available-=1
            db.commit()
        for identity,token in selected: self.active[identity]=(self.pool.submit(self.run,identity,token),token)
        for identity,(future,token) in list(self.active.items()):
            if future.done():
                self.active.pop(identity); future.result()

    def loop(self):
        while not self.stopping.is_set():
            try: self.tick()
            except Exception: pass  # 下一轮调度重试；业务错误在任务记录中返回。
            self.stopping.wait(0.5)

    def stop(self):
        self.stopping.set(); self.thread.join(timeout=5); self.pool.shutdown(wait=True)


def start_runtime(*,managed=False):
    engine=get_engine(); identity=id(engine)
    with _guard:
        if managed: _refs[identity]=_refs.get(identity,0)+1
        runtime=_runtimes.get(identity)
        if runtime is None:
            if managed:
                with Session(engine) as db:
                    pending=db.scalar(select(LessonPlanTask.id).where(LessonPlanTask.status.in_(["queued","running"])).limit(1))
                if pending is None: return None
            runtime=Runtime(engine,upload_root()); _runtimes[identity]=runtime
        return runtime


def stop_runtime(*,managed=False):
    identity=id(get_engine())
    with _guard:
        if managed:
            count=max(0,_refs.get(identity,0)-1)
            if count: _refs[identity]=count; return
            _refs.pop(identity,None)
        runtime=_runtimes.pop(identity,None)
    if runtime: runtime.stop()
