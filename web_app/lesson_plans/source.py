"""保存原始 DOCX 与课程关联；教案读取时由自己的 read_outline 解析。"""
import hashlib
import io
import uuid
import zipfile
from pathlib import Path
from sqlalchemy import select
from web_app.db import Course, LessonPlanSource, NotFound, utcnow, get_engine
from sqlalchemy.orm import Session
from web_app.syllabus import compatibility
import json
from web_app.storage import upload_root, allocate_path
from web_app.resource_limits import validate_document, computation_slot
from . import engine
from .errors import LessonPlanError

MAX_BYTES = 5 * 1024 * 1024


def source_path(source):
    root=upload_root(); path=(root/source.stored_path).resolve()
    if root not in path.parents or not path.is_file(): raise LessonPlanError("原始大纲文件已不可用，请重新上传。",404)
    return path


def validate_bytes(filename,data):
    if Path(filename).suffix.lower()!=".docx": raise LessonPlanError("教学大纲只接受 .docx 文件。")
    if len(data)>MAX_BYTES: raise LessonPlanError("教学大纲不能超过 5MB。",413)
    validate_document(data,filename)
    if not zipfile.is_zipfile(io.BytesIO(data)): raise LessonPlanError("无法读取 Word 大纲，请检查文件格式。")
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        if "word/document.xml" not in archive.namelist(): raise LessonPlanError("上传内容不是 Word 大纲。")


def parse_bytes(filename,data, *, snapshot=None):
    validate_bytes(filename,data)
    try:
        with computation_slot(): outline=engine.read_outline(io.BytesIO(data))
    except engine.OutlineError:
        outline = None
    except LessonPlanError: raise
    except Exception: raise LessonPlanError("无法读取教学大纲，请检查文件是否损坏。") from None
    try:
        if snapshot is not None:
            outline = compatibility.normalize_outline(snapshot)
        else:
            compatible = compatibility.load(filename, data)["outline"]
            if outline is None:
                outline = compatibility.normalize_outline(compatible)
            else:
                outline = compatibility.enrich_outline(outline, compatible)
    except Exception:
        if outline is None:
            raise LessonPlanError("没有找到可用课次，请在“从大纲建课”中开启 AI 纠错与补全后再生成教案。") from None
    ids=[engine.pick(item,"课次") for item in outline["lessons"]]
    if len(ids)>100 or len(ids)!=len(set(ids)) or any(not v.isdigit() or int(v)<1 for v in ids):
        raise LessonPlanError("大纲课次应为不重复的正整数，最多支持 100 课次。")
    return outline


def saved_snapshot(source):
    if source.course_id is None:
        return None
    with Session(get_engine()) as db:
        course = db.get(Course, source.course_id)
        settings = json.loads(course.settings_json) if course else {}
    snapshot = settings.get("teaching_outline")
    if isinstance(snapshot, dict) and snapshot.get("source_digest") == source.digest:
        return snapshot
    return None


def read_source(source):
    path=source_path(source)
    return parse_bytes(source.filename,path.read_bytes(),snapshot=saved_snapshot(source))


def owned_source(db,user_id,source_id):
    source=db.get(LessonPlanSource,source_id)
    if source is None or source.user_id!=user_id: raise NotFound()
    return source


def store_source(db,user_id,course_id,filename,data):
    validate_bytes(filename,data)
    digest=hashlib.sha256(data).hexdigest()
    existing=db.scalar(select(LessonPlanSource).where(LessonPlanSource.user_id==user_id,LessonPlanSource.course_id==course_id,LessonPlanSource.digest==digest))
    if existing is not None:
        try:
            source_path(existing)
            # 相同文件再次作为新版大纲上传时，课程仍应指向本次选择的原文件。
            existing.created_at=utcnow()
            return existing
        except LessonPlanError: pass
    identity=uuid.uuid4().hex
    if course_id is not None: path,_=allocate_path(user_id,course_id,filename)
    else:
        folder=upload_root()/".lesson-plans"/str(user_id)/"sources"; folder.mkdir(parents=True,exist_ok=True)
        path=folder/(identity+".docx")
    path.write_bytes(data); db.info.setdefault("new_blob_paths",[]).append(path)
    row=LessonPlanSource(id=identity,user_id=user_id,course_id=course_id,filename=Path(filename.replace("\\","/")).name[:255],
                         stored_path=path.relative_to(upload_root()).as_posix(),digest=digest)
    db.add(row); db.flush(); return row


def public_source(source):
    outline=read_source(source)
    return {"id":source.id,"filename":source.filename,"course_id":source.course_id,"course_name":outline["课程名称"],
            "teacher":outline["课程负责人"],"lesson_count":len(outline["lessons"]),
            "outline_version":outline.get("completion_id","")}
