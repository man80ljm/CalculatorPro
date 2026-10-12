"""Production SQLAlchemy store; imported records remain administrator-only."""
import json,uuid
from contextlib import contextmanager
from sqlalchemy import select,update,text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session
from web_app.db import ScheduleDocument,ScheduleDraft,ScheduleRecord,ScheduleRecognitionChunk,utcnow
from .api_settings import ApiError

def dump(value):return json.dumps(value,ensure_ascii=False,separators=(',',':'))
def iso(value):return value.isoformat(timespec='seconds')+'Z'

class ScheduleDatabase:
    def __init__(self,engine):self.engine=engine
    @contextmanager
    def connect(self,write=False):
        with Session(self.engine,expire_on_commit=False) as db:
            try:
                if write and self.engine.dialect.name=='sqlite':db.execute(text('BEGIN IMMEDIATE'))
                yield db
                db.commit()
            except Exception:
                db.rollback();raise
    def read_chunk(self,key):
        with self.connect() as db:r=db.get(ScheduleRecognitionChunk,key)
        return {'payload':json.loads(r.payload),'usage':json.loads(r.usage)} if r else None
    def save_chunk(self,key,payload,usage):
        with self.connect(True) as db:
            row=db.get(ScheduleRecognitionChunk,key)
            if row:
                row.payload=dump(payload);row.usage=dump(usage)
            else:
                try:
                    with db.begin_nested():
                        db.add(ScheduleRecognitionChunk(cache_key=key,payload=dump(payload),usage=dump(usage)));db.flush()
                except IntegrityError:
                    # A concurrent identical batch already cached the validated result.
                    if db.get(ScheduleRecognitionChunk,key) is None:raise
    def recover_interrupted(self):
        with self.connect(True) as db:
            db.execute(update(ScheduleDraft).where(ScheduleDraft.status.in_(['extracting','recognizing'])).values(status='error',error='服务重启中断了识别，请重新选择同一份 PDF；已完成部分会自动复用。'))
    def find_source(self,digest):
        with self.connect() as db:
            document=db.scalar(select(ScheduleDocument.id).where(ScheduleDocument.digest==digest))
            if document:return {'duplicate':True,'documentId':document}
            draft=db.scalar(select(ScheduleDraft).where(ScheduleDraft.digest==digest,ScheduleDraft.status.in_(['extracting','recognizing','ready'])).order_by(ScheduleDraft.created_at.desc()).limit(1))
            if draft:return {'jobId':draft.id,'reused':True}
    def create_draft(self,identifier,filename,digest,model):
        with self.connect(True) as db:db.add(ScheduleDraft(id=identifier,filename=filename,digest=digest,model=model,status='extracting'))
    def update_draft(self,identifier,status,result=None,error=None):
        with self.connect(True) as db:
            row=db.scalar(select(ScheduleDraft).where(ScheduleDraft.id==identifier).with_for_update())
            if not row:raise ApiError('识别任务不存在。',404)
            row.status=status;row.error=error
            if result is not None:row.result=dump(result)
    def draft(self,identifier):
        with self.connect() as db:row=db.get(ScheduleDraft,identifier)
        if not row:raise ApiError('识别任务不存在。',404)
        return {'id':row.id,'filename':row.filename,'model':row.model,'status':row.status,'result':json.loads(row.result) if row.result else None,'message':row.error}
    def save_draft(self,identifier,result):
        try:
            with self.connect(True) as db:
                source=db.scalar(select(ScheduleDraft).where(ScheduleDraft.id==identifier).with_for_update())
                if not source:raise ApiError('识别任务不存在。',404)
                if source.status=='saved':return
                if source.status!='ready':raise ApiError('此任务还不能保存。')
                if db.scalar(select(ScheduleDocument.id).where(ScheduleDocument.digest==source.digest)):
                    raise ApiError('此 PDF 已保存，不会重复入库。')
                document_id=uuid.uuid4().hex
                db.add(ScheduleDocument(id=document_id,filename=source.filename,digest=source.digest,term=dump(result['term']),model=source.model));db.flush()
                for record in result['records']:
                    db.add(ScheduleRecord(id=uuid.uuid4().hex,document_id=document_id,term_id=result['term']['id'],course_name=record['courseName'],payload=dump(record)))
                source.status='saved';source.result=dump(dict(result,pages=json.loads(source.result)['pages']))
        except IntegrityError as exc:
            with self.connect() as db:
                source=db.get(ScheduleDraft,identifier)
                duplicate=source and db.scalar(select(ScheduleDocument.id).where(ScheduleDocument.digest==source.digest))
            if duplicate:raise ApiError('此 PDF 已保存，不会重复入库。') from exc
            raise
    def edit_record(self,identifier,value):
        from .pdf_schedule import normalize_result
        with self.connect(True) as db:
            row=db.scalar(select(ScheduleRecord).where(ScheduleRecord.id==identifier).with_for_update())
            if not row:raise ApiError('开课记录不存在。',404)
            document=db.get(ScheduleDocument,row.document_id);term=json.loads(document.term)
            parts=normalize_result({'term':{'academicYear':term['academicYear'],'semester':term['semester'],'startDate':term['start'],'weeks':term['weeks']},'records':[value]})['records']
            for index,part in enumerate(parts):
                if index==0:row.course_name=part['courseName'];row.payload=dump(part)
                else:db.add(ScheduleRecord(id=uuid.uuid4().hex,document_id=row.document_id,term_id=row.term_id,course_name=part['courseName'],payload=dump(part)))
    def snapshot(self):
        with self.connect() as db:
            documents=[{'id':r.id,'filename':r.filename,'term':json.loads(r.term),'model':r.model,'createdAt':iso(r.created_at)} for r in db.scalars(select(ScheduleDocument).order_by(ScheduleDocument.created_at.desc())).all()]
            records=[dict(json.loads(r.payload),id=r.id,documentId=r.document_id,termId=r.term_id) for r in db.scalars(select(ScheduleRecord).order_by(ScheduleRecord.term_id.desc(),ScheduleRecord.course_name)).all()]
            drafts=[{'id':r.id,'filename':r.filename,'status':r.status,'message':r.error} for r in db.scalars(select(ScheduleDraft).where(ScheduleDraft.status!='saved').order_by(ScheduleDraft.created_at.desc()).limit(10)).all()]
        return {'documents':documents,'records':records,'drafts':drafts}