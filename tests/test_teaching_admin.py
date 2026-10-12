"""Formal timetable/settings: access, encrypted persistence, parsing, atomic save and limits."""
import base64,io,json,time,threading
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch
import pytest
from fastapi.testclient import TestClient
from sqlalchemy import select,event
from sqlalchemy.orm import Session
from web_app.db import AdminApiConfiguration,ScheduleDocument,ScheduleDraft,ScheduleRecord,get_engine
from web_app.teaching.api_settings import ApiSettings,ApiError
from web_app.teaching.pdf_schedule import normalize_result,extract_pdf
from web_app.teaching.schedule_db import ScheduleDatabase

KEY='sk-private-offline-settings-fixture'

@pytest.fixture
def app(tmp_path,monkeypatch):
    monkeypatch.setenv('SECRET_KEY','teaching-test-secret')
    monkeypatch.setenv('ALLOW_SQLITE','1')
    monkeypatch.setenv('DATABASE_URL','sqlite:///'+(tmp_path/'teaching.db').as_posix())
    monkeypatch.setenv('UPLOAD_DIR',str(tmp_path/'uploads'))
    monkeypatch.setenv('COOKIE_SECURE','false')
    monkeypatch.setenv('ADMIN_PASSWORD','admin-teaching-test-pass')
    for name in ('ADMIN_PASSWORD_HASH','ADMIN_HOST','ADMIN_USERNAME','DEEPSEEK_API_KEY','DEEPSEEK_KEYS_A','DEEPSEEK_KEYS_B','DEEPSEEK_KEYS_A_FILE','DEEPSEEK_KEYS_B_FILE'):
        monkeypatch.delenv(name,raising=False)
    from web_app.limiter import reset_limiters
    reset_limiters()
    from web_app.app import create_app
    return create_app()

@pytest.fixture
def admin(app):
    with TestClient(app) as client:
        assert client.post('/admin/api/login',json={'username':'admin','password':'admin-teaching-test-pass'}).status_code==200
        config=client.get('/admin/api/settings');assert config.status_code==200
        client.headers['X-Admin-CSRF']=config.json()['csrf']
        yield client

def fixture():
    return {'term':{'academicYear':'2026-2027','semester':1,'startDate':'2026-09-07','weeks':19},'records':[{'courseName':'Drawing','teachers':['TeacherA'],'classes':['ClassA'],'weeks':[1,2],'credits':3,'totalHours':48,'timePattern':'Tuesday','sourcePage':1,'sourceText':'Drawing TeacherA ClassA','issues':[]}],'notes':[]}

def pdf_bytes():
    from reportlab.pdfgen import canvas
    data=io.BytesIO();p=canvas.Canvas(data,pagesize=(300,200));p.setFont('Helvetica',10)
    p.drawString(20,180,'2026-2027 semester 1 Drawing TeacherA ClassA')
    for x in [20,70,170,220,270]:p.line(x,60,x,150)
    p.line(120,120,120,150)
    for y in [60,120,150]:p.line(20,y,270,y)
    for x,value in zip([25,75,125,175,225],['Class','1','2','3','4']):p.drawString(x,133,value)
    p.drawString(25,90,'ClassA');p.drawString(75,90,'Drawing');p.save();return data.getvalue()

class Response(io.BytesIO):
    def __enter__(self):return self
    def __exit__(self,*args):self.close()
class OfflineProvider:
    def __init__(self):self.calls=[]
    def open(self,request,timeout):
        self.calls.append(request.full_url)
        assert request.full_url.startswith('https://api.deepseek.com/')
        if request.data is None:return Response(json.dumps({'data':[{'id':'deepseek-flash'}]}).encode())
        body=json.loads(request.data);assert body['response_format']=={'type':'json_object'}
        source=json.loads(body['messages'][1]['content'].split('json：',1)[1])
        items=[{'id':u['id'],'courses':[{'name':'Drawing','teachers':['TeacherA'],'time':'Tuesday'}]} for u in source['units']]
        return Response(json.dumps({'choices':[{'finish_reason':'stop','message':{'content':json.dumps({'items':items})}}],'usage':{'total_tokens':10}}).encode())

@pytest.mark.parametrize('path',['/admin/api/settings','/admin/api/schedule','/admin/api/schedule/jobs/'+'a'*32])
def test_admin_read_permissions(app,path):
    with TestClient(app) as client:assert client.get(path).status_code==401

def test_settings_and_schedule_writes_require_admin_and_csrf(app,admin):
    for path in ('/admin/api/settings','/admin/api/settings/test','/admin/api/schedule/analyze','/admin/api/schedule/save','/admin/api/schedule/record'):
        with TestClient(app) as client:assert client.post(path,json={}).status_code==401
        assert admin.post(path,json={},headers={'X-Admin-CSRF':''}).status_code==403
        assert admin.post(path,json={},headers={'Origin':'https://evil.example'}).status_code==403
    assert admin.post('/admin/api/settings',json={'source':[],'model':'deepseek-flash'}).status_code==400
    assert admin.post('/admin/api/settings',json={'url':'http://internal.example'}).status_code==400
    assert admin.post('/admin/api/schedule/analyze',content='x'*17000000,headers={'Content-Type':'application/json'}).status_code==413

def test_teacher_cookie_cannot_read_timetables_or_configuration(app,admin):
    code=admin.post('/admin/api/invites',json={}).json()['invitation']['code']
    with TestClient(app) as teacher:
        assert teacher.post('/api/register',json={'username':'teacher-only','password':'teacher-only-pass','invite_code':code}).status_code==200
        assert teacher.post('/api/login',json={'username':'teacher-only','password':'teacher-only-pass'}).status_code==200
        assert teacher.get('/admin/api/schedule').status_code==401
        assert teacher.post('/admin/api/settings',json={'key':KEY}).status_code==401

def test_settings_are_encrypted_masked_and_survive_restart(app,admin):
    saved=admin.post('/admin/api/settings',json={'source':'custom','key':KEY,'model':'deepseek-flash'})
    assert saved.status_code==200 and KEY not in saved.text
    info=admin.get('/admin/api/settings');assert info.json()['settings']['configured'] and KEY not in info.text
    engine=get_engine()
    with Session(engine) as db:
        row=db.get(AdminApiConfiguration,1);assert KEY not in row.encrypted_config
    assert ApiSettings(engine).resolve({})['key']==KEY
    assert admin.post('/admin/api/settings',json={'source':'custom','key':'','model':'deepseek-flash'}).status_code==200
    assert ApiSettings(engine).resolve({})['key']==KEY

def test_project_source_includes_existing_multikey_pool_without_copying_keys(admin,monkeypatch):
    monkeypatch.setenv('DEEPSEEK_KEYS_A',KEY+',sk-second-offline-key')
    settings=ApiSettings(get_engine());assert settings.metadata()['projectAvailable']
    settings.save({'source':'project','model':'deepseek-flash'})
    assert settings.load()['key']==''
    assert settings.resolve({})['key']==KEY
    assert KEY not in json.dumps(settings.metadata())

def test_models_connection_and_last_test_persist(admin,monkeypatch):
    provider=OfflineProvider();monkeypatch.setattr('web_app.teaching.api_settings.urllib.request.build_opener',lambda *args:provider)
    assert admin.post('/admin/api/settings',json={'source':'custom','key':KEY,'model':'deepseek-flash'}).status_code==200
    result=admin.post('/admin/api/settings/test',json={'source':'custom','model':'deepseek-flash'})
    assert result.status_code==200 and result.json()['ok'] and KEY not in result.text
    assert provider.calls==['https://api.deepseek.com/models']
    assert ApiSettings(get_engine()).metadata()['lastTest']['ok']
    for _ in range(4):assert admin.post('/admin/api/settings/test',json={}).status_code==200
    assert admin.post('/admin/api/settings/test',json={}).status_code==429

def test_pdf_real_parsing_background_review_and_transactional_save(app,admin,monkeypatch):
    provider=OfflineProvider();monkeypatch.setattr('web_app.teaching.provider.urllib.request.build_opener',lambda *args:provider)
    assert admin.post('/admin/api/settings',json={'source':'custom','key':KEY,'model':'deepseek-flash'}).status_code==200
    raw=pdf_bytes();parsed=extract_pdf(raw)
    cell=next(c for c in parsed[0]['tables'][0]['cells'] if c['text']=='Drawing')
    assert cell['columnSpan']==2 and cell['coveredWeeks']==[1,2] and cell['classNames']==['ClassA']
    payload={'filename':'fixture.pdf','pdf':base64.b64encode(raw).decode()}
    response=admin.post('/admin/api/schedule/analyze',json=payload);assert response.status_code==202
    identifier=response.json()['jobId']
    deadline=time.monotonic()+10
    while time.monotonic()<deadline:
        job=admin.get('/admin/api/schedule/jobs/'+identifier).json()
        if job['status'] in {'ready','error'}:break
        time.sleep(.03)
    assert job['status']=='ready',job
    assert job['result']['records'][0]['weeks']==[1,2]
    assert admin.get('/admin/api/schedule').json()['records']==[]
    body={'jobId':identifier,'result':job['result']}
    saved=admin.post('/admin/api/schedule/save',json=body);assert saved.status_code==200
    assert len(saved.json()['documents'])==1 and len(saved.json()['records'])==1
    assert admin.post('/admin/api/schedule/save',json=body).status_code==200
    duplicate=admin.post('/admin/api/schedule/analyze',json=payload)
    assert duplicate.json()['duplicate'] and len(provider.calls)==1
    record=saved.json()['records'][0];record['teachers']=['TeacherA','TeacherB'];record['reviewed']=True
    edited=admin.post('/admin/api/schedule/record',json={'id':record['id'],'record':record})
    assert edited.status_code==200 and len(edited.json()['records'])==2
    assert all(len(r['teachers'])==1 for r in edited.json()['records'])
    assert len(ScheduleDatabase(get_engine()).snapshot()['records'])==2

def ready_repository():
    db=ScheduleDatabase(get_engine());identifier='b'*32
    db.create_draft(identifier,'fixture.pdf','0'*64,'deepseek-flash')
    result=dict(normalize_result(fixture()),pages=1);db.update_draft(identifier,'ready',result)
    return db,identifier,result

def test_database_save_rolls_back_and_concurrent_save_is_idempotent(app,admin):
    repo,identifier,result=ready_repository()
    def fail(db,context,instances):
        if any(isinstance(row,ScheduleRecord) for row in db.new):raise RuntimeError('simulated record failure')
    event.listen(Session,'before_flush',fail)
    try:
        with pytest.raises(RuntimeError):repo.save_draft(identifier,result)
    finally:event.remove(Session,'before_flush',fail)
    assert repo.snapshot()['documents']==[] and repo.draft(identifier)['status']=='ready'
    with ThreadPoolExecutor(max_workers=4) as pool:list(pool.map(lambda _:repo.save_draft(identifier,result),range(4)))
    assert len(repo.snapshot()['documents'])==1 and len(repo.snapshot()['records'])==1

def test_restart_recovers_partial_jobs_preserving_ready_drafts_and_chunks(app,admin):
    repo,identifier,result=ready_repository();repo.create_draft('c'*32,'interrupted.pdf','1'*64,'deepseek-flash')
    repo.save_chunk('a'*64,{'items':[]},{'total_tokens':1})
    restarted=ScheduleDatabase(get_engine());restarted.recover_interrupted()
    assert restarted.draft('c'*32)['status']=='error'
    assert restarted.draft(identifier)['status']=='ready'
    assert restarted.read_chunk('a'*64)['usage']['total_tokens']==1

def test_settings_and_timetable_pages_and_host_mapping(app,admin,monkeypatch):
    for page in ('schedule','settings'):
        response=admin.get('/admin/'+page);assert response.status_code==200
        assert 'static/'+('schedule' if page=='schedule' else 'settings')+'.js' in response.text
        assert '本地演示' not in response.text and '原型' not in response.text
    assert admin.get('/admin/static/teaching-ui.js').status_code==200
    assert admin.get('/admin/static/api-settings.sealed').status_code==404
    monkeypatch.setenv('ADMIN_HOST','admin.calc.geekhuang.com')
    assert admin.get('/schedule',headers={'Host':'admin.calc.geekhuang.com'}).status_code==200
    assert admin.get('/api/settings',headers={'Host':'admin.calc.geekhuang.com'}).status_code==200

def test_independent_key_requests_share_global_ai_ceiling(monkeypatch):
    from web_app import deepseek_pool as pool
    monkeypatch.setenv('DEEPSEEK_API_KEY',KEY)
    monkeypatch.setenv('DEEPSEEK_GLOBAL_CONCURRENCY','1')
    pool.reset_pool()
    with pool.external_slot():
        ticket=pool.reserve('syllabus')
        with pytest.raises(pool.QueueTimeout):ticket.wait(.02)
    with pool.acquire('syllabus',wait_timeout=.1):pass

def test_connection_rotates_failed_keys_in_project_pool(admin,monkeypatch):
    import urllib.error
    monkeypatch.setenv("DEEPSEEK_KEYS_A", "sk-expired-offline-key,"+KEY)
    used=[]
    class Rotating(OfflineProvider):
        def open(self,request,timeout):
            used.append(request.get_header("Authorization"))
            if used[-1]=="Bearer sk-expired-offline-key":
                raise urllib.error.HTTPError(request.full_url,401,"denied",{},None)
            return super().open(request,timeout)
    provider=Rotating()
    monkeypatch.setattr("web_app.teaching.provider.urllib.request.build_opener",lambda *args:provider)
    response=admin.post("/admin/api/settings/test",json={"source":"project","model":"deepseek-flash"})
    assert response.status_code==200 and response.json()["ok"]
    assert used==["Bearer sk-expired-offline-key","Bearer "+KEY]
    assert KEY not in response.text
