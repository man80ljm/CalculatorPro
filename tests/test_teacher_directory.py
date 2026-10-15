"""Staff directory import/CRUD and safe lesson metadata linkage."""
import base64
import copy
import io
import json
from unittest.mock import Mock

import pytest
from sqlalchemy import select
from sqlalchemy.orm import Session
from fastapi.testclient import TestClient
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas
from tests.test_teaching_admin import app, admin
from tests.test_web_flow import client
from tests.test_course_calendar import create_course, calendar_data
from tests.test_teaching_assignments import add_schedule, make_plan
from web_app.db import TeacherProfile, TeacherImport, get_engine
from web_app.teaching.teacher_directory import TeacherDirectory, parse_pdf, teacher_info
from web_app.teaching.teacher_routes import upload_limiter


@pytest.fixture(autouse=True)
def reset_uploads():
    upload_limiter.reset()
    yield
    upload_limiter.reset()


def pdf_bytes(*, department="美术教育系", date="2026年5月", name="教师甲", phone="13800000001", duplicate=False):
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    stream=io.BytesIO(); document=canvas.Canvas(stream,pagesize=(600,800));document.setFont('STSong-Light',11)
    document.drawString(30,765,f"合成学院教职工联系表（{4 if duplicate else 3}人） {date}")
    for x in (30,325): document.drawString(x,735,'姓名 移动电话 办公室电话')
    document.drawString(30,705,f"{department}（{3 if duplicate else 2}人）")
    document.drawString(30,680,name+' '+phone+' 3690001')
    document.drawString(30,655,'教师乙 13800000002')
    if duplicate: document.drawString(30,630,name+' 13800000004')
    document.drawString(325,705,'办公室（1人）')
    document.drawString(325,680,'行政甲 13800000003 3690002')
    document.save();return stream.getvalue()


def upload(admin, data=None):
    response=admin.post('/admin/api/teachers/import/preview',json={'filename':'合成联系表.pdf','pdf':base64.b64encode(data or pdf_bytes()).decode()})
    assert response.status_code==200,response.text
    return response.json()


def confirm(admin,draft):
    response=admin.post('/admin/api/teachers/import/confirm',json={'id':draft['id'],'revision':draft['revision']})
    assert response.status_code==200,response.text
    return response.json()


def fields(name='教师甲',department='传达与媒体设计系',unit=''):
    return {'name':name,'college':'合成学院','department':department,'teaching_unit':unit,
            'group_name':department,'mobile':'13800000001','office_phone':'3690001','notes':'管理员备注'}


def test_pdf_column_groups_counts_and_source_date_are_preserved():
    result=parse_pdf('合成.pdf',pdf_bytes())
    assert len(result['rows'])==result['expected_count']==3 and result['warnings']==[]
    assert result['source_date']=='2026-05'
    rows={row['fields']['name']:row['fields'] for row in result['rows']}
    assert rows['教师甲']['department']=='美术教育系' and rows['教师甲']['teaching_unit']=='美术教育'
    assert rows['教师乙']['department']=='美术教育系' and rows['教师乙']['teaching_unit']=='美术教育'
    assert rows['行政甲']['department']==rows['行政甲']['teaching_unit']=='' and rows['行政甲']['group_name']=='办公室'
    assert rows['教师甲']['office_phone']=='3690001'
    assert [group['actual'] for group in result['groups']]==[2,1]


def test_import_is_reviewed_before_atomic_save_and_same_pdf_is_idempotent(admin,monkeypatch):
    forbidden=Mock(side_effect=AssertionError('staff import must not call AI'))
    monkeypatch.setattr('web_app.lesson_plans.engine.deepseek',forbidden)
    raw=pdf_bytes();draft=upload(admin,raw)
    assert draft['counts']=={'new':3,'update':0,'skip':0,'ambiguous':0}
    assert admin.get('/admin/api/teachers').json()['total']==0
    result=confirm(admin,draft);assert result['counts']['new']==3
    listing=admin.get('/admin/api/teachers').json()
    assert listing['total']==3 and listing['latest_import']['source_date']=='2026-05'
    assert admin.post('/admin/api/teachers/import/confirm',json={'id':draft['id'],'revision':draft['revision']}).json()['duplicate']
    assert upload(admin,raw)['status']=='imported'
    assert admin.get('/admin/api/teachers').json()['total']==3
    forbidden.assert_not_called()


def test_manual_edit_search_status_and_optimistic_revision(admin):
    response=admin.post('/admin/api/teachers/record',json={'fields':fields(unit='传达与媒体设计')})
    assert response.status_code==200,response.text
    listing=admin.get('/admin/api/teachers',params={'query':'教师 甲'}).json()
    assert listing['total']==1
    row=listing['records'][0];changed=fields(unit='新版教研室')
    assert admin.post('/admin/api/teachers/record',json={'id':row['id'],'revision':row['revision'],'fields':changed}).status_code==200
    assert admin.post('/admin/api/teachers/record',json={'id':row['id'],'revision':row['revision'],'fields':changed}).status_code==409
    row=admin.get('/admin/api/teachers',params={'query':'新版教研室'}).json()['records'][0]
    assert admin.post('/admin/api/teachers/status',json={'id':row['id'],'revision':row['revision'],'active':False}).status_code==200
    assert admin.get('/admin/api/teachers').json()['total']==0
    inactive=admin.get('/admin/api/teachers',params={'status':'inactive'}).json()['records'][0]
    with Session(get_engine()) as db: assert teacher_info(db,'教师甲','合成学院')['status']=='not_found'
    assert admin.post('/admin/api/teachers/status',json={'id':inactive['id'],'revision':inactive['revision'],'active':True}).status_code==200
    assert admin.get('/admin/api/teachers').json()['total']==1


def test_pdf_updates_match_existing_and_preserve_explicit_manual_unit_in_same_department(admin):
    original=upload(admin);confirm(admin,original)
    row=admin.get('/admin/api/teachers',params={'query':'教师甲'}).json()['records'][0]
    edited={key:row[key] for key in fields()};edited['teaching_unit']='管理员维护名称'
    admin.post('/admin/api/teachers/record',json={'id':row['id'],'revision':row['revision'],'fields':edited}).raise_for_status()
    newer=upload(admin,pdf_bytes(date='2026年6月',phone='13800000009'))
    assert newer['counts']['update']==3
    candidate=next(item for item in newer['rows'] if item['fields']['name']=='教师甲')
    assert candidate['fields']['teaching_unit']=='管理员维护名称'
    confirm(admin,newer)
    updated=admin.get('/admin/api/teachers',params={'query':'教师甲'}).json()['records'][0]
    assert updated['id']==row['id'] and updated['mobile']=='13800000009' and updated['teaching_unit']=='管理员维护名称'
    changed=upload(admin,pdf_bytes(date='2026年7月',department='环境设计系',phone='13800000009'))
    changed_row=next(item for item in changed['rows'] if item['fields']['name']=='教师甲')
    assert changed_row['fields']['teaching_unit']=='环境设计'
    confirm(admin,changed)
    assert admin.get('/admin/api/teachers',params={'query':'教师甲'}).json()['records'][0]['department']=='环境设计系'


def test_stale_preview_requires_refresh_before_saving(admin):
    draft=upload(admin)
    assert admin.post('/admin/api/teachers/record',json={'fields':fields(name='另一教师')}).status_code==200
    assert admin.post('/admin/api/teachers/import/confirm',json={'id':draft['id'],'revision':draft['revision']}).status_code==409
    refreshed=admin.post('/admin/api/teachers/import/review',json={'id':draft['id'],'revision':draft['revision'],'refresh':True})
    assert refreshed.status_code==200 and not refreshed.json()['is_stale']
    confirm(admin,refreshed.json())
    assert admin.get('/admin/api/teachers').json()['total']==4


def test_homonyms_are_not_silently_selected_and_import_requires_explicit_choice(admin):
    for unit in ('第一教研室','第二教研室'):
        response=admin.post('/admin/api/teachers/record',json={'fields':fields(department='美术教育系',unit=unit),'allow_duplicate':True})
        assert response.status_code==200,response.text
    with Session(get_engine()) as db: assert teacher_info(db,'教师甲','合成学院')['status']=='ambiguous'
    draft=upload(admin)
    assert draft['counts']['ambiguous']==1
    assert admin.post('/admin/api/teachers/import/confirm',json={'id':draft['id'],'revision':draft['revision']}).status_code==400
    item=next(row for row in draft['rows'] if row['fields']['name']=='教师甲')
    reviewed=admin.post('/admin/api/teachers/import/review',json={'id':draft['id'],'revision':draft['revision'],'index':item['index'],'fields':item['fields'],'action':'skip'})
    assert reviewed.status_code==200,reviewed.text
    confirm(admin,reviewed.json())
    assert admin.get('/admin/api/teachers').json()['total']==4


def test_duplicate_names_in_one_pdf_are_explicitly_reviewed(admin):
    draft=upload(admin,pdf_bytes(duplicate=True))
    assert draft['counts']['ambiguous']==2
    assert admin.post('/admin/api/teachers/import/confirm',json={'id':draft['id'],'revision':draft['revision']}).status_code==400
    assert admin.get('/admin/api/teachers').json()['total']==0


@pytest.mark.parametrize('filename,content,status',[('bad.docx',b'not-pdf',400),('bad.pdf',b'broken',400),('large.pdf',b'%PDF-'+b'x'*(5*1024*1024),413)],ids=['wrong_type','corrupt_pdf','too_large'])
def test_invalid_uploads_are_rejected(admin,filename,content,status):
    result=admin.post('/admin/api/teachers/import/preview',json={'filename':filename,'pdf':base64.b64encode(content).decode()})
    assert result.status_code==status,result.text
    assert admin.get('/admin/api/teachers').json()['total']==0


def test_admin_auth_csrf_origin_and_teacher_cookie_protect_staff_contacts(admin,client,app):
    assert client.get('/admin/api/teachers').status_code==401
    for route in ('import/preview','import/review','import/confirm','record','status'):
        assert client.post('/admin/api/teachers/'+route,json={}).status_code==401
        assert admin.post('/admin/api/teachers/'+route,json={},headers={'X-Admin-CSRF':''}).status_code==403
        assert admin.post('/admin/api/teachers/'+route,json={},headers={'Origin':'https://other.example'}).status_code==403
    for path in ('/admin/static/teachers.js','/admin/static/teachers.css','/admin/teachers'):
        assert admin.get(path).status_code==200
    assert client.get('/admin/teachers',follow_redirects=False).status_code==303


def test_lesson_metadata_uses_directory_unit_but_does_not_expose_contacts(admin,client):
    profile=admin.post('/admin/api/teachers/record',json={'fields':fields(unit='')})
    assert profile.status_code==200
    course=create_course(client,calendar_data(),name='合成教案课程')
    add_schedule('合成教案课程',[('教师甲',['A班'])])
    plan=make_plan(client,course)
    assert plan['context']['教研室']=='传达与媒体设计'
    serialized=json.dumps(plan['teaching_assignment'],ensure_ascii=False)
    for private in ('13800000001','3690001','mobile','office_phone','管理员备注'):
        assert private not in serialized
    context=copy.deepcopy(plan['context']);context['教研室']='前端伪造教研室'
    saved=client.patch(f"/api/lesson-plans/{plan['id']}",json={'revision':plan['revision'],'context':context})
    assert saved.status_code==200 and saved.json()['context']['教研室']=='传达与媒体设计'
    row=admin.get('/admin/api/teachers').json()['records'][0]
    update=fields(unit='管理员确认名称')
    admin.post('/admin/api/teachers/record',json={'id':row['id'],'revision':row['revision'],'fields':update}).raise_for_status()
    assert client.get(f"/api/lesson-plans/{plan['id']}").json()['context']['教研室']=='管理员确认名称'


def test_same_name_in_other_college_is_not_used(admin,client):
    entry=fields();entry['college']='另一学院'
    admin.post('/admin/api/teachers/record',json={'fields':entry}).raise_for_status()
    course=create_course(client,calendar_data(),name='合成教案课程')
    add_schedule('合成教案课程',[('教师甲',['A班'])])
    plan=make_plan(client,course)
    assert plan['context']['教研室']=='' and plan['teaching_assignment']['teacher_info']['status']=='not_found'


def test_manually_entered_unit_also_removes_department_suffix(admin):
    entry=fields(unit="传达与媒体设计系")
    assert admin.post('/admin/api/teachers/record',json={'fields':entry}).status_code==200
    row=admin.get('/admin/api/teachers').json()['records'][0]
    assert row['department']=='传达与媒体设计系' and row['teaching_unit']=='传达与媒体设计'
    with Session(get_engine()) as db: assert teacher_info(db,'教师甲','合成学院')['teaching_unit']=='传达与媒体设计'


def test_valid_pdf_larger_than_default_json_limit_can_be_previewed(admin):
    content=pdf_bytes()+b"\n%"+b" "*(1024*1024)
    draft=upload(admin,content)
    assert len(draft['rows'])==3 and not draft['warnings']
    assert admin.get('/admin/api/teachers').json()['total']==0
