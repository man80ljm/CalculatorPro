const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const test = require('node:test');
const script = fs.readFileSync('web_app/static/lesson_plans.js', 'utf8');
const pageHtml = fs.readFileSync('web_app/static/index.html', 'utf8');
const html = pageHtml.slice(pageHtml.indexOf('id="dlgLessonPlan"'), pageHtml.indexOf('id="dlgCalendar"'));

function plan(sourceId = 'source-one', name = '合成课程') {
  return {
    id: `plan-${sourceId}`, source_id: sourceId, revision: 2, job: null, warnings: [], pages: {}, pages_current: false, download_url: null,
    limits: { '教学目标': [120,180], '教学内容': [220,300], '教学手段与方法': [100,150], '思考讨论作业': [50,80] },
    context: { '课程名称': name, '学年起': '2026', '学年止': '2027', '学期': '1', '教研室': '', '授课对象': '', '授课教师': '合成教师',
      '学习资料': [], '学习网站': [], '评价': { '评价方式': '已编辑的评价' },
      lessons: [{ '课次': 1, '授课题目': '合成章节', '学时': '2', '教学目标': ['教师已编辑的目标'], '教学内容': ['教师已编辑的内容'], '教学手段与方法': [], '思考讨论作业': [] }],
    },
  };
}
function setup({ course = 1, saved = { id:'source-one' }, existing = plan(), responder } = {}) {
  const nodes = new Map(), calls = [];
  class Element {
    constructor(tag='div') { this.tagName=tag.toUpperCase(); this.children=[]; this.handlers={}; this.hidden=false; this.disabled=false; this.value=''; this.files=[]; this.textContent=''; this.classList={toggle(){}}; }
    append(...children) { this.children.push(...children); }
    replaceChildren(...children) { this.children=children; }
    addEventListener(name,callback) { const before=this.handlers[name]; this.handlers[name]=before ? async(event)=>{ const first=before(event); const second=callback(event); await Promise.all([first,second]); } : callback; }
    querySelectorAll(selector) {
      const tags = new Set(selector.toUpperCase().split(',').map(tag=>tag.trim()));
      const result=[];
      const visit=node=>{ for(const child of node.children) { if(tags.has(child.tagName)) result.push(child); visit(child); } };
      visit(this); return result;
    }
    click() { if(this.handlers.click) return this.handlers.click(); }
  }
  function node(id) {
    if(!nodes.has(id)) nodes.set(id,new Element(id==='lpOutlineFile'?'input':id==='lpLessonSelect'?'select':['lpMaterials','lpWebsites'].includes(id)?'textarea':'div'));
    return nodes.get(id);
  }
  node('lpEditor').append(...['lpMetaFields','lpLessonSelect','lpLessonFields','lpMaterials','lpWebsites','lpEvaluation','lpTeacherChoice'].map(node));
  const context=vm.createContext({
    document:{ getElementById:node, createElement:tag=>new Element(tag) }, window:{}, courseId:course, sessionStopped:false,
    Date, FormData, Response, URLSearchParams, delay:async()=>{throw new Error('opening must not poll without a running task');}, openDialog(){},
    errorMessage:async(response)=>(await response.json()).detail || '读取失败',
    api:async(path,options)=>{
      calls.push({path,options});
      if(responder) return new Response(JSON.stringify(await responder(path,options)));
      if(path.endsWith('/lesson-plan-outline')) return new Response(JSON.stringify({source:saved,message:'此课程尚未保存原始 Word 大纲，请上传一次后复用。'}));
      if(path.endsWith('/lesson-plans/current')) return new Response(JSON.stringify({plan:existing}));
      if(path==='/api/lesson-plans') return new Response(JSON.stringify(plan(saved.id)));
      throw new Error(`Unexpected request ${path}`);
    },
  });
  vm.runInContext(script,context);
  return {node,calls,context,open:()=>node('lessonPlanBtn').click()};
}

test('saved outline and edited plan open directly without source controls or AI calls',async()=>{
  const page=setup(); await page.open();
  assert.equal(page.node('lpSourcePanel').hidden,true);
  assert.equal(page.node('lpEditor').hidden,false);
  assert.equal(page.node('lpLoading').hidden,true);
  assert.equal(page.node('lpGenerate').disabled,false);
  assert.equal(page.node('lpLessonFields').children[0].children[1].value,'教师已编辑的目标');
  assert.equal(page.calls.length,2);
  assert(!html.includes('lpSourceChoice'));
  assert(!html.includes('>读取大纲<'));
});

test('first opening with a saved outline automatically prepares a reusable preview',async()=>{
  const page=setup({existing:null}); await page.open();
  const request=page.calls.find(call=>call.path==='/api/lesson-plans');
  assert.deepEqual(JSON.parse(request.options.body),{source_id:'source-one',reuse:true});
  assert.equal(page.node('lpSourcePanel').hidden,true);
  assert.equal(page.node('lpEditor').hidden,false);
  assert(!page.calls.some(call=>call.path.endsWith('/generate')));
});

test('old course asks only for a missing original file, then validates the upload',async()=>{
  const page=setup({saved:null,existing:null}); await page.open();
  assert.equal(page.node('lpSourcePanel').hidden,false);
  assert(page.node('lpSourceHint').textContent.includes('上传'));
  assert.equal(page.node('lpEditor').hidden,true);
  assert.equal(page.node('lpGenerate').disabled,true);
  await page.node('lpReadOutline').click();
  assert(page.node('lpError').textContent.includes('.docx'));
  assert.equal(page.calls.length,2);
});

test('slow response from the previous course cannot replace the current course preview',async()=>{
  let release;
  const held=new Promise(resolve=>{release=resolve;});
  const page=setup({responder:async(path)=>{
    if(path==='/api/courses/1/lesson-plan-outline') return held;
    if(path==='/api/courses/1/lesson-plans/current') return {plan:plan('one','上一课程')};
    if(path==='/api/courses/2/lesson-plan-outline') return {source:{id:'two'}};
    if(path==='/api/courses/2/lesson-plans/current') return {plan:plan('two','当前课程')};
    throw new Error('stale opening must not create a preview');
  }});
  const before=page.open(); page.context.courseId=2; await page.open(); release({source:{id:'one'}}); await before;
  assert.equal(page.node('lpMetaFields').children[0].children[1].value,'当前课程');
  assert.equal(page.node('lpSourcePanel').hidden,true);
});


test('a missing timetable allows direct manual entry and submits explicit metadata edits',async()=>{
  const preview=plan(); preview.teaching_assignment={...assignment('missing','',[]),options:[],override_revision:0,college:'合成学院',
    editable:{teacher:true,audience:true,teaching_unit:true},manual_fields:[],has_manual:false,message:'可手动填写'};
  let submitted;
  const page=setup({responder:async(path,options)=>{
    if(path.endsWith('/lesson-plan-outline')) return {source:{id:preview.source_id}};
    if(path.endsWith('/lesson-plans/current')) return {plan:preview};
    if(path.startsWith('/api/teaching/teachers?')) return {teachers:[{name:'目录教师',college:'合成学院',teaching_unit:'设计'}]};
    if(path==='/api/lesson-plans/'+preview.id) {
      submitted=JSON.parse(options.body);
      const current=plan(); current.context=submitted.context;current.revision=3;
      current.teaching_assignment={...preview.teaching_assignment,status:'manual',has_manual:true,manual_fields:['teacher','audience'],
        teacher:'目录教师',classes:['新班级'],teacher_info:{status:'ready',teaching_unit:'设计',unit_source:'department'},override_revision:1};
      return current;
    }
    throw new Error('Unexpected '+path);
  }});
  await page.open();
  const fields=page.node('lpMetaFields').children;
  assert.equal(fields[4].children[1].readOnly,false);
  assert.equal(fields[5].children[1].readOnly,false);
  assert.equal(fields[6].children[1].readOnly,false);
  const teacher=fields[6].children[1], audience=fields[5].children[1];
  teacher.value='目录教师';await teacher.handlers.input();
  audience.value='新班级';audience.handlers.input();
  assert.equal(fields[4].children[1].value,'设计');
  await page.node('lpSave').click();
  assert.deepEqual(submitted.teaching_metadata,{mode:'manual',revision:0,fields:{teacher:'目录教师',audience:'新班级'}});
  assert.equal(page.node('lpMetaFields').children[6].children[1].value,'目录教师');
  assert.equal(page.node('lpUseTimetable').hidden,false);
});
test('manual mode can override an unresolved timetable choice without blocking the generate button',async()=>{
  const preview=plan();preview.teaching_assignment=assignment('selection_required','',[]);
  const page=setup({existing:preview});await page.open();
  assert.equal(page.node('lpGenerate').disabled,true);
  await page.node('lpUseManual').click();
  assert.equal(page.node('lpMetaFields').children[6].children[1].readOnly,false);
  assert.equal(page.node('lpGenerate').disabled,false);
});

function assignment(status = 'ready', teacher = '课表教师', classes = ['课表A班']) {
  return { status, teacher, classes, options: [{teacher:'课表教师',classes:['课表A班']},{teacher:'另一教师',classes:['另一班级']}], message:'任课教师、授课对象已从后台课表自动填入。' };
}
test('timetable metadata is read-only and an unresolved teacher choice blocks generation',async()=>{
  const preview=plan(); preview.teaching_assignment=assignment('selection_required','',[]);
  const page=setup({existing:preview}); await page.open();
  const fields=page.node('lpMetaFields').children;
  assert.equal(fields[4].children[1].readOnly,true);
  assert.equal(fields[5].children[1].readOnly,true); assert.equal(fields[6].children[1].readOnly,true);
  assert.equal(fields[5].children[1].value,''); assert.equal(fields[6].children[1].value,'');
  assert.equal(page.node('lpTeacherChoiceWrap').hidden,false);
  assert.equal(page.node('lpGenerate').disabled,true); assert.equal(page.node('lpExport').disabled,true);
  assert.equal(page.node('lpSave').disabled,false);
  await page.node('lpGenerate').click();
  assert(!page.calls.some(call=>call.path.endsWith('/generate')));
});
test('choosing a timetable teacher fills only their classes and sends the remembered choice with edits',async()=>{
  const preview=plan(); preview.teaching_assignment=assignment('selection_required','',[]);
  preview.teaching_assignment.options[1].teacher_info={status:'ready',teaching_unit:'传达与媒体设计',unit_source:'department'};
  let submitted;
  const page=setup({responder:async(path,options)=>{
    if(path.endsWith('/lesson-plan-outline')) return {source:{id:preview.source_id}};
    if(path.endsWith('/lesson-plans/current')) return {plan:preview};
    if(path===`/api/lesson-plans/${preview.id}`) {
      submitted=JSON.parse(options.body);
      const result=plan(); result.context=submitted.context; result.revision=3; result.teaching_assignment=assignment('ready','另一教师',['另一班级']); result.teaching_assignment.teacher_info=preview.teaching_assignment.options[1].teacher_info; return result;
    }
    throw new Error(`Unexpected request ${path}`);
  }});
  await page.open(); page.node('lpTeacherChoice').value='另一教师';
  page.node('lpTeacherChoice').handlers.change({target:page.node('lpTeacherChoice')});
  assert.equal(page.node('lpMetaFields').children[4].children[1].value,'传达与媒体设计');
  assert.equal(page.node('lpMetaFields').children[5].children[1].value,'另一班级');
  assert.equal(page.node('lpMetaFields').children[6].children[1].value,'另一教师');
  assert.equal(page.node('lpGenerate').disabled,false);
  await page.node('lpSave').click();
  assert.equal(submitted.teacher_choice,'另一教师');
  assert.equal(submitted.context['授课对象'],'另一班级');
  assert.equal(submitted.context['教研室'],'传达与媒体设计');
  assert.deepEqual(submitted.context.lessons[0]['教学目标'],['教师已编辑的目标']);
});
test('changing the academic period refreshes timetable data without losing edited lesson text',async()=>{
  const preview=plan(); preview.teaching_assignment=assignment();
  const page=setup({responder:async(path)=>{
    if(path.endsWith('/lesson-plan-outline')) return {source:{id:preview.source_id}};
    if(path.endsWith('/lesson-plans/current')) return {plan:preview};
    if(path.includes('/teaching-assignment?')) {
      assert(path.includes('year_start=2025') && path.includes('year_end=2026') && path.includes('semester=1'));
      return {...assignment('ready','另一学期教师',['另一学期班级']),teacher_info:{status:'ready',teaching_unit:'环境设计',unit_source:'department'}};
    }
    throw new Error(`Unexpected request ${path}`);
  }});
  await page.open();
  const start=page.node('lpMetaFields').children[1].children[1], end=page.node('lpMetaFields').children[2].children[1];
  start.value='2025'; start.handlers.input(); end.value='2026'; end.handlers.input();
  assert.equal(page.node('lpGenerate').disabled,true);
  await end.handlers.change();
  assert.equal(page.node('lpMetaFields').children[4].children[1].value,'环境设计');
  assert.equal(page.node('lpMetaFields').children[5].children[1].value,'另一学期班级');
  assert.equal(page.node('lpMetaFields').children[6].children[1].value,'另一学期教师');
  assert.equal(page.node('lpLessonFields').children[0].children[1].value,'教师已编辑的目标');
  assert.equal(page.node('lpGenerate').disabled,false);
});

test('a similar timetable course is offered for explicit confirmation and then shows its teachers',async()=>{
  const preview=plan();
  preview.teaching_assignment={...assignment('missing','',[]),options:[],course_matches:['3D设计软件基础'],
    matched_course_name:'',override_revision:0,message:'课表中有相近课程名',editable:{teacher:true,audience:true,teaching_unit:true}};
  let submitted;
  const page=setup({existing:preview,responder:async(path,options)=>{
    if(path.endsWith('/lesson-plan-outline')) return {source:{id:preview.source_id}};
    if(path.endsWith('/lesson-plans/current')) return {plan:preview};
    if(path==='/api/lesson-plans/'+preview.id) {
      submitted=JSON.parse(options.body);
      const current=plan(); current.context=submitted.context; current.revision=3;
      current.teaching_assignment={...assignment('selection_required','',[]),course_matches:['3D设计软件基础'],
        matched_course_name:'3D设计软件基础',override_revision:1,message:'请选择本学期任课教师'};
      return current;
    }
    throw new Error('Unexpected '+path);
  }});
  await page.open();
  assert(html.includes('id="lpCourseMatch"'));
  assert.equal(page.node('lpCourseMatchWrap').hidden,false);
  assert.equal(page.node('lpCourseMatch').children[1].textContent,'3D设计软件基础');
  page.node('lpCourseMatch').value='3D设计软件基础';
  await page.node('lpCourseMatch').handlers.change({target:page.node('lpCourseMatch')});
  assert.deepEqual(submitted.teaching_metadata,{mode:'course_match',revision:0,course_name:'3D设计软件基础'});
  assert.equal(page.node('lpTeacherChoiceWrap').hidden,false);
  assert.equal(page.node('lpGenerate').disabled,true);
  assert.equal(page.node('lpMetaFields').children[0].children[1].value,'合成课程');
});
