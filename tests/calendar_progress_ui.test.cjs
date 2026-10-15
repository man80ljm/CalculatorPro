const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const test = require('node:test');
const source = fs.readFileSync('web_app/static/app.js','utf8');
const block = source.slice(source.indexOf('let actionBusy = false;'), source.indexOf('function setBusy(busy)'));
const html = fs.readFileSync('web_app/static/index.html','utf8');
const bindings = source.split('\n').filter(line => /getElementById\("calendar(Btn|Retry|CourseChoice|ChooseCourse|SkipCourseMatch|TeacherChoice|ChooseTeacher)"\)\.addEventListener/.test(line) || line.includes('window.addEventListener("pagehide"')).join('\n');
const bytes = Buffer.from('synthetic XLS bytes');
const frames = (summary={hours_match:true}) => [
  {type:'progress',percent:15,message:'正在准备模板…'},
  {type:'progress',percent:80,message:'正在生成 .xls 文件…'},
  {type:'file',filename:'合成测试.xls',size:bytes.length,summary},
  {type:'chunk',data:bytes.toString('base64')},{type:'complete'}
];
function response(events=frames(),fragment=0) {
  const data = new TextEncoder().encode(events.map(e => JSON.stringify(e)+'\n').join(''));
  const size = fragment || data.length;
  return new Response(new ReadableStream({start(controller){for(let i=0;i<data.length;i+=size)controller.enqueue(data.slice(i,i+size));controller.close();}}),{headers:{'Content-Type':'application/x-ndjson'}});
}
function setup(options={}) {
  const nodes = new Map(), calls=[], downloads=[], revoked=[], created=[], percentages=[], timers=new Map(), winHandlers={};
  let nextTimer=0;
  function node(id){
    if(!nodes.has(id)) {
      let value=0;
      const entry={id,hidden:false,disabled:false,textContent:'',dataset:{},attrs:{},handlers:{},open:false,
        setAttribute(key,v){this.attrs[key]=String(v);},removeAttribute(key){delete this.attrs[key];if(key==='href')delete this.href;if(key==='download')delete this.download;},
        append(...items){this.children=(this.children||[]).concat(items);},replaceChildren(...items){this.children=items;},addEventListener(name,callback){this.handlers[name]=callback;},querySelectorAll(){return [node('close0'),node('close1')];},
        click(){if(this.handlers.click)return this.handlers.click();downloads.push({url:this.href,filename:this.download});if(options.blockAuto && downloads.length===1)throw new Error('Browser blocked automatic download');}
      };
      Object.defineProperty(entry,'value',{get(){return value;},set(v){value=v;if(id==='calendarProgress')percentages.push(v);}});
      nodes.set(id,entry);
    }
    return nodes.get(id);
  }
  const addTimer=(callback,ms)=>{const id=++nextTimer;timers.set(id,{callback,ms});return id;};
  const context=vm.createContext({courseId:42,sessionStopped:false,courseFiles:[],document:{getElementById:node,createElement:()=>({value:'',textContent:''})},TextDecoder,Blob,Uint8Array,atob,AbortController,
    Date, URL:{createObjectURL(blob){const url='blob:synthetic-'+(created.length+1);created.push({url,blob});return url;},revokeObjectURL(url){revoked.push(url);}},
    setTimeout:addTimer,setInterval:addTimer,clearTimeout:id=>timers.delete(id),clearInterval:id=>timers.delete(id),
    window:{addEventListener:(name,callback)=>{winHandlers[name]=callback;}},flushSave:async()=>{},openDialog:id=>{node(id).open=true;},
    api:async(path,opts)=>{calls.push({path,opts});return response(frames(),3);},errorMessage:async()=> '课程资料不可用，请更新大纲。',filenameFromDisposition:()=> '原接口.xls',...options});
  vm.runInContext(block+'\n'+bindings,context);
  return {context,node,calls,downloads,revoked,created,percentages,timers,winHandlers};
}
const settle=()=>new Promise(setImmediate);

test('popup replaces upload flow and offers only XLS',()=>{
  assert(html.includes('id="dlgCalendar"') && html.includes('id="calendarDownload"'));
  assert(!html.includes('calendarFeedback') && !html.includes('calendarXlsxBtn') && !html.includes('下载 .xlsx'));
  assert(!source.includes('教学日历已下载'));
});
test('no course produces actionable error without sending request',async()=>{
  const t=setup({courseId:null});await t.context.downloadTeachingCalendar();
  assert.equal(t.calls.length,0);assert(t.node('calendarError').textContent.includes('选择已有课程'));
});
test('split UTF8 progress stream completes and keeps a manual download link',async()=>{
  const t=setup();await t.context.downloadTeachingCalendar();
  assert.equal(t.calls.length,1);assert.equal(t.calls[0].path,'/api/courses/42/teaching-calendar?progress=1');
  assert.equal(t.calls[0].opts.method,'POST');assert.equal(t.node('calendarProgress').value,100);
  assert.equal(t.node('calendarStage').textContent,'文件已生成');assert.equal(t.node('calendarDownload').hidden,false);
  assert.equal(t.downloads[0].filename,'合成测试.xls');assert.equal(t.revoked.length,0);
  assert.equal(t.timers.size,0);assert.equal(t.node('close0').disabled,false);assert.equal(t.node('calendarBtn').disabled,false);
  assert.equal(await t.created[0].blob.text(),bytes.toString());
  assert.equal(t.node('calendarDownloadHint').hidden,false);assert.equal(t.node('calendarError').hidden,true);
  await t.node('calendarDownload').click();await t.node('calendarDownload').click();
  assert.equal(t.calls.length,1);assert.equal(t.downloads.length,3);assert.equal(t.revoked.length,0);
});
test('automatic download rejection leaves completed file clickable',async()=>{
  const t=setup({blockAuto:true});await t.context.downloadTeachingCalendar();
  assert.equal(t.node('calendarProgress').value,100);assert.equal(t.node('calendarDialogError').hidden,true);
  assert.equal(t.node('calendarDownload').href,t.created[0].url);await t.node('calendarDownload').click();
  assert.equal(t.downloads.length,2);assert.equal(t.calls.length,1);
});
test('server stages stay below 100 until the whole file arrives',async()=>{
  let controller;
  const live=new Response(new ReadableStream({start(c){controller=c;}}),{headers:{'Content-Type':'application/x-ndjson'}});
  const t=setup({api:async()=>live});const pending=t.context.downloadTeachingCalendar();await settle();
  controller.enqueue(new TextEncoder().encode(JSON.stringify({type:'progress',percent:80,message:'生成文件'})+'\n'));await settle();
  assert.equal(t.node('calendarProgress').value,80);assert.equal(t.node('calendarDownload').hidden,true);assert.equal(t.node('close0').disabled,true);
  t.timers.values().next().value.callback();assert(t.node('calendarElapsed').textContent.includes('已等待'));
  controller.enqueue(new TextEncoder().encode(frames().slice(2).map(e=>JSON.stringify(e)+'\n').join('')));controller.close();await pending;
  assert.equal(t.node('calendarProgress').value,100);assert(t.percentages.every((p,i,a)=>i===0||p>=a[i-1]));
});
test('truncated response never exposes a download or falsely shows 100',async()=>{
  const t=setup({api:async()=>response(frames().slice(0,-1))});await t.context.downloadTeachingCalendar();
  assert(t.node('calendarProgress').value<100);assert.equal(t.node('calendarDownload').hidden,true);assert.equal(t.created.length,0);
  assert(t.node('calendarDialogError').textContent.includes('中断'));assert.equal(t.node('calendarRetry').hidden,false);assert.equal(t.timers.size,0);
});
test('size mismatch and invalid chunks are rejected',async()=>{
  for(const bad of [frames().map(e=>e.type==='file'?{...e,size:bytes.length+1}:e),frames().map(e=>e.type==='chunk'?{...e,data:'%%%'}:e)]) {
    const t=setup({api:async()=>response(bad)});await t.context.downloadTeachingCalendar();
    assert.equal(t.downloads.length,0);assert.equal(t.node('calendarDialogError').hidden,false);assert(t.node('calendarProgress').value<100);
  }
});
test('stream business error displays retry and does not invent completion',async()=>{
  const t=setup({api:async()=>response([{type:'progress',percent:15,message:'准备模板'},{type:'error',message:'转换超时，请重试。'}])});await t.context.downloadTeachingCalendar();
  assert.equal(t.node('calendarDialogError').textContent,'转换超时，请重试。');assert.equal(t.node('calendarStage').textContent,'生成未完成');assert.equal(t.downloads.length,0);
});
test('authentication and initial HTTP errors reset busy controls',async()=>{
  let t=setup({api:async()=>({ok:false})});await t.context.downloadTeachingCalendar();assert(t.node('calendarDialogError').textContent.includes('更新大纲'));
  t=setup();t.context.api=async()=>{t.context.sessionStopped=true;throw new Error('未登录');};await t.context.downloadTeachingCalendar();
  assert(t.node('calendarDialogError').textContent.includes('重新登录'));assert.equal(t.node('calendarRetry').hidden,true);assert.equal(t.node('close0').disabled,false);
});
test('duplicate clicks wait for save and issue only one request',async()=>{
  let release;const t=setup({flushSave:()=>new Promise(resolve=>{release=resolve;})});
  const first=t.node('calendarBtn').handlers.click();const second=t.node('calendarBtn').handlers.click();
  assert.equal(t.node('calendarBtn').disabled,true);assert.equal(t.node('dlgCalendar').open,true);release();await Promise.all([first,second]);assert.equal(t.calls.length,1);
});
test('course changes cannot receive the previous course file',async()=>{
  let release;const t=setup({flushSave:()=>new Promise(resolve=>{release=resolve;})});
  const pending=t.context.downloadTeachingCalendar();t.context.courseId=99;release();await pending;
  assert.equal(t.calls.length,0);assert(t.node('calendarDialogError').textContent.includes('课程已切换'));assert.equal(t.node('calendarDownload').hidden,true);
});
test('deadline ends the spinner and exposes a meaningful retry',async()=>{
  const t=setup({api:async(path,opts)=>new Promise((resolve,reject)=>{opts.signal.addEventListener('abort',()=>reject(new Error('aborted')));})});
  const pending=t.context.downloadTeachingCalendar();await settle();
  [...t.timers.values()].find(timer=>timer.ms===180000).callback();await pending;
  assert(t.node('calendarDialogError').textContent.includes('超时'));assert.equal(t.node('calendarBtn').disabled,false);assert.equal(t.node('close0').disabled,false);assert.equal(t.timers.size,0);
});
test('only regeneration or real page exit releases the retained file URL',async()=>{
  const t=setup();await t.context.downloadTeachingCalendar();t.winHandlers.pagehide({persisted:true});assert.equal(t.revoked.length,0);
  await t.node('calendarRetry').handlers.click();assert.equal(t.revoked[0],t.created[0].url);assert.equal(t.calls.length,2);
  assert.equal(t.node('calendarDownload').href,t.created[1].url);t.winHandlers.pagehide({persisted:false});assert.equal(t.revoked.length,2);assert.equal(t.node('calendarDownload').hidden,true);
});
test('old binary endpoint is compatible and hours warning appears in dialog',async()=>{
  const t=setup({api:async()=>new Response(bytes,{headers:{'Content-Type':'application/vnd.ms-excel','X-Calendar-Hours-Match':'false','X-Calendar-Total-Hours':'4','X-Calendar-Course-Hours':'48'}})});
  await t.context.downloadTeachingCalendar();assert.equal(t.node('calendarProgress').value,100);assert.equal(t.node('calendarDownload').download,'原接口.xls');
  assert(t.node('calendarDialogError').textContent.includes('4') && t.node('calendarDialogError').textContent.includes('48'));assert.equal(t.downloads.length,1);
});

test('multiple teachers are selected in the existing progress dialog and then generation resumes',async()=>{
  const requests=[];
  const teaching={status:'selection_required',teacher:'',classes:[],year_start:'2026',year_end:'2027',semester:'1',message:'请选择本学期的任课教师',options:[{teacher:'教师甲',classes:['A班']},{teacher:'教师乙',classes:['B班']}]};
  let generated=0;
  const t=setup({api:async(path,opts)=>{
    requests.push({path,opts});
    if(path.endsWith('/teaching-assignment')) {
      assert.deepEqual(JSON.parse(opts.body),{teacher:'教师乙',year_start:'2026',year_end:'2027',semester:'1'});
      return new Response(JSON.stringify({...teaching,status:'ready',teacher:'教师乙',classes:['B班']}));
    }
    if(++generated===1) return new Response(JSON.stringify({code:'teaching_assignment_required',detail:teaching.message,teaching_assignment:teaching}),{status:409});
    return response(frames({hours_match:true,teaching_assignment:{...teaching,status:'ready',teacher:'教师乙',classes:['B班']}}));
  }});
  await t.context.downloadTeachingCalendar();
  assert.equal(t.node('calendarAssignment').hidden,false); assert.equal(t.node('calendarStage').textContent,'请选择任课教师');
  assert.equal(t.node('calendarDownload').hidden,true); assert.equal(t.downloads.length,0);
  t.node('calendarTeacherChoice').value='教师乙'; t.node('calendarTeacherChoice').handlers.change({target:t.node('calendarTeacherChoice')});
  await t.node('calendarChooseTeacher').click();
  assert.equal(requests.length,3); assert.equal(t.node('calendarProgress').value,100);
  assert.equal(t.node('calendarAssignment').hidden,true); assert.equal(t.downloads.length,1);
});
test('missing timetable data stays visible as a warning beside the generated download',async()=>{
  const t=setup({api:async()=>response(frames({hours_match:true,teaching_assignment:{status:'missing',message:'后台课表未找到本课程授课安排，请联系管理员。'}}))});
  await t.context.downloadTeachingCalendar();
  assert.equal(t.node('calendarProgress').value,100); assert.equal(t.node('calendarDownload').hidden,false);
  assert(t.node('calendarDialogError').textContent.includes('后台课表未找到'));
});
test('similar course is confirmed before timetable teachers and calendar generation',async()=>{
  const calls=[];
  const missing={status:'missing',course_name:'3D软件基础',course_matches:['3D设计软件基础'],matched_course_name:'',
    teacher:'',classes:[],options:[],year_start:'2026',year_end:'2027',semester:'1',override_revision:0,message:'请确认'};
  let stage=0;
  const t=setup({api:async(path,opts)=>{
    calls.push({path,opts});
    if(path.endsWith('/teaching-metadata')) {
      assert.deepEqual(JSON.parse(opts.body),{mode:'course_match',course_name:'3D设计软件基础',revision:0,
        year_start:'2026',year_end:'2027',semester:'1'});
      return new Response(JSON.stringify({...missing,status:'selection_required',matched_course_name:'3D设计软件基础',
        options:[{teacher:'黄楚峰',classes:['25数字媒体A']},{teacher:'王阳子',classes:['25数字媒体B']}]}));
    }
    if(path.endsWith('/teaching-assignment')) return new Response(JSON.stringify({...missing,status:'ready',teacher:'黄楚峰',classes:['25数字媒体A']}));
    if(stage++===0) return new Response(JSON.stringify({code:'teaching_course_match_required',detail:'请确认课程',teaching_assignment:missing}),{status:409});
    return response(frames());
  }});
  await t.context.downloadTeachingCalendar();
  assert.equal(t.node('calendarCourseMatch').hidden,false);
  assert.equal(t.node('calendarDownload').hidden,true);
  assert.equal(t.node('calendarCourseChoice').children[1].textContent,'3D设计软件基础');
  t.node('calendarCourseChoice').value='3D设计软件基础';
  t.node('calendarCourseChoice').handlers.change({target:t.node('calendarCourseChoice')});
  await t.node('calendarChooseCourse').click();
  assert.equal(t.node('calendarAssignment').hidden,false);
  assert.equal(t.downloads.length,0);
  t.node('calendarTeacherChoice').value='黄楚峰';
  t.node('calendarTeacherChoice').handlers.change({target:t.node('calendarTeacherChoice')});
  await t.node('calendarChooseTeacher').click();
  assert.equal(calls.length,4);
  assert.equal(t.node('calendarProgress').value,100);
  assert.equal(t.downloads.length,1);
});

test('teacher can skip an unrelated similar course without binding it',async()=>{
  const missing={status:'missing',teacher:'',classes:[],course_matches:['3D设计软件基础'],
    year_start:'2026',year_end:'2027',semester:'1',override_revision:0};
  const calls=[];
  const t=setup({api:async(path)=>{
    calls.push(path);
    if(calls.length===1) return new Response(JSON.stringify({code:'teaching_course_match_required',detail:'请确认课程',teaching_assignment:missing}),{status:409});
    return response(frames({hours_match:true,teaching_assignment:{status:'missing',message:'未关联课表'}}));
  }});
  await t.context.downloadTeachingCalendar();
  await t.node('calendarSkipCourseMatch').click();
  assert.equal(calls[1],'/api/courses/42/teaching-calendar?progress=1&skip_course_match=1');
  assert.equal(t.node('calendarDownload').hidden,false);
  assert.equal(t.downloads.length,1);
});
