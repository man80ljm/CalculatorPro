const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const test=require('node:test');
const source=fs.readFileSync('web_app/static/app.js','utf8');
const script=source.slice(source.indexOf('let syllabusWatch = 0;'),source.indexOf('document.getElementById("wizardRecheck").addEventListener'));
function setup(responder) {
  const nodes=new Map(),calls=[],storage=new Map(),drafts=[];
  const get=id=>{
    if(!nodes.has(id)) nodes.set(id,{value:0,checked:true,files:[],textContent:'',hidden:false,disabled:false,
      handlers:{},addEventListener(name,callback){this.handlers[name]=callback;}});
    return nodes.get(id);
  };
  const context=vm.createContext({
    document:{getElementById:get},wizard:null,courseId:4,Date,FormData,Response,encodeURIComponent,
    sessionStorage:{setItem:(k,v)=>storage.set(k,v),getItem:k=>storage.get(k),removeItem:k=>storage.delete(k)},
    flushSave:async()=>{},delay:async()=>{},
    errorMessage:async response=>(await response.json()).detail||'failed',
    showWizardDraft:draft=>{drafts.push(draft);context.wizard={completionJobId:draft.completion_job_id};},
    api:async(path,options)=>{calls.push({path,options});return new Response(JSON.stringify(await responder(path,options)),{status:options?.method==='POST'?202:200});}
  });
  vm.runInContext(script,context);
  return {node:get,calls,context,storage,drafts,click:id=>get(id).handlers.click()};
}
test('file validation reports an error without starting a background task',async()=>{
  const page=setup(()=>{throw new Error('should not call');});
  await page.click('syllabusReadBtn');
  assert.match(page.node('wizardError').textContent,/请选择/);
  page.node('syllabusFile').files=[new File(['x'],'bad.xls')];
  await page.click('syllabusReadBtn');
  assert.equal(page.calls.length,0);
});
test('background progress finishes with a saved draft and API count',async()=>{
  let poll=0;
  const page=setup(path=>{
    if(path==='/api/syllabus/imports') return {job_id:'job1'};
    if(++poll===1) return {status:'running',percent:45,api_calls:1,stage_label:'AI 正在补全'};
    return {status:'success',percent:100,api_calls:1,stage_label:'已完成',draft:{completion_job_id:'job1',source_course_id:4,expected_version:2}};
  });
  await page.click('wizardSavedRead');
  const submitted=JSON.parse(page.calls[0].options.body);
  assert.deepEqual(submitted,{course_id:4,allow_completion:true});
  assert.equal(page.context.wizard.expectedVersion,2);
  assert.equal(page.context.wizard.sourceCourseId,4);
  assert.equal(page.node('wizardReadProgress').value,100);
  assert.equal(page.node('wizardWait').hidden,true);
  assert.equal(page.storage.get('calculatorpro.syllabusJob'),'job1');
});
test('closing a dialog while a request is pending cannot replace a later draft',async()=>{
  let release;
  const pending=new Promise(resolve=>{release=resolve;});
  const page=setup(async()=>{await pending;return {job_id:'old-job'};});
  const opening=page.click('wizardSavedRead');
  await Promise.resolve();await Promise.resolve();
  vm.runInContext('syllabusWatch++',page.context);
  release();await opening;
  assert.equal(page.drafts.length,0);
  assert.equal(page.calls.length,1);
});
