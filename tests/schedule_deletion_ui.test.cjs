const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const test = require('node:test');
const source = fs.readFileSync('web_app/static/admin/schedule.js', 'utf8');
const ui = fs.readFileSync('web_app/static/admin/teaching-ui.js', 'utf8');
const flush = () => new Promise(setImmediate);
const term = { id:'2026-1', label:'2026–2027 学年第 1 学期' };
const documentRow = (id='doc-one') => ({id, filename:id+'.pdf', term, createdAt:'2026-10-14T01:00:00Z'});
const record = (id='record-one', documentId='doc-one', name='课程甲') => ({id,documentId,termId:term.id,courseName:name,teachers:['教师甲'],classes:['班级甲'],weeks:[1,2],issues:[]});
const snapshot = () => ({documents:[documentRow(),documentRow('doc-two')], records:[record(),record('record-two','doc-two','课程乙')],drafts:[],trash:[]});
const preview = overrides => ({scope:'records',label:'选中的课表记录',recordCount:1,documentCount:1,removedDocumentCount:1,filenames:['doc-one.pdf'],terms:[term.label],confirmation:'confirmed-hash',fallbackWarning:'删除后旧课表中的安排将重新生效。',...overrides});

async function setup(responder = () => snapshot()) {
  const nodes = new Map(), elements = new Set(), handlers = {}, calls=[];
  function node(id, attrs={}) {
    const el = {id:id||'',dataset:{},hidden:'hidden' in attrs,disabled:'disabled' in attrs,checked:'checked' in attrs,indeterminate:false,value:attrs.value||'',textContent:'',open:'open' in attrs,isConnected:true,_children:[],_handlers:{},classList:{toggle(){}},
      addEventListener(type,fn){(this._handlers[type] ||= []).push(fn);},
      showModal(){this.open=true;},
      close(){this.open=false;for(const fn of this._handlers.close||[])fn({});},
      focus(){},
      closest(){return this.dataset.action?this:null;},
    };
    for(const [key,value] of Object.entries(attrs))if(key.startsWith('data-'))el.dataset[key.slice(5).replace(/-([a-z])/g,(_,letter)=>letter.toUpperCase())]=value;
    Object.defineProperty(el,'innerHTML',{get(){return this._html||'';},set(value){
      this._html=value;
      const remove=child=>{child._children.forEach(remove);elements.delete(child);if(child.id&&nodes.get(child.id)===child)nodes.delete(child.id);};
      this._children.forEach(remove);this._children=[];
      const stack=[this],voidTags=new Set(['input','br','hr','img','meta','link']);
      for(const match of value.matchAll(/<(\/)?([a-z][\w-]*)\b([^>]*)>/g)){
        if(match[1]){if(stack.length>1&&stack.at(-1)._tag===match[2])stack.pop();continue;}
        const attrs={};for(const attr of match[3].matchAll(/([\w-]+)(?:="([^"]*)")?/g))attrs[attr[1]]=attr[2]||'';
        const child=node(attrs.id,attrs);child._tag=match[2];stack.at(-1)._children.push(child);
        if(!voidTags.has(match[2]))stack.push(child);
      }
    }});
    if(id)nodes.set(id,el);elements.add(el);return el;
  }
  function query(selector) {
    const selectors=selector.split(',').map(value=>value.trim().replace(/^#appDialog\s+/,''));
    return [...elements].filter(el=>selectors.some(value=>{
      const match=value.match(/^\[data-([\w-]+)(?:="([^"]*)")?\]$/);if(!match)return false;
      const key=match[1].replace(/-([a-z])/g,(_,letter)=>letter.toUpperCase());
      return key in el.dataset&&(match[2]===undefined||el.dataset[key]===match[2]);
    }));
  }
  for(const id of ['adminContent','pageTitle','appDialog','dialogContent','toast'])node(id);
  nodes.get('pageTitle').textContent='课表管理';
  const context=vm.createContext({window:{},document:{activeElement:null,getElementById:id=>nodes.get(id)||null,querySelector:selector=>query(selector)[0]||null,querySelectorAll:query,addEventListener:(type,fn)=>{(handlers[type]||=[]).push(fn);}},setTimeout:()=>0,clearTimeout(){},Date});
  vm.runInContext(ui,context);
  context.window.TeachingAdmin.api={request:async(path,body)=>{calls.push({path,body:body&&JSON.parse(JSON.stringify(body))});return await responder(path,body);}};
  vm.runInContext(source,context);await flush();
  async function fire(type,target){for(const fn of handlers[type]||[])fn({target});await flush();}
  return {calls,node:id=>nodes.get(id),query,context,
    click:async(action,id)=>{const control=query('[data-action="'+action+'"]').find(el=>id===undefined||el.dataset.id===id);assert(control,'missing control '+action);if(!control.disabled)await fire('click',control);},
    select:async(id,checked=true,mobile=false)=>{const inputs=query('[data-record-select="'+id+'"]');assert.equal(inputs.length,2);const input=inputs[mobile?1:0];input.checked=checked;await fire('change',input);},
    all:async(checked=true)=>{const control=nodes.get('pdfSelectAll');control.checked=checked;await fire('change',control);},
    filter:async(id,value)=>{const control=nodes.get(id);control.value=value;await fire(id==='pdfQuery'?'input':'change',control);},
  };
}

test('desktop and mobile record selection stay synchronized with accessible checkboxes',async()=>{
  const page=await setup();
  const html=page.node('pdfRecordTable').innerHTML;
  assert(html.includes('pdf-selectable'));
  assert(html.includes('<div class="data-cards">'));
  assert(html.includes('aria-label="选择 课程甲 · 教师甲 · 班级甲"'));
  await page.select('record-one',true,true);
  assert(page.query('[data-record-select="record-one"]').every(input=>input.checked));
  assert.equal(page.node('pdfSelectedCount').textContent,'已选 1 条');
  assert.equal(page.node('pdfSelectAll').indeterminate,true);
});

test('all selects only matching records and every filter change clears selection',async()=>{
  const page=await setup();
  await page.filter('pdfQuery','课程甲');await page.all();
  assert.equal(page.node('pdfSelectedCount').textContent,'已选 1 条');
  assert.equal(page.query('[data-record-select="record-two"]').length,0);
  for(const [id,value] of [['pdfQuery',''],['pdfStatus','complete'],['pdfTerm',term.id]]){
    await page.all();await page.filter(id,value);
    assert.equal(page.node('pdfSelectedCount').textContent,'已选 0 条');
    assert.equal(page.query('[data-action="pdf-delete-selected"]')[0].disabled,true);
  }
});

test('deletion requires fresh server preview and explicit confirmation, then clears selection',async()=>{
  let pendingDelete;
  const page=await setup((path,body)=>{
    if(path==='/schedule/delete-preview')return preview();
    if(path==='/schedule/delete')return new Promise(resolve=>{pendingDelete=resolve;});
    return snapshot();
  });
  await page.select('record-one');await page.click('pdf-delete-selected');
  assert.deepEqual(page.calls.find(call=>call.path==='/schedule/delete-preview').body,{scope:'records',ids:['record-one']});
  assert.equal(page.calls.filter(call=>call.path==='/schedule/delete').length,0);
  const dialog=page.node('dialogContent').innerHTML;
  assert(dialog.includes('1 条开课记录'));assert(dialog.includes(term.label));assert(dialog.includes('doc-one.pdf'));
  assert(dialog.includes('旧课表中的安排将重新生效'));assert(dialog.includes('用户课程、大纲、成绩、教师资料'));
  await page.click('pdf-delete-confirm');await page.click('pdf-delete-confirm');
  assert.equal(page.calls.filter(call=>call.path==='/schedule/delete').length,1);
  assert.equal(page.calls.at(-1).body.confirmation,'confirmed-hash');
  pendingDelete({...snapshot(),records:[record('record-two','doc-two','课程乙')],trash:[{id:'trash-one',label:'选中的课表记录',recordCount:1,createdAt:'2026-10-14T01:00:00Z',filenames:['doc-one.pdf']}]});await flush();
  assert.equal(page.node('pdfSelectedCount').textContent,'已选 0 条');assert.equal(page.node('appDialog').open,false);
  assert(page.node('adminContent').innerHTML.includes('最近删除（1 批）'));
});

test('whole-term clearing is unavailable for all terms and ignores narrower search filters',async()=>{
  const page=await setup(path=>path==='/schedule/delete-preview'?preview({scope:'term',recordCount:2,documentCount:2}):snapshot());
  assert.equal(page.query('[data-action="pdf-delete-term"]')[0].hidden,true);
  await page.filter('pdfTerm',term.id);await page.filter('pdfQuery','课程甲');
  assert.equal(page.query('[data-action="pdf-delete-term"]')[0].hidden,false);
  await page.click('pdf-delete-term');
  assert.deepEqual(page.calls.at(-1).body,{scope:'term',termId:term.id});
  assert(page.node('dialogContent').innerHTML.includes('包括当前搜索或状态筛选没有显示的记录'));
  assert(page.node('dialogContent').innerHTML.includes('2 条开课记录'));
});

test('import file deletion scopes to one PDF and escapes displayed file names',async()=>{
  const initial=snapshot();initial.documents[0].filename='<img src=x>.pdf';
  const page=await setup(path=>path==='/schedule/delete-preview'?preview({scope:'document',filenames:['<img src=x>.pdf']}):initial);
  assert(page.node('adminContent').innerHTML.includes('按导入文件管理'));
  assert(!page.node('adminContent').innerHTML.includes('<img'));
  await page.click('pdf-delete-document','doc-one');
  assert.deepEqual(page.calls.at(-1).body,{scope:'document',documentId:'doc-one'});
  assert(page.node('dialogContent').innerHTML.includes('&lt;img src=x&gt;.pdf'));
});

test('changed-data errors invalidate the old confirmation and require another preview',async()=>{
  let previews=0,deletes=0;
  const page=await setup(path=>{
    if(path==='/schedule/delete-preview')return preview({confirmation:'hash-'+(++previews)});
    if(path==='/schedule/delete'){deletes++;throw new Error('课表已变化，请重新核对删除范围。');}
    return snapshot();
  });
  await page.select('record-one');await page.click('pdf-delete-selected');await page.click('pdf-delete-confirm');
  assert(page.node('dialogContent').innerHTML.includes('课表已变化'));
  assert.equal(page.query('[data-action="pdf-delete-confirm"]').length,0);
  await page.click('pdf-delete-recheck');assert.equal(previews,2);
  await page.click('pdf-delete-confirm');assert.equal(deletes,2);
  assert.equal(page.calls.filter(call=>call.path==='/schedule/delete').at(-1).body.confirmation,'hash-2');
});

test('forbidden preview errors show the server message and never submit a deletion',async()=>{
  const page=await setup(path=>{if(path==='/schedule/delete-preview')throw new Error('权限已失效，请重新登录。');return snapshot();});
  await page.select('record-one');await page.click('pdf-delete-selected');
  assert(page.node('dialogContent').innerHTML.includes('权限已失效，请重新登录。'));
  assert.equal(page.calls.filter(call=>call.path==='/schedule/delete').length,0);
});

test('recent deletions can be restored once while pending and disappear after success',async()=>{
  const initial={documents:[],records:[],drafts:[],trash:[{id:'trash-one',label:'已删除学期',recordCount:2,filenames:['doc-one.pdf'],createdAt:'2026-10-14T01:00:00Z'}]};
  let complete;
  const page=await setup(path=>path==='/schedule/restore'?new Promise(resolve=>{complete=resolve;}):initial);
  await page.click('pdf-restore','trash-one');await page.click('pdf-restore','trash-one');
  assert.equal(page.calls.filter(call=>call.path==='/schedule/restore').length,1);
  assert.deepEqual(page.calls.at(-1).body,{id:'trash-one'});
  complete(snapshot());await flush();
  assert(!page.node('adminContent').innerHTML.includes('最近删除（'));
  assert(page.node('pdfRecordTable').innerHTML.includes('课程甲'));
});


test('older deletions append once without resetting current filters, selection, or expanded details',async()=>{
  const deleted=id=>({id,label:'删除批次 '+id,recordCount:1,filenames:['课表.pdf'],createdAt:'2026-10-14T01:00:00Z'});
  const initial={...snapshot(),trash:[deleted('new')],trashNextCursor:'cursor+/1'};
  let complete;
  const page=await setup(path=>path.startsWith('/schedule/trash?')?new Promise(resolve=>{complete=resolve;}):initial);
  await page.filter('pdfQuery','课程甲');await page.select('record-one');
  page.node('pdfTrashDetails').open=true;
  await page.click('pdf-trash-more');await page.click('pdf-trash-more');
  assert.equal(page.calls.filter(call=>call.path.startsWith('/schedule/trash?')).length,1);
  assert.equal(page.calls.at(-1).path,'/schedule/trash?before=cursor%2B%2F1');
  assert.equal(page.node('pdfTrashStatus').textContent,'正在加载更早的删除记录……');
  complete({items:[deleted('new'),deleted('old'),deleted('old')],nextCursor:null});await flush();
  assert.equal(page.node('pdfTrashDetails').open,true);
  assert.equal(page.node('pdfQuery').value,'课程甲');
  assert.equal(page.node('pdfSelectedCount').textContent,'已选 1 条');
  assert(page.query('[data-record-select="record-one"]').every(input=>input.checked));
  assert.equal(page.query('[data-action="pdf-restore"]').filter(control=>control.dataset.id==='old').length,1);
  assert.equal(page.query('[data-action="pdf-trash-more"]').length,0);
});

test('older-deletion loading errors remain local and retry the same cursor',async()=>{
  const initial={...snapshot(),trash:[{id:'new',label:'删除批次',recordCount:1,filenames:[],createdAt:'2026-10-14T01:00:00Z'}],trashNextCursor:'older'};
  let count=0;
  const page=await setup(path=>{if(path.startsWith('/schedule/trash?')){if(++count===1)throw new Error('暂时无法读取，请重试。');return {items:[],nextCursor:null};}return initial;});
  page.node('pdfTrashDetails').open=true;await page.click('pdf-trash-more');
  assert.equal(page.node('pdfTrashError').hidden,false);
  assert.equal(page.node('pdfTrashError').textContent,'暂时无法读取，请重试。');
  assert.equal(page.query('[data-action="pdf-trash-more"]')[0].disabled,false);
  await page.click('pdf-trash-more');
  assert.equal(count,2);assert.equal(page.node('pdfTrashError').hidden,true);
  assert.equal(page.node('pdfTrashDetails').open,true);
});
