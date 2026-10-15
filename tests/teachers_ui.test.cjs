const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const test = require('node:test');

const script = fs.readFileSync('web_app/static/admin/teachers.js', 'utf8');
const uiScript = fs.readFileSync('web_app/static/admin/teaching-ui.js', 'utf8');
const sharedCss = fs.readFileSync('web_app/static/admin/admin.css', 'utf8');

function record(id = 'teacher-one', name = '合成教师') {
  return { id, name, college:'合成学院', department:'传达与媒体设计系', teaching_unit:'传达与媒体设计',
    group_name:'', mobile:'13800000001', office_phone:'3690001', active:true, revision:1,
    source_filename:'合成联系表.pdf', source_date:'2026-05', notes:'' };
}
async function setup(responder) {
  const nodes = new Map(), handlers = {}, calls = [];
  const node = id => {
    if (!nodes.has(id)) nodes.set(id, { innerHTML:'', textContent:'', hidden:false, open:false,
      classList:{ toggle(){} }, addEventListener(){}, showModal(){this.open=true;}, close(){this.open=false;} });
    return nodes.get(id);
  };
  const context = vm.createContext({
    window:{}, document:{ activeElement:null, getElementById:node, addEventListener:(event,handler)=>{
      (handlers[event] ||= []).push(handler);
    } }, URLSearchParams, setTimeout, clearTimeout,
  });
  vm.runInContext(uiScript, context);
  context.window.TeachingAdmin.api = { request:async(path,body)=>{
    calls.push({path,body});
    return await responder(path,body);
  } };
  vm.runInContext(script, context);
  await new Promise(setImmediate);
  return { node, calls, click:async(action,data={})=>{
    const control = {disabled:false,dataset:{action,...data}};
    const event = {target:{closest:()=>control}};
    await Promise.all(handlers.click.map(handler=>handler(event)));
  } };
}
const listing = (records, page=1, total=records.length) => ({
  records, page, total, pages:Math.max(1,Math.ceil(total/20)), departments:['传达与媒体设计系'], pending:[],
});
const cards = html => [...html.matchAll(/<article class="data-card staff-card">([\s\S]*?)<\/article>/g)].map(match=>match[1]);

test('teacher records have mobile cards alongside the desktop table',async()=>{
  const page = await setup(()=>listing([record()]));
  const html = page.node('adminContent').innerHTML;
  assert(html.includes('<table class="staff-table">'));
  assert(html.includes('<div class="data-cards">'));
  const [card] = cards(html);
  assert(card, 'mobile table hiding must leave a rendered teacher card');
  for (const value of ['合成教师','合成学院','传达与媒体设计系','传达与媒体设计','13800000001','3690001','合成联系表.pdf','2026-05','启用'])
    assert(card.includes(value), value+' must be available on mobile');
  assert(card.includes('data-action="staff-edit" data-id="teacher-one"'));
  assert(card.includes('data-action="staff-status" data-id="teacher-one"'));
  assert(/\.data-cards\s*{\s*display:\s*block/.test(sharedCss), 'cards use the existing mobile visibility rule');
});

test('a mobile card opens the matching editable record and escapes teacher data',async()=>{
  const row = {...record(), name:'教师 <img src=x onerror=alert(1)>', department:null, teaching_unit:'', group_name:'办公室'};
  const page = await setup(()=>listing([row]));
  const [card] = cards(page.node('adminContent').innerHTML);
  assert(card.includes('教师 &lt;img'));
  assert(!card.includes('<img'));
  assert(card.includes('办公室'));
  assert(card.includes('待补充'));
  await page.click('staff-edit',{id:row.id});
  assert.equal(page.node('appDialog').open,true);
  const dialog = page.node('dialogContent').innerHTML;
  assert(dialog.includes('id="staff-name" value="教师 &lt;img'));
  assert(dialog.includes('id="staff-mobile" value="13800000001"'));
  assert(dialog.includes('id="staff-teaching_unit"'));
});

test('pagination replaces mobile cards with the requested page',async()=>{
  const page = await setup(path=>{
    const requested=Number(new URLSearchParams(path.split('?')[1]).get('page'));
    return listing([record('teacher-'+requested,'第'+requested+'页教师')],requested,79);
  });
  await page.click('staff-next');
  const html = page.node('adminContent').innerHTML;
  const [card] = cards(html);
  assert(card.includes('第2页教师'));
  assert(!card.includes('第1页教师'));
  assert(html.includes('第 2 / 4 页 · 共 79 条'));
  assert(page.calls[1].path.includes('page=2'));
});

test('PDF review cards remain visible and keep the original row index after paging',async()=>{
  const rows=Array.from({length:21},(_,index)=>({
    index,fields:record('draft-'+index,'待核对教师'+index),action:'new',candidates:[],source_text:'合成原文',
  }));
  const draft={id:'import-one',filename:'合成联系表.pdf',source_date:'2026-05',rows,warnings:[],
    counts:{new:21,update:0,skip:0,ambiguous:0},is_stale:false};
  const page=await setup(path=>path.startsWith('/teachers/import/')?draft:listing([]));
  await page.click('staff-resume',{id:draft.id});
  assert.equal(cards(page.node('dialogContent').innerHTML).length,20);
  await page.click('staff-review-next');
  const html=page.node('dialogContent').innerHTML;
  const [card]=cards(html);
  assert.equal(cards(html).length,1);
  assert(card.includes('待核对教师20'));
  assert(card.includes('新增'));
  assert(card.includes('data-action="staff-draft-edit" data-index="20"'));
  await page.click('staff-draft-edit',{index:'20'});
  assert(page.node('dialogContent').innerHTML.includes('id="staff-name" value="待核对教师20"'));
});
