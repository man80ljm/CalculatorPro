(function(P){
  const {esc,button}=P.ui;
  let data={documents:[],records:[],drafts:[]},loaded=false,file=null,query='',termFilter='',statusFilter='',activeJob='',draft=null,job=null,correction=null,timer=null,epoch=0,busy=false;
  const clone=value=>JSON.parse(JSON.stringify(value));
  function current(){return document.getElementById('pageTitle').textContent==='课表管理';}
  function issues(row){return row.issues||[];}
  function weeksText(values){
    const weeks=[...new Set(values||[])].sort((a,b)=>a-b);if(!weeks.length)return '周次待确认';
    let start=weeks[0],end=start,ranges=[];
    for(let i=1;i<=weeks.length;i++){if(weeks[i]===end+1){end=weeks[i];continue;}ranges.push(start===end?String(start):start+'—'+end);start=end=weeks[i];}
    return '第 '+ranges.join('、')+' 周';
  }
  function documentFor(row){return data.documents.find(d=>d.id===row.documentId);}
  function names(values){return (values||[]).join('、')||'待确认';}
  function shortNames(values,limit,unit){return !values?.length?'待确认':values.slice(0,limit).join('、')+(values.length>limit?' 等 '+values.length+' '+unit:'');}
  function rows(values,review=false){
    return values.map((row,index)=>'<tr><td><b>'+esc(row.courseName)+'</b>'+(row.credits?'<span class="subtext">'+row.credits+' 学分</span>':'')+'</td><td class="teacher-names" title="'+esc(names(row.teachers))+'">'+esc(shortNames(row.teachers,3,'位教师'))+'</td><td title="'+esc(names(row.classes))+'">'+esc(shortNames(row.classes,2,'个班级'))+'</td><td>'+esc(weeksText(row.weeks))+(row.timePattern?'<span class="subtext">'+esc(row.timePattern)+'</span>':'')+(row.room?'<span class="subtext">教室：'+esc(row.room)+'</span>':'')+'</td><td><span class="pill '+(issues(row).length?'warn':'')+'">'+(issues(row).length?'待核对':'信息齐全')+'</span>'+(issues(row).length?'<span class="subtext">'+esc(issues(row)[0])+'</span>':'')+'</td><td><button class="text-button" data-action="'+(review?'pdf-correct-draft':'pdf-correct-saved')+'" data-id="'+(review?index:row.id)+'">核对 / 修正</button></td></tr>').join('');
  }
  function table(values,review=false){
    const cards=values.map((row,index)=>'<article class="data-card"><h3 class="data-card-title">'+esc(row.courseName)+'</h3><dl class="data-fields"><div class="data-field"><dt>任课教师</dt><dd>'+esc(names(row.teachers))+'</dd></div><div class="data-field"><dt>教学班级</dt><dd>'+esc(names(row.classes))+'</dd></div><div class="data-field"><dt>授课安排</dt><dd>'+esc(weeksText(row.weeks))+(row.timePattern?'<span class="subtext">'+esc(row.timePattern)+'</span>':'')+'</dd></div><div class="data-field"><dt>核对状态</dt><dd><span class="pill '+(issues(row).length?'warn':'')+'">'+(issues(row).length?'待核对':'信息齐全')+'</span></dd></div></dl><button class="text-button" data-action="'+(review?'pdf-correct-draft':'pdf-correct-saved')+'" data-id="'+esc(review?index:row.id)+'">核对 / 修正</button></article>').join('');
    return '<div class="table-wrap"><table class="pdf-records"><thead><tr><th>课程名称</th><th>任课教师</th><th>教学班级</th><th>授课安排</th><th>核对状态</th><th>操作</th></tr></thead><tbody>'+rows(values,review)+'</tbody></table><div class="data-cards">'+cards+'</div></div>';
  }
  function html(){
    const terms=[...new Map(data.documents.map(d=>[d.term.id,d.term])).values()];
    const last=data.documents[0];
    const pending=data.drafts.filter(d=>['ready','extracting','recognizing'].includes(d.status));
    return '<section class="pdf-upload panel"><div><span class="pill neutral">PDF 课表</span><h2>上传课表，自动整理开课信息</h2><p>识别学年学期、课程、教师、班级和教学周次，核对后保存。</p></div><div class="pdf-upload-actions"><input type="file" id="pdfScheduleFile" accept=".pdf,application/pdf" aria-label="上传 PDF 课表"><div class="pdf-file-name" id="pdfFileName">'+(file?esc(file.name)+' · '+Math.ceil(file.size/1024)+' KB':'选择学校公布的 PDF 课表')+'</div>'+button('开始识别','pdf-recognize',!file||busy?'disabled':'')+'</div></section>'+
      pending.map(d=>'<div class="pdf-pending"><div><b>'+esc(d.filename)+'</b><span class="subtext">'+(d.status==='ready'?'识别完成，等待核对保存':'正在识别，可继续查看其他页面')+'</span></div>'+button(d.status==='ready'?'查看识别结果':'查看进度','pdf-resume','data-id="'+d.id+'"',true)+'</div>').join('')+
      '<div class="pdf-list-head"><div><h2>已保存的课表</h2><p class="hint">'+data.documents.length+' 份 PDF · '+data.records.length+' 条开课记录'+(last?' · 最近保存：'+esc(last.filename):'')+'</p></div></div>'+
      (data.records.length?'<div class="panel"><div class="filter-bar pdf-filters"><input id="pdfQuery" aria-label="查询课表" placeholder="搜索课程、教师或班级" value="'+esc(query)+'"><select id="pdfStatus" aria-label="筛选核对状态">'+P.ui.options([{id:'',label:'全部状态'},{id:'complete',label:'信息齐全'},{id:'pending',label:'待核对'}],statusFilter)+'</select><select id="pdfTerm" aria-label="查询学期"><option value="">全部学期</option>'+P.ui.options(terms,termFilter)+'</select></div><div id="pdfRecordTable"></div><div id="pdfRecordCount" class="panel-foot"></div></div>':'<div class="pdf-empty"><span>▤</span><h3>还没有入库的 PDF 课表</h3><p>上传一次，全学期的开课信息统一保存。</p></div>')+
      '<p class="hint pdf-scope">已保存的课表集中保管，可按学期、课程、教师或班级查询。</p>';
  }
  function renderRows(){
    if(!document.getElementById('pdfRecordTable'))return;
    const normalized=P.timetable.normalized;
    const values=data.records.filter(r=>(!termFilter||r.termId===termFilter)&&(!statusFilter||(statusFilter==='pending'?issues(r).length>0:issues(r).length===0))&&(!query||normalized([r.courseName,...r.teachers,...r.classes].join(' ')).includes(normalized(query))));
    document.getElementById('pdfRecordTable').innerHTML=values.length?table(values):'<div class="empty-state">没有符合条件的课程。</div>';
    document.getElementById('pdfRecordCount').textContent='显示 '+values.length+' 条 · '+values.filter(r=>issues(r).length).length+' 条待核对';
  }
  async function render(){
    const revision=++epoch,root=document.getElementById('adminContent');
    if(loaded){root.innerHTML=html();renderRows();}else root.innerHTML='<div class="notice-box">正在读取课表库……</div>';
    try{const snapshot=await P.api.request('/schedule');data=snapshot;loaded=true;if(revision===epoch&&current()){root.innerHTML=html();renderRows();}}
    catch(e){if(revision===epoch&&current())root.innerHTML='<div class="notice-box warn">'+esc(e.message)+'</div>'+button('重新读取课表库','pdf-reload');}
  }
  function progress(status){
    P.ui.show({title:'正在识别 PDF 课表',subtitle:job?.filename||file?.name||'',body:'<div class="pdf-progress" id="pdfProgress"><div class="pdf-spinner"></div><h3>'+(status==='extracting'?'读取 PDF 表格和合并单元格':'正在提取课程、班级、时间和老师')+'</h3>'+(job?.result?.progress?'<p>已处理 '+job.result.progress.done+' / '+job.result.progress.total+' 部分 · 已提取 '+job.result.progress.records+' 条记录</p>':'<p>系统自动处理整份 PDF，完成后核对保存。</p>')+'<ol class="pdf-stages"><li class="done">上传 PDF</li><li class="'+(status==='recognizing'?'active':'')+'">API 识别</li><li>核对入库</li></ol></div>',foot:'<span>关闭弹窗后，识别任务仍会继续。</span>'+button('关闭','close-dialog','',true)});
  }
  function showFailure(message){
    P.ui.show({title:'课表识别未完成',body:'<div class="notice-box warn">'+esc(message)+'</div><p class="hint">已完成的提取会保留，重新识别时自动复用。完整提取后再保存到课表库。</p>',foot:'<span>已保存的课表继续保留。</span><div class="row-actions">'+button('关闭','close-dialog','',true)+(file?button('重新识别','pdf-recognize'):button('关闭后重新选择 PDF','close-dialog'))+'</div>'});
  }
  async function poll(identifier,show=true){
    clearTimeout(timer);activeJob=identifier;
    try{
      const value=await P.api.request('/schedule/jobs/'+identifier);job=value;
      if(value.status==='ready'){
        busy=false;draft=clone(value.result);
        if(show||document.getElementById('pdfProgress'))review();else{P.ui.toast('PDF 识别完成，可在课表管理查看结果。');if(current())render();}
        return;
      }
      if(value.status==='error'){busy=false;activeJob='';if(show||document.getElementById('pdfProgress'))showFailure(value.message);else P.ui.toast(value.message);return;}
      if(value.status==='saved'){busy=false;P.ui.close();if(current())render();return;}
      busy=true;
      if(show||document.getElementById('pdfProgress'))progress(value.status);
      timer=setTimeout(()=>poll(identifier,false),2000);
    }catch(e){busy=false;if(show||document.getElementById('pdfProgress'))showFailure(e.message);else P.ui.toast(e.message);}
  }
  function encodeFile(value){return new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result).split(',')[1]);reader.onerror=()=>reject(new Error('PDF 文件读取失败。'));reader.readAsDataURL(value);});}
  async function start(){
    if(!file||busy)return;
    busy=true;job={filename:file.name};progress('extracting');
    try{
      const result=await P.api.request('/schedule/analyze',{filename:file.name,pdf:await encodeFile(file)});
      if(result.duplicate){busy=false;P.ui.close();P.ui.toast('这份 PDF 已保存，已打开课表库；没有重复调用 API。');render();return;}
      activeJob=result.jobId;await poll(activeJob,Boolean(document.getElementById('pdfProgress')));
    }catch(e){busy=false;showFailure(e.message);}
  }
  function review(){
    if(!draft)return;
    const count=draft.records.filter(r=>issues(r).length).length;
    P.ui.show({title:'核对 PDF 识别结果',subtitle:job.filename,wide:true,body:'<div class="pdf-review-summary"><b>'+esc(draft.term.label)+'</b><span class="pill">'+draft.records.length+' 条开课记录</span><span class="pill '+(count?'warn':'')+'">'+count+' 条待核对</span></div><p class="hint">主要检查课程、教师、班级和周次；只有需要修正的记录才展开填写。</p>'+(draft.notes?.length?'<details class="pdf-notes"><summary>识别说明（'+draft.notes.length+' 项）</summary>'+draft.notes.map(n=>'<p>'+esc(n)+'</p>').join('')+'</details>':'')+table(draft.records,true)+'<p id="pdfReviewError" class="error-line" role="alert" hidden></p>',foot:'<span>未确定的信息会保留“待核对”标记。</span><div class="row-actions">'+button('稍后核对','close-dialog','',true)+button('保存到课表库','pdf-save')+'</div>'});
  }
  function parseWeeks(value,max){
    if(!value.trim())return [];
    const out=[];
    for(const part of value.replace(/[—–~～]/g,'-').split(/[、,，\s]+/).filter(Boolean)){
      if(!/^\d{1,2}(-\d{1,2})?$/.test(part))throw new Error('周次请填写 1-5, 8-12 这样的格式。');
      const [start,finish]=part.split('-').map(Number),end=finish||start;
      if(start<1||end<start||end>max)throw new Error('教学周次应在 1—'+max+' 范围内。');
      for(let week=start;week<=end;week++)out.push(week);
    }
    return [...new Set(out)].sort((a,b)=>a-b);
  }
  function correct(id,saved){
    const row=saved?data.records.find(r=>r.id===id):draft?.records[Number(id)];if(!row)return;
    correction={id,saved,row:clone(row)};
    P.ui.show({title:'核对开课信息',subtitle:row.courseName,body:'<div class="pdf-evidence"><b>PDF 原文'+(row.sourcePage?' · 第 '+row.sourcePage+' 页':'')+'</b><p>'+esc(row.sourceText||'原文引用未识别，请对照 PDF 核对。')+'</p></div><label class="field">课程名称<input id="pdfCourseName" aria-label="课表课程名称" value="'+esc(row.courseName)+'" maxlength="160"></label><div class="fields" style="margin-top:16px"><label class="field">任课教师<textarea id="pdfTeachers" aria-label="课表任课教师" rows="3">'+esc(row.teachers.join('、'))+'</textarea></label><label class="field">教学班级<textarea id="pdfClasses" aria-label="课表教学班级" rows="3">'+esc(row.classes.join('、'))+'</textarea></label></div><label class="field" style="margin-top:16px">教学周次<input id="pdfWeeks" aria-label="课表教学周次" placeholder="例如：1-5, 8-12" value="'+esc(row.weeks.join(', '))+'"></label><details class="pdf-more"><summary>更多排课信息</summary><div class="fields"><label class="field">学分<input id="pdfCredits" aria-label="课表学分" type="number" min="0.5" max="30" step="0.5" value="'+(row.credits??'')+'"></label><label class="field">总学时（原文有标注时填写）<input id="pdfHours" aria-label="课表总学时" type="number" min="1" max="500" value="'+(row.totalHours??'')+'"></label></div><label class="field" style="margin-top:14px">授课安排<textarea id="pdfTimePattern" aria-label="课表授课安排" rows="3">'+esc(row.timePattern)+'</textarea></label></details><p id="pdfCorrectionError" class="error-line" role="alert" hidden></p>',foot:'<span>留空表示尚未确定，不自动推算。</span><div class="row-actions">'+button('返回','pdf-correction-back','',true)+button('保存修正','pdf-correction-save')+'</div>'});
  }
  async function saveCorrection(){
    if(!correction)return;
    try{
      const term=correction.saved?documentFor(correction.row).term:draft.term;
      const split=value=>value.split(/[、，,；;\n]+/).map(s=>s.trim()).filter(Boolean);
      const row={...correction.row,courseName:document.getElementById('pdfCourseName').value.trim(),teachers:split(document.getElementById('pdfTeachers').value),classes:split(document.getElementById('pdfClasses').value),weeks:parseWeeks(document.getElementById('pdfWeeks').value,term.weeks),
        credits:document.getElementById('pdfCredits').value?Number(document.getElementById('pdfCredits').value):null,
        totalHours:document.getElementById('pdfHours').value?Number(document.getElementById('pdfHours').value):null,
        timePattern:document.getElementById('pdfTimePattern').value.trim(),issues:[],reviewed:true};
      if(!row.courseName)throw new Error('课程名称不能为空。');
      if(!row.teachers.length)row.issues.push('任课教师未确定');if(!row.classes.length)row.issues.push('教学班级未确定');if(!row.weeks.length)row.issues.push('教学周次未确定');
      if(correction.saved){data=await P.api.request('/schedule/record',{id:correction.id,record:row});P.ui.close();render();}
      else{draft.records[Number(correction.id)]=row;review();}
      correction=null;
    }catch(e){const el=document.getElementById('pdfCorrectionError');if(el){el.hidden=false;el.textContent=e.message;}}
  }
  async function save(){
    if(!draft||!job)return;
    const control=document.querySelector('[data-action="pdf-save"]');if(control.disabled)return;control.disabled=true;
    try{
      const result=clone(draft);
      result.records.forEach(row=>{row.outlineId='';});
      data=await P.api.request('/schedule/save',{jobId:job.id,result});loaded=true;draft=null;activeJob='';P.ui.close();render();P.ui.toast('已保存到课表库，待核对记录也已保留标记。');
    }catch(e){const el=document.getElementById('pdfReviewError');if(el){el.hidden=false;el.textContent=e.message;}control.disabled=false;}
  }
  document.addEventListener('change',e=>{
    if(e.target.id==='pdfScheduleFile'){
      const candidate=e.target.files[0];if(!candidate)return;
      if(!/\.pdf$/i.test(candidate.name)||candidate.size>8*1024*1024){P.ui.toast('请选择 8 MB 以内的 PDF。');e.target.value='';return;}
      file=candidate;document.getElementById('pdfFileName').textContent=file.name+' · '+Math.ceil(file.size/1024)+' KB';document.querySelector('[data-action="pdf-recognize"]').disabled=busy;
    }
    if(e.target.id==='pdfStatus'){statusFilter=e.target.value;renderRows();}
    if(e.target.id==='pdfTerm'){termFilter=e.target.value;renderRows();}
  });
  document.addEventListener('input',e=>{if(e.target.id==='pdfQuery'){query=e.target.value;renderRows();}});
  document.addEventListener('click',e=>{
    const el=e.target.closest('[data-action]');if(!el)return;
    switch(el.dataset.action){
      case 'pdf-recognize':start();break;case 'pdf-reload':render();break;
      case 'pdf-resume':poll(el.dataset.id,true);break;
      case 'pdf-save':save();break;
      case 'pdf-correct-draft':correct(el.dataset.id,false);break;
      case 'pdf-correct-saved':correct(el.dataset.id,true);break;
      case 'pdf-correction-save':saveCorrection();break;
      case 'pdf-correction-back':correction?.saved?P.ui.close():review();correction=null;break;
    }
  });
  P.pdfSchedule={render};
})(window.TeachingAdmin);

window.TeachingAdmin.pdfSchedule.render();
