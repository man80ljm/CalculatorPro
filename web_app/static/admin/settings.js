(function(P){
  const {esc,button,options}=P.ui;
  let csrf='',saved=null,busy=false,epoch=0;
  async function api(path,body){
    const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),20000);
    try{
      const response=await fetch('api/settings'+path,{method:body?'POST':'GET',headers:body?{'Content-Type':'application/json','X-Admin-CSRF':csrf}:{Accept:'application/json'},body:body?JSON.stringify(body):undefined,signal:controller.signal,credentials:'same-origin',cache:'no-store'});
      const result=await response.json();
      if(response.status===401){location.href='login';throw new Error('请重新登录。');}if(!response.ok)throw new Error(result.message||result.detail||'设置服务返回错误。');
      if(result.csrf)csrf=result.csrf;
      return result;
    }catch(e){if(e.name==='AbortError')throw new Error('连接测试超时，请稍后重试。');throw e;}
    finally{clearTimeout(timer);}
  }
  function time(value){return value?new Date(value).toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false}):'尚未保存';}
  function lockControls(){document.querySelectorAll('#apiSource,#apiKey,#apiModel,[data-action="api-test"],[data-action="api-save"]').forEach(el=>el.disabled=busy);}
  function panel(data){
    return '<div class="settings-grid"><div class="panel settings-panel"><div class="panel-head"><h2>AI API 配置</h2><span class="pill '+(data.configured?'':'neutral')+'">'+(data.configured?'已配置 · '+(data.source==='project'?'沿用服务器配置':'独立密钥'):'未配置')+'</span></div><div class="settings-body"><p class="hint">配置 PDF 课表识别使用的密钥和模型。密钥保存后不回显。</p><div class="fields" style="margin-top:18px"><label class="field">服务商<input value="DeepSeek（与原系统一致）" readonly></label><label class="field">接口地址<input value="'+esc(data.baseUrl)+'" readonly></label></div><label class="field" style="margin-top:16px">密钥来源<select id="apiSource" aria-label="密钥来源">'+options([{id:'custom',label:'填写独立密钥'},{id:'project',label:'沿用服务器配置'+(data.projectAvailable?'（已检测到密钥）':'（未检测到密钥）')}],data.source)+'</select></label><div id="apiKeyWrap" '+(data.source==='project'?'hidden':'')+'><label class="field" style="margin-top:16px">API 密钥（SK）<input id="apiKey" type="password" aria-label="API 密钥" autocomplete="off" spellcheck="false" placeholder="'+(data.configured&&data.source==='custom'?'已保存；留空保留原密钥':'sk-…')+'"></label></div><div id="apiProjectHint" class="notice-box" '+(data.source!=='project'?'hidden':'')+'>使用服务器已有的密钥池，密钥不会在页面显示。</div><label class="field" style="margin-top:16px">模型名称<input id="apiModel" aria-label="模型名称" value="'+esc(data.model)+'" maxlength="100"></label><div class="row-actions" style="margin-top:20px">'+button('测试连接','api-test')+button('保存配置','api-save','',true)+'</div><div id="apiConnectionResult" class="connection-result" role="status">'+(data.lastTest?'<span class="pill">认证连接已验证</span><p>'+esc(data.lastTest.message)+'</p><span class="subtext">'+esc(data.lastTest.scope)+'</span>':'<span class="muted">尚未测试当前配置。</span>')+'</div><p class="hint">测试只请求模型列表，不生成教学内容。上次保存：'+esc(time(data.savedAt))+'</p></div></div></div>';
  }
  async function render(){
    const revision=++epoch,el=document.getElementById('adminContent');el.innerHTML='<div class="notice-box">正在读取设置服务……</div>';
    try{const result=await api('');saved=result.settings;if(revision===epoch&&document.getElementById('pageTitle').textContent==='后台设置'){el.innerHTML=panel(saved);lockControls();}}
    catch(e){if(revision===epoch&&document.getElementById('pageTitle').textContent==='后台设置')el.innerHTML='<div class="notice-box warn">设置服务暂时不可用。请稍后刷新页面或重新登录管理员账号。</div><p class="error-line">'+esc(e.message)+'</p>'+button('重新读取','settings-reload');}
  }
  function draft(){return {source:document.getElementById('apiSource').value,model:document.getElementById('apiModel').value,key:document.getElementById('apiKey').value};}
  async function operate(test){
    if(busy||!saved)return;const body=draft(),revision=epoch;busy=true;lockControls();
    const buttons=[...document.querySelectorAll('[data-action="api-test"],[data-action="api-save"]')];buttons.forEach(b=>b.disabled=true);
    const result=document.getElementById('apiConnectionResult');result.className='connection-result';result.textContent=test?'正在连接 DeepSeek 官方接口……':'正在保存 API 配置……';
    try{
      const data=await api(test?'/test':'',body);
      if(test){const current=document.getElementById('apiConnectionResult');if(current&&revision===epoch){current.className='connection-result ok';current.innerHTML='<b>'+esc(data.message)+'</b><p>响应耗时：'+data.latencyMs+' ms</p><span class="subtext">'+esc(data.scope)+' 此结果对应刚刚测试的输入；修改配置后应重新测试。</span>';}}
      else{saved=data.settings;P.ui.toast('API 配置已保存；密钥不会在页面回显。');if(revision===epoch&&document.getElementById('apiConnectionResult'))await render();}
    }catch(e){const current=document.getElementById('apiConnectionResult');if(current&&revision===epoch){current.className='connection-result error';current.textContent=e.message;}else P.ui.toast(e.message);}
    finally{busy=false;lockControls();document.querySelectorAll('[data-action="api-test"],[data-action="api-save"]').forEach(b=>b.disabled=false);}
  }
  document.addEventListener('click',e=>{const el=e.target.closest('[data-action]');if(!el)return;if(el.dataset.action==='api-test')operate(true);if(el.dataset.action==='api-save')operate(false);if(el.dataset.action==='settings-reload')render();});
  document.addEventListener('change',e=>{if(e.target.id==='apiSource'){const project=e.target.value==='project';document.getElementById('apiKeyWrap').hidden=project;document.getElementById('apiProjectHint').hidden=!project;if(project)document.getElementById('apiModel').value=saved.projectModel;document.getElementById('apiKey').value='';}});
  document.addEventListener('input',e=>{if(['apiKey','apiModel','apiSource'].includes(e.target.id)){const result=document.getElementById('apiConnectionResult');result.className='connection-result';result.textContent='输入已改变，请重新测试。';}});
  P.settings={render};
})(window.TeachingAdmin);

window.TeachingAdmin.settings.render();
