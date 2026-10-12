(function(P){
  let csrf='';
  async function request(path,body){
    if(body&&!csrf){const info=await request('/settings');csrf=info.csrf;}
    const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),20000);
    try{
      const response=await fetch('api'+path,{method:body?'POST':'GET',headers:body?{'Content-Type':'application/json','X-Admin-CSRF':csrf}:{Accept:'application/json'},body:body?JSON.stringify(body):undefined,credentials:'same-origin',cache:'no-store',signal:controller.signal});
      const data=await response.json();if(data.csrf)csrf=data.csrf;
      if(!response.ok){if(response.status===401){location.href='login';throw new Error('请重新登录。');}if(response.status===403)csrf='';throw new Error(data.message||data.detail||'服务器返回错误。');}
      return data;
    }catch(e){if(e.name==='AbortError')throw new Error('请求超时，请刷新页面查看任务状态。');throw e;}
    finally{clearTimeout(timer);}
  }
  P.api={request};
})(window.TeachingAdmin);
