"""Fixed official endpoint, bounded responses, global leases and key failover."""
import urllib.request,urllib.error
from contextlib import contextmanager
from web_app.deepseek_pool import acquire,external_slot,PoolError
from .api_settings import ApiError,BASE_URL,NoRedirect

class ProviderOpener:
    def __init__(self,config,*,wait_timeout=180):
        self.config=config
        self.wait_timeout=wait_timeout
    @contextmanager
    def open(self,request,timeout):
        if request.full_url not in {BASE_URL+'/chat/completions',BASE_URL+'/models'}:raise ApiError('识别接口地址不允许。')
        if self.config['source']=='project':
            try:lease=acquire('syllabus',wait_timeout=3 if request.full_url.endswith('/models') else self.wait_timeout)
            except PoolError:raise ApiError('识别请求排队失败，请稍后重试。',503) from None
            with lease:
                with self._open(request,timeout,lease.account.keys) as response:yield response
        else:
            try:
                with external_slot(wait_timeout=3 if request.full_url.endswith('/models') else self.wait_timeout):
                    with self._open(request,timeout,[self.config['key']]) as response:yield response
            except PoolError:raise ApiError('识别请求排队超时，请稍后重试。',503) from None
    @contextmanager
    def _open(self,request,timeout,keys):
        client=urllib.request.build_opener(NoRedirect())
        headers={name:value for name,value in request.header_items() if name.lower()!='authorization'}
        response=None
        for index,key in enumerate(keys):
            try:
                candidate=urllib.request.Request(request.full_url,data=request.data,headers=dict(headers,Authorization='Bearer '+key),method=request.get_method())
                response=client.open(candidate,timeout=timeout)
                break
            except urllib.error.HTTPError as exc:
                if exc.code in {401,403,429} and index+1<len(keys):
                    exc.close();continue
                raise
        try:yield response
        finally:
            if response is not None:response.close()