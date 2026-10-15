"""只读取服务器环境密钥，复用原有全站 API 池；不读取后台独立密钥。"""
import json
import os
import threading
import urllib.error
import urllib.request
from contextlib import contextmanager
from web_app.deepseek_pool import primary_key, PoolError
from web_app.teaching.provider import ProviderOpener
from .errors import LessonPlanError

_local = threading.local()

@contextmanager
def observe_calls(callback):
    previous = getattr(_local, "callback", None)
    _local.callback = callback
    try: yield
    finally: _local.callback = previous


def api_key():
    try: return primary_key()
    except PoolError:
        raise LessonPlanError("服务器 API 配置无法读取，请联系管理员。",503) from None


def complete_json(system, user, *, temperature=0.5):
    key = api_key()
    if not key: raise LessonPlanError("尚未配置服务器 DeepSeek API，请联系管理员。",503,code="ai_not_configured")
    callback = getattr(_local,"callback",None)
    if callback: callback()
    config = {"source":"project", "key":key, "model":os.environ.get("LESSON_PLAN_MODEL") or os.environ.get("DEEPSEEK_MODEL") or "deepseek-flash"}
    payload = {"model":config["model"], "messages":[{"role":"system","content":system},{"role":"user","content":user}],
               "temperature":temperature, "response_format":{"type":"json_object"}, "max_tokens":4000, "stream":False,
               "thinking":{"type":"disabled"}}
    request = urllib.request.Request("https://api.deepseek.com/chat/completions",data=json.dumps(payload,ensure_ascii=False).encode(),
                                    headers={"Authorization":"Bearer "+key,"Content-Type":"application/json"},method="POST")
    try:
        with ProviderOpener(config,wait_timeout=180).open(request,timeout=180) as response:
            data=response.read(2*1024*1024+1)
        if len(data)>2*1024*1024: raise ValueError()
        body=json.loads(data); choice=body["choices"][0]
        if choice.get("finish_reason")=="length":
            raise LessonPlanError("API 返回的教案内容被截断，请重新生成。",502,code="ai_truncated")
        value=json.loads(choice["message"]["content"])
        if not isinstance(value,dict): raise ValueError()
        if system.startswith("你是有经验的高校教师"):
            for field in ("教学目标","基本内容","教学重点","教学难点","教学手段与方法"):
                lines=value.get(field)
                if not isinstance(lines,list) or len(lines)>200 or any(not isinstance(line,str) or len(line)>6000 for line in lines): raise ValueError()
            if any(not isinstance(value.get(field),str) for field in ("思考题","讨论题","作业")): raise ValueError()
        elif any(not isinstance(item,str) for item in value.values()): raise ValueError()
        return value
    except urllib.error.HTTPError as exc:
        message={401:"API 密钥认证失败",402:"API 余额不足",403:"API 无权访问",429:"API 请求限流"}.get(exc.code,"API 服务暂时不可用")
        exc.close(); raise LessonPlanError(message+"，请稍后重试或联系管理员。",502,code="ai_failed") from None
    except LessonPlanError: raise
    except Exception:
        raise LessonPlanError("API 调用超时或结果无法读取，请重试。",502,code="ai_failed") from None
