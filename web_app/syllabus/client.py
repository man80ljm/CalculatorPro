"""建课 AI：优先使用后台保存的配置，否则沿用服务器密钥池。"""
from __future__ import annotations

import json
import time
import urllib.error
import urllib.request

from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.orm import Session

from web_app.db import AdminApiConfiguration, get_engine
from web_app.deepseek_pool import PoolError, chat_completions, configured, primary_key
from web_app.teaching.api_settings import ApiError, ApiSettings, BASE_URL
from web_app.teaching.provider import ProviderOpener


class AICallError(RuntimeError):
    def __init__(self, message, status=502):
        super().__init__(message)
        self.status = status


def _saved_config():
    engine = get_engine()
    with Session(engine) as db:
        saved = db.get(AdminApiConfiguration, 1) is not None
    if not saved:
        return None
    return ApiSettings(engine).resolve({})


def api_key() -> str:
    try:
        saved = _saved_config()
        return saved["key"] if saved is not None else primary_key()
    except (PoolError, ApiError, SQLAlchemyError):
        # 让正式调用给出脱敏后的配置/排队错误，而不是误报未配置环境变量。
        return "1"


def complete(messages: list[dict], max_tokens: int = 8000, timeout: int = 120) -> dict:
    try:
        config = _saved_config()
    except (ApiError, SQLAlchemyError):
        raise AICallError("后台 API 配置无法读取，请检查后台设置。", 503) from None
    if config is None:
        try:
            ready = configured()
        except PoolError:
            ready = True
        if not ready:
            raise AICallError("尚未配置可用 API，请检查后台设置或服务器配置。", 503)
        return chat_completions("syllabus", messages, max_tokens=max_tokens, timeout=timeout, temperature=0.1)

    payload = {"model": config["model"], "messages": messages, "max_tokens": int(max_tokens),
               "temperature": 0.1, "stream": False, "response_format": {"type": "json_object"}}
    request = urllib.request.Request(BASE_URL + "/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Authorization": "Bearer " + config["key"], "Content-Type": "application/json"}, method="POST")
    started = time.perf_counter()
    try:
        # 复用后台的官方地址限制、禁止重定向、全站并发和项目密钥故障转移。
        with ProviderOpener(config, wait_timeout=min(10, timeout)).open(request, timeout=timeout) as response:
            raw = response.read(2 * 1024 * 1024 + 1)
        if len(raw) > 2 * 1024 * 1024:
            raise AICallError("API 返回内容过大，请稍后重试。")
        body = json.loads(raw)
        choice = body["choices"][0]
        if choice.get("finish_reason") == "length":
            raise AICallError("API 返回内容不完整，请稍后重试。")
        content = choice["message"]["content"]
        if not isinstance(content, str) or not content.strip():
            raise AICallError("API 未返回有效内容，请稍后重试。")
        usage = body.get("usage") or {}
        if not isinstance(usage, dict):
            usage = {}
        usage = {key: value for key, value in usage.items() if key in {
            "prompt_tokens", "completion_tokens", "total_tokens"} and isinstance(value, int)}
    except urllib.error.HTTPError as exc:
        message = {401: "API 密钥认证失败，请检查后台设置。", 402: "API 余额不足，请检查服务商账户。",
                   403: "当前 API 配置无权使用此接口。", 429: "API 请求限流，请稍后重试。"}.get(
                       exc.code, "API 服务暂时不可用，请稍后重试。")
        exc.close()
        raise AICallError(message) from None
    except ApiError:
        raise AICallError("API 请求排队或配置检查失败，请稍后重试并检查后台设置。", 503) from None
    except AICallError:
        raise
    except (OSError, ValueError, KeyError, TypeError, IndexError):
        raise AICallError("API 调用失败、超时或返回内容无法读取，请稍后重试。") from None
    from core_app.ai_report import log_deepseek_call
    elapsed = round(time.perf_counter() - started, 3)
    log_deepseek_call("syllabus", config["model"], usage, elapsed, True)
    return {"content": content, "model": config["model"], "elapsed_seconds": elapsed, **usage}
