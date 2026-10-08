"""建课读取用的一次 DeepSeek 调用。单元测试应替换 complete。"""
from __future__ import annotations

from web_app.deepseek_pool import PoolError, chat_completions, configured, primary_key


class AICallError(RuntimeError):
    pass


def api_key() -> str:
    try:
        return primary_key()
    except PoolError:
        return "1"


def complete(messages: list[dict], max_tokens: int = 8000, timeout: int = 120) -> dict:
    try:
        ready = configured()
    except PoolError:
        ready = True
    if not ready:
        raise AICallError("未配置密钥")
    return chat_completions("syllabus", messages, max_tokens=max_tokens, timeout=timeout, temperature=0.1)
