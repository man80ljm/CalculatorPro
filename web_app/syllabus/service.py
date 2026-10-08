"""大纲建课：检查文件、预过滤、一次模型调用、程序规则。不写库。"""
from __future__ import annotations

import time

from web_app.deepseek_pool import PoolError, QueueTimeout
from web_app.grade_register import markdown_for_ai, parse_register
from web_app.service import ServiceError
from web_app.syllabus.client import AICallError, api_key, complete
from web_app.syllabus.extract_text import (
    DOC_MESSAGE,
    SCAN_MESSAGE,
    TYPE_MESSAGE,
    docx_to_markdown,
    pdf_to_markdown,
    suffix_of,
)
from web_app.syllabus.prefilter import prefilter, still_has_roster
from web_app.syllabus.rules import build_draft, load_prompt, parse_model_json

AI_MISSING = "服务器未配置 DEEPSEEK_API_KEY，无法读取大纲。手填新建课程、计算和导出仍可用。"
NETWORK = "读取大纲失败：网络问题，草稿未保存。可以重试或改为手填。"
UNREADABLE = "读取大纲失败：结果无法读取，草稿未保存。可以重试或改为手填。"


def _syllabus_text(filename: str, data: bytes) -> str:
    suffix = suffix_of(filename)
    if suffix == ".doc":
        raise ServiceError(DOC_MESSAGE)
    if suffix == ".docx":
        return docx_to_markdown(data)
    if suffix == ".pdf":
        text = pdf_to_markdown(data)
        if len("".join(text.split())) < 20:
            raise ServiceError(SCAN_MESSAGE)
        return text
    raise ServiceError(TYPE_MESSAGE)


def _register_bundle(filename: str, data: bytes) -> tuple[dict | None, str]:
    if not data:
        return None, ""
    suffix = suffix_of(filename)
    if suffix == ".doc":
        raise ServiceError(DOC_MESSAGE)
    if suffix not in {".pdf", ".xlsx", ".xls", ".docx"}:
        raise ServiceError("成绩登记表支持 Excel（.xlsx、.xls）、Word（.docx）或带文字的 PDF。")
    parsed = parse_register(filename, data)
    safe = markdown_for_ai(parsed)
    if still_has_roster(safe):
        raise ServiceError("成绩登记表折叠后仍含学生名单，已停止读取。")
    return parsed, safe


def extract_draft(syllabus_name: str, syllabus_bytes: bytes, register_name: str | None = None, register_bytes: bytes | None = None) -> dict:
    if not api_key():
        raise ServiceError(AI_MISSING)
    syllabus_full = _syllabus_text(syllabus_name, syllabus_bytes)
    syllabus = prefilter(syllabus_full)
    if still_has_roster(syllabus):
        raise ServiceError("大纲片段里仍有学生学号，已停止读取。")
    parsed = None
    register_text = ""
    if register_bytes:
        parsed, register_text = _register_bundle(register_name or "register", register_bytes)
    user = "文档A（教学大纲）：\n" + syllabus
    if register_text:
        user += "\n\n文档B（成绩登记表，名单已折叠）：\n" + register_text
    messages = [
        {"role": "system", "content": load_prompt()},
        {"role": "user", "content": user},
    ]
    started = time.perf_counter()
    last_error = ""
    meta = {"calls": 0, "model": "", "elapsed_seconds": 0, "prompt_tokens": 0, "completion_tokens": 0}
    parsed_ai = None
    for _attempt in range(2):
        try:
            result = complete(messages, max_tokens=12000)
        except QueueTimeout as exc:
            raise ServiceError(str(exc), status=503) from None
        except AICallError:
            raise ServiceError(AI_MISSING) from None
        except PoolError as exc:
            raise ServiceError(f"读取大纲失败：{exc}") from None
        except Exception:
            last_error = NETWORK
            meta["calls"] += 1
            continue
        meta["calls"] += 1
        meta["model"] = result.get("model") or ""
        meta["prompt_tokens"] = (meta["prompt_tokens"] or 0) + int(result.get("prompt_tokens") or 0)
        meta["completion_tokens"] = (meta["completion_tokens"] or 0) + int(result.get("completion_tokens") or 0)
        try:
            parsed_ai = parse_model_json(result.get("content") or "")
            break
        except (ValueError, TypeError, KeyError):
            last_error = UNREADABLE
            parsed_ai = None
    if parsed_ai is None:
        raise ServiceError(last_error or UNREADABLE)
    draft = build_draft(parsed_ai, syllabus, register_text, parsed, matrix_text=syllabus_full)
    meta["elapsed_seconds"] = round(time.perf_counter() - started, 3)
    draft["meta"] = meta
    draft["sent_excerpt_has_roster"] = False
    return draft
