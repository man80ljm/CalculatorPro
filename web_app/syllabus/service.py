"""大纲建课：检查文件、预过滤、一次模型调用、程序规则。不写库。"""
from __future__ import annotations

import time
import io
import json

from docx import Document

from web_app.resource_limits import computation_slot, validate_document
from web_app.teaching_calendar.source import (apply_calendar_ai, calendar_brief, calendar_prompt,
                                               capture_calendar_data, retain_rule_calendar)
from web_app.teaching_calendar.service import MAX_FILE_BYTES

from web_app.deepseek_pool import PoolError, QueueTimeout
from web_app.grade_register import markdown_for_ai, parse_register
from web_app.service import ServiceError
from web_app.syllabus.client import AICallError, api_key, complete
from web_app.syllabus.extract_text import (
    DOC_MESSAGE,
    DocumentError,
    SCAN_MESSAGE,
    TYPE_MESSAGE,
    docx_to_markdown,
    pdf_to_markdown,
    suffix_of,
)
from web_app.syllabus.prefilter import prefilter, still_has_roster
from web_app.syllabus.rules import build_draft, load_prompt, parse_model_json

AI_MISSING = "尚未配置可用 AI API，请检查后台设置或服务器 DEEPSEEK_API_KEY。手填新建课程、计算和导出仍可用。"
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


def _syllabus_bundle(filename: str, data: bytes):
    if suffix_of(filename) != ".docx":
        return _syllabus_text(filename, data), {"version": 1, "source_name": filename,
            "error": "教学日历需要包含课次、学时表格的 .docx 大纲，请用 Word 版更新课程。"}
    validate_document(data, filename)
    with computation_slot():
        try:
            document = Document(io.BytesIO(data))
        except Exception as exc:
            raise DocumentError("无法读取这份 Word 文件，请确认它是 .docx。") from exc
        calendar = (capture_calendar_data(document, filename) if len(data) <= MAX_FILE_BYTES else
            {"version": 1, "source_name": filename, "error": "教学日历的 Word 大纲不能超过 5MB。"})
    text = docx_to_markdown(data, document=document)
    return text, calendar


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
    syllabus_full, calendar = _syllabus_bundle(syllabus_name, syllabus_bytes)
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
    system = load_prompt()
    calendar_requested = not calendar.get("error")
    if calendar_requested:
        brief = json.dumps(calendar_brief(calendar), ensure_ascii=False)
        if still_has_roster(brief):
            calendar_requested = False
            calendar = retain_rule_calendar(calendar, "日历原文含学生名单信息，未发送给 AI，已按规则保存。")
        else:
            system += calendar_prompt()
            user += "\n\n附加教学日历原文（仅供 teaching_calendar 整理）：\n" + brief
    messages = [
        {"role": "system", "content": system},
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
        except AICallError as exc:
            raise ServiceError(f"读取大纲失败：{exc}", status=exc.status) from None
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
    if calendar_requested:
        calendar = apply_calendar_ai(calendar, parsed_ai.get("teaching_calendar"))
    meta["elapsed_seconds"] = round(time.perf_counter() - started, 3)
    draft["meta"] = meta
    draft["sent_excerpt_has_roster"] = False
    draft["teaching_calendar"] = calendar
    return draft
