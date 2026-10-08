"""从班级名推导上课专业。只按文字规则，不调用 AI。"""
from __future__ import annotations

import re

_NOISE = re.compile(r"^[\d\s级班（）()A-Za-z]+$")
_GRADE = re.compile(r"^(?:(?:19|20)\d{2}|\d{2})级?")
_TRAIL = re.compile(r"(?:[（(]\s*\d+\s*[）)]\s*班?|\d+\s*班|班|[A-Za-z])$")
_SPLIT = re.compile(r"[、,，;；\n|]+")

SOURCE = "成绩登记表·班级（推导）"


def derive_major_from_class(raw: str | None) -> str | None:
    """去掉年级前缀和班号。推不出时返回 None（空值、纯数字、只剩年级）。"""
    text = (raw or "").strip()
    if not text or _NOISE.fullmatch(text):
        return None
    text = _GRADE.sub("", text, count=1).strip()
    if not text:
        return None
    while True:
        nxt = _TRAIL.sub("", text).strip()
        if nxt == text:
            break
        text = nxt
    if not text or re.fullmatch(r"[\d\s]+", text):
        return None
    return text


def split_class_names(raw: str | None) -> list[str]:
    return [part.strip() for part in _SPLIT.split(raw or "") if part.strip()]


def resolve_majors(names: list[str]) -> dict:
    """多个班级推出同一专业名则已填；不同则需手填并给出候选。"""
    found: list[str] = []
    for name in names:
        major = derive_major_from_class(name)
        if major and major not in found:
            found.append(major)
    if len(found) == 1:
        return {
            "status": "已填",
            "value": found[0],
            "source": SOURCE,
            "reason": "",
            "candidates": [],
        }
    if len(found) > 1:
        return {
            "status": "需手填",
            "value": "",
            "source": "",
            "reason": "多个班级推出的专业名不同，请选择一个。一个班对应一条学期记录。",
            "candidates": [{"label": f"{SOURCE}：{item}", "value": item} for item in found],
        }
    return {
        "status": "需手填",
        "value": "",
        "source": "",
        "reason": "无法从班级推导上课专业",
        "candidates": [],
    }
