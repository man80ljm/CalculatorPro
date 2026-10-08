"""老师下载资料时使用的简短名称，与服务器内部文件名无关。"""
from __future__ import annotations

import re


def archive_filename(course_name: str, year_start="", year_end="", semester="", school_year_term="") -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(course_name or "课程资料")).strip(" .") or "课程资料"
    start, end, term = (str(value or "").strip() for value in (year_start, year_end, semester))
    def valid():
        return bool(re.fullmatch(r"\d{4}", start) and re.fullmatch(r"\d{4}", end) and term in {"1", "2"})
    if not valid():
        match = re.search(r"(\d{4})\s*[-—–/]\s*(\d{4}).*?(?:第\s*)?([12])\s*(?:学期)?$", str(school_year_term or ""))
        if match:
            start, end, term = match.groups()
    if valid():
        suffix = f"{start}-{end}学年_第{term}学期"
    else:
        suffix = "未填写学期"
    return f"{name}_{suffix}.zip"


def settings_archive_filename(course_name: str, settings: dict) -> str:
    opened = settings.get("course_open_info") or {}
    basic = settings.get("course_basic_info") or {}
    return archive_filename(course_name, opened.get("year_start"), opened.get("year_end"),
                            opened.get("semester") or opened.get("term"), basic.get("school_year_term"))
