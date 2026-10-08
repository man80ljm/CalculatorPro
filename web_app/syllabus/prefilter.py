"""保留建课需要的章节，并把含学号/姓名的表折成人数和班级。"""
from __future__ import annotations

import re

KEEP_SECTION = ["基本信息", "课程简介", "课程描述", "课程概述", "课程目标", "毕业要求", "考核", "成绩"]
DROP_SECTION = ["教学内容", "学时分配", "教学进度", "评分标准", "参考", "教材", "网络资源", "教学方法", "实验项目", "思政"]
SUBHEADING = re.compile(r"^[（(][一二三四五六七八九十]+[）)]")
HEADING = re.compile(r"^([一二三四五六七八九十]+、|第[一二三四五六七八九十]+[章部分])")
DROP_TABLE_HEADER = ["课次", "章目", "优秀", "评分标准"]
ROSTER_HEADER = ["学号", "姓名"]
_LONG_ID = re.compile(r"\d{8,}")


def split_blocks(markdown: str):
    lines = markdown.split("\n")
    index = 0
    while index < len(lines):
        if re.match(r"^\[表格[^\]/]*\]$", lines[index]):
            end = index
            while end < len(lines) and not lines[end].startswith("[/表格"):
                end += 1
            if end >= len(lines):
                yield "text", lines[index]
                index += 1
                continue
            yield "table", lines[index : end + 1]
            index = end + 1
        else:
            yield "text", lines[index]
            index += 1


def summarize_roster(table: list[str]) -> list[str]:
    rows = [[cell.strip() for cell in line.strip("|").split("|")] for line in table[1:-1]]
    if not rows:
        return table[:1] + ["（学生名单已省略：共 0 行学生记录）"] + table[-1:]
    head, body = rows[0], rows[1:]
    class_index = head.index("班级") if "班级" in head else None
    classes = sorted({row[class_index] for row in body if class_index is not None and len(row) > class_index and row[class_index]})
    summary = f"（学生名单已省略：共 {len(body)} 行学生记录"
    if classes:
        summary += f"；班级取值：{'、'.join(classes)}"
    summary += "）"
    return [table[0], "| " + " | ".join(head) + " |", summary, table[-1]]


def prefilter(markdown: str) -> str:
    kept: list[str] = []
    keep = True
    parent_keep = True
    for kind, block in split_blocks(markdown or ""):
        if kind == "text":
            top = HEADING.match(block)
            sub = SUBHEADING.match(block)
            if top or sub:
                if any(token in block for token in KEEP_SECTION):
                    keep = True
                elif any(token in block for token in DROP_SECTION):
                    keep = False
                else:
                    keep = False if top else parent_keep
                if top:
                    parent_keep = keep
            if keep and not re.match(r"^(制定人|审订人|审批人|教师 ?：|教研室)", block) and not re.match(r"^\d{4}年 ?\d+月", block):
                kept.append(block)
        else:
            header = block[1] if len(block) > 1 else ""
            if any(token in header for token in ROSTER_HEADER):
                kept.extend(summarize_roster(block))
            elif keep and not any(token in header for token in DROP_TABLE_HEADER):
                kept.extend(block)
    return "\n".join(kept)


def still_has_roster(text: str) -> bool:
    """折叠失败时的最后一道闸：表头和长学号还在同一段里就不送出。"""
    if "学号" in text and "姓名" in text and _LONG_ID.search(text):
        return True
    return False
