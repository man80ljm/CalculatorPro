"""从大纲表格解析课程目标与毕业要求矩阵。不调用模型。"""
from __future__ import annotations

import re

NAME_ONLY_NOTE = "大纲未给出毕业要求编号，按矩阵行标签填写"
NO_MATRIX_REASON = "大纲没有课程目标与毕业要求的矩阵"
MATRIX_SOURCE = "大纲·矩阵行标签"
REQ_DEF_SOURCE = "大纲·毕业要求定义"
IND_DEF_SOURCE = "大纲·指标点定义"
STRENGTH_SOURCE = "大纲·矩阵"

_OBJ = re.compile(r"课程目标\s*(\d+)")
_DEF_REQ = re.compile(r"毕业要求\s*(\d+)\s*[：:]\s*([^\n|]{1,80})")
_DEF_IND = re.compile(r"指标点\s*(\d+)\s*([.\-－])\s*(\d+)\s*[：:]\s*([^\n|]{1,120})")
_IND_LABEL = re.compile(r"指标点\s*(\d+)\s*([.\-－])\s*(\d+)\s*[：:．.]?\s*(.*)")
_REQ_LABEL = re.compile(r"毕业要求\s*(\d+)\s*[：:．.]?\s*(.*)")
_CODE_LABEL = re.compile(r"(\d+)\s*([.\-－])\s*(\d+)\s*[：:．.]?\s*(.*)")
_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")


def _norm(text: str) -> str:
    return " ".join((text or "").replace("\u3000", " ").translate(_DIGITS).split())


def _field(value: str = "", status: str = "需手填", source: str = "", reason: str = "", note: str = "") -> dict:
    filled = status == "已填" and str(value or "").strip() not in {"", "需手填", "手填"}
    data = {
        "value": str(value).strip() if filled else "",
        "status": "已填" if filled else "需手填",
        "source": source if filled else "",
        "reason": "" if filled else (reason or "需要老师填写"),
        "candidates": [],
    }
    if filled and note:
        data["note"] = note
    return data


def _sources(items: list[str]) -> str:
    ordered: list[str] = []
    for item in items:
        if item and item not in ordered:
            ordered.append(item)
    return "；".join(ordered)


def _split_row(line: str) -> list[str]:
    raw = line.strip()
    if raw.startswith("|"):
        raw = raw[1:]
    if raw.endswith("|"):
        raw = raw[:-1]
    return [_norm(cell) for cell in raw.split("|")]


def iter_markdown_tables(text: str) -> list[list[list[str]]]:
    lines = (text or "").splitlines()
    tables: list[list[list[str]]] = []
    index = 0
    while index < len(lines):
        stripped = lines[index].strip()
        if re.match(r"^\[表格[^\]/]*\]$", stripped):
            block: list[list[str]] = []
            index += 1
            while index < len(lines) and not lines[index].strip().startswith("[/表格"):
                if lines[index].strip().startswith("|"):
                    block.append(_split_row(lines[index]))
                index += 1
            if block:
                tables.append(_pad(block))
            continue
        if stripped.startswith("|"):
            block = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                block.append(_split_row(lines[index]))
                index += 1
            if block:
                tables.append(_pad(block))
            continue
        index += 1
    return tables


def _pad(rows: list[list[str]]) -> list[list[str]]:
    width = max(len(row) for row in rows)
    return [row + [""] * (width - len(row)) for row in rows]


def _header_kind(cell: str) -> str:
    text = cell.replace(" ", "")
    has_req = "毕业要求" in text
    has_ind = "指标点" in text
    if has_req and has_ind:
        return "combined"
    if has_ind:
        return "indicator"
    if has_req:
        return "requirement"
    return "label"


def _objective_columns(rows: list[list[str]]) -> tuple[dict[int, int], int]:
    found: dict[int, int] = {}
    header_end = 0
    scan = rows[:4]
    for row_index, row in enumerate(scan):
        for col_index, cell in enumerate(row):
            match = _OBJ.search(cell.replace(" ", ""))
            if match and col_index not in found:
                found[col_index] = int(match.group(1))
                header_end = max(header_end, row_index + 1)
    if found:
        return found, header_end
    for row_index, row in enumerate(scan[:-1]):
        if not any("课程目标" in cell.replace(" ", "") and not _OBJ.search(cell.replace(" ", "")) for cell in row):
            continue
        for col_index, cell in enumerate(scan[row_index + 1]):
            bare = re.fullmatch(r"(\d+)", cell.replace(" ", ""))
            if bare:
                found[col_index] = int(bare.group(1))
        if found:
            return found, row_index + 2
    return {}, 0


def _column_kind(rows: list[list[str]], col: int, header_end: int) -> str:
    best = "label"
    best_score = -1
    priority = {"combined": 3, "indicator": 2, "requirement": 1, "label": 0}
    for row_index in range(header_end):
        if col >= len(rows[row_index]):
            continue
        kind = _header_kind(rows[row_index][col])
        if kind == "label":
            continue
        score = row_index * 10 + priority[kind]
        if score >= best_score:
            best = kind
            best_score = score
    return best


def _strengths(cell: str) -> list[str]:
    letters = re.sub(r"[^A-Za-z]", "", cell or "")
    if not letters or set(letters.upper()) - set("HML"):
        return []
    return re.findall(r"[HML]", cell.upper())


def _definitions(text: str) -> tuple[dict[str, str], dict[str, str]]:
    requirements: dict[str, str] = {}
    indicators: dict[str, str] = {}
    for match in _DEF_REQ.finditer(text or ""):
        requirements.setdefault(match.group(1), _norm(match.group(2)))
    for match in _DEF_IND.finditer(text or ""):
        left, right = match.group(1), match.group(3)
        name = _norm(match.group(4))
        for key in (f"{left}.{right}", f"{left}-{right}", f"{left}{match.group(2)}{right}"):
            indicators.setdefault(key, name)
    return requirements, indicators


def _lookup_indicator(defs: dict[str, str], code: str) -> str:
    if code in defs:
        return defs[code]
    swapped = code.replace(".", "-") if "." in code else code.replace("-", ".").replace("－", ".")
    return defs.get(swapped, "")


def _split_label(text: str) -> dict:
    raw = _norm(text)
    parsed = {
        "requirement_code": "",
        "requirement_name": "",
        "indicator_code": "",
        "indicator_name": "",
        "plain": "",
    }
    if not raw:
        return parsed
    indicator = _IND_LABEL.fullmatch(raw)
    if indicator:
        parsed["requirement_code"] = indicator.group(1)
        parsed["indicator_code"] = f"{indicator.group(1)}{indicator.group(2)}{indicator.group(3)}"
        parsed["indicator_name"] = indicator.group(4).strip()
        return parsed
    requirement = _REQ_LABEL.fullmatch(raw)
    if requirement:
        parsed["requirement_code"] = requirement.group(1)
        rest = requirement.group(2).strip()
        nested = _IND_LABEL.match(rest) or _CODE_LABEL.match(rest)
        if nested:
            parsed["indicator_code"] = f"{nested.group(1)}{nested.group(2)}{nested.group(3)}"
            parsed["indicator_name"] = nested.group(4).strip()
            if nested.group(1) != parsed["requirement_code"]:
                parsed["requirement_code"] = nested.group(1)
        else:
            parsed["requirement_name"] = rest
        return parsed
    coded = _CODE_LABEL.fullmatch(raw)
    if coded:
        parsed["requirement_code"] = coded.group(1)
        parsed["indicator_code"] = f"{coded.group(1)}{coded.group(2)}{coded.group(3)}"
        parsed["indicator_name"] = coded.group(4).strip()
        return parsed
    parsed["plain"] = raw
    return parsed


def _format_requirement(parsed: dict, req_defs: dict[str, str]) -> tuple[str, list[str]]:
    code = parsed["requirement_code"]
    if not code:
        return "", []
    name = parsed["requirement_name"]
    sources = [MATRIX_SOURCE]
    defined = req_defs.get(code, "")
    if defined and defined not in name:
        name = f"{name}：{defined}" if name else defined
        sources.append(REQ_DEF_SOURCE)
    value = f"毕业要求{code}" + (f"：{name}" if name else "")
    return value, sources


def _format_indicator(parsed: dict, ind_defs: dict[str, str]) -> tuple[str, list[str]]:
    if parsed["indicator_code"]:
        code = parsed["indicator_code"]
        text = code + (f" {parsed['indicator_name']}" if parsed["indicator_name"] else "")
        sources = [MATRIX_SOURCE]
        defined = _lookup_indicator(ind_defs, code)
        if defined and defined not in text:
            if parsed["indicator_name"]:
                text = f"{text}：{defined}"
            else:
                text = f"{code} {defined}"
            sources.append(IND_DEF_SOURCE)
        return text, sources
    if parsed["plain"]:
        return parsed["plain"], [MATRIX_SOURCE]
    return "", []


def _cell_parts(text: str, role: str, req_defs: dict[str, str], ind_defs: dict[str, str]) -> tuple[str, list[str], str, list[str], bool]:
    """最后一项为真：这一格只是名称，毕业要求和指标点都用它。"""
    parsed = _split_label(text)
    requirement, req_sources = _format_requirement(parsed, req_defs)
    indicator, ind_sources = _format_indicator(parsed, ind_defs)
    if role == "requirement" and not requirement and text and not parsed["indicator_code"]:
        return text, [MATRIX_SOURCE], "", [], False
    if role == "indicator" and not indicator and text:
        return requirement, req_sources, text, [MATRIX_SOURCE], False
    if role == "combined" and parsed["plain"]:
        return parsed["plain"], [MATRIX_SOURCE], parsed["plain"], [MATRIX_SOURCE], True
    return requirement, req_sources, indicator, ind_sources, False


def _parse_table(rows: list[list[str]], req_defs: dict[str, str], ind_defs: dict[str, str]) -> dict[int, list[dict]] | None:
    objectives, header_end = _objective_columns(rows)
    if not objectives or header_end <= 0:
        return None
    header_blob = "".join(cell for row in rows[:header_end] for cell in row)
    if "毕业要求" not in header_blob and "指标点" not in header_blob:
        return None
    kinds = {
        col: _column_kind(rows, col, header_end)
        for col in range(len(rows[0]))
        if col not in objectives
    }
    grouped: dict[int, list[dict]] = {}
    saw_mark = False
    for row in rows[header_end:]:
        requirement = ""
        req_sources: list[str] = []
        indicator = ""
        ind_sources: list[str] = []
        label_modes: list[bool] = []
        for col, kind in kinds.items():
            if col >= len(row) or not row[col]:
                continue
            req, req_src, ind, ind_src, name_only = _cell_parts(row[col], kind if kind != "label" else "combined", req_defs, ind_defs)
            label_modes.append(name_only)
            if req and req not in requirement:
                requirement = req if not requirement else f"{requirement}；{req}"
                req_sources.extend(req_src)
            if ind and ind not in indicator:
                indicator = ind if not indicator else f"{indicator}；{ind}"
                ind_sources.extend(ind_src)
        marks: dict[int, list[str]] = {}
        for col, number in objectives.items():
            if col >= len(row):
                continue
            found = _strengths(row[col])
            if found:
                marks[number] = found
                saw_mark = True
        if not marks:
            continue
        for number, found in marks.items():
            grouped.setdefault(number, []).append(
                {
                    "requirement": requirement,
                    "requirement_sources": req_sources,
                    "indicator": indicator,
                    "indicator_sources": ind_sources,
                    "name_only": bool(label_modes) and all(label_modes),
                    "strength": "、".join(found),
                }
            )
    if not saw_mark:
        return None
    return grouped


def _merge_objective(number: int, rows: list[dict]) -> dict:
    requirements: list[str] = []
    req_sources: list[str] = []
    indicators: list[str] = []
    ind_sources: list[str] = []
    strengths: list[str] = []
    missing_requirement = False
    missing_indicator = False
    name_only = True
    for row in rows:
        name_only = name_only and bool(row.get("name_only"))
        if row["requirement"]:
            if row["requirement"] not in requirements:
                requirements.append(row["requirement"])
            req_sources.extend(row["requirement_sources"])
        else:
            missing_requirement = True
        if row["indicator"]:
            indicators.append(row["indicator"])
            ind_sources.extend(row["indicator_sources"])
        else:
            missing_indicator = True
        if row["strength"]:
            strengths.append(row["strength"])
    note = NAME_ONLY_NOTE if name_only and requirements and indicators else ""
    if requirements and not missing_requirement:
        requirement = _field("；".join(requirements), "已填", _sources(req_sources), note=note)
    elif requirements and missing_requirement:
        requirement = _field("；".join(requirements), "已填", _sources(req_sources), note=note)
        requirement["reason"] = "部分矩阵行未给出毕业要求"
    else:
        requirement = _field("", "需手填", reason="矩阵这一行只给出指标点编号，未给出毕业要求")
    if indicators and not missing_indicator:
        indicator = _field("；".join(indicators), "已填", _sources(ind_sources), note=note)
    elif indicators:
        indicator = _field("；".join(indicators), "已填", _sources(ind_sources), note=note)
        indicator["reason"] = "部分矩阵行没有指标点"
    else:
        indicator = _field("", "需手填", reason="矩阵这一行只给出毕业要求，未给出指标点")
    return {
        "objective": f"课程目标{number}",
        "requirement": requirement,
        "indicator": indicator,
        "strength": "；".join(strengths),
        "strength_source": STRENGTH_SOURCE if strengths else "",
    }


def matrix_grad_items(text: str, objective_count: int) -> list[dict] | None:
    """解析到矩阵就返回每个课程目标一项；没有这样的表则返回 None，交给原有逻辑。"""
    req_defs, ind_defs = _definitions(text or "")
    grouped: dict[int, list[dict]] | None = None
    for table in iter_markdown_tables(text or ""):
        parsed = _parse_table(table, req_defs, ind_defs)
        if parsed:
            grouped = parsed
            break
    if grouped is None:
        return None
    highest = max(grouped) if grouped else 0
    count = max(int(objective_count or 0), highest)
    items = []
    for number in range(1, count + 1):
        rows = grouped.get(number) or []
        if not rows:
            items.append(
                {
                    "objective": f"课程目标{number}",
                    "requirement": _field("", "需手填", reason=f"大纲矩阵里课程目标{number}没有标 H/M/L"),
                    "indicator": _field("", "需手填", reason=f"大纲矩阵里课程目标{number}没有标 H/M/L"),
                    "strength": "",
                    "strength_source": "",
                }
            )
            continue
        items.append(_merge_objective(number, rows))
    return items
