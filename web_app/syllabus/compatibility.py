"""Legacy syllabus adapters. Existing extraction, calculator rules and templates stay intact."""
import copy
import hashlib
import io
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

from docx import Document
from web_app.service import ServiceError
from web_app.resource_limits import computation_slot, validate_document
from .extract_text import docx_to_markdown, pdf_to_markdown
from .grad_matrix import iter_markdown_tables
from .rules import build_draft, field, validate_relation_grid
from web_app.teaching_calendar import engine as calendar_engine
from web_app.teaching_calendar.source import normalize_calendar_data

MAX_BYTES = 5 * 1024 * 1024
MAX_GOALS = 8  # The existing calculator editor supports eight objective columns.
BASIC = {"course_name": ("课程中文名称", "课程名称"), "credits": ("学分",), "hours": ("学时", "总学时"),
         "course_type": ("课程类别", "课程性质"), "course_code": ("课程编号", "课程代码", "课程编码"),
         "college": ("开课单位", "开课学院")}
GOAL = re.compile(r"^(?:课程)?目标\s*([1-9]\d*)\s*[:：、.．]?\s*(.*)", re.S)


def compact(value):
    return re.sub(r"\s+", "", str(value or ""))


def hours(value):
    raw = str(value).strip()
    if not re.fullmatch(r"\d+(?:\.\d+)?", raw):
        raise ServiceError("学时必须是非负数字。")
    number = Decimal(raw)
    if not number.is_finite() or number > 10000:
        raise ServiceError("学时数字超出范围。")
    return number


def clean_detail(value):
    # The original rule_split expects the goal before the content.
    parts = {}
    matches = list(re.finditer(r"(教学目标|教学要求|教学内容)\s*[:：]", value))
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(value)
        parts[match[1]] = value[match.end():end].strip()
    goal = parts.get("教学目标", parts.get("教学要求", ""))
    content = parts.get("教学内容", value if not matches else "")
    return f"教学目标：{goal}\n教学内容：{content}" if goal else f"教学内容：{content}"


def references(markdown, textbook):
    materials, sites = [], []
    active = False
    for line in markdown.splitlines():
        if re.match(r"^[一二三四五六七八九十0-9]+[、.．]", line) and any(k in line for k in ("参考资料", "参考文献", "参考书", "教材与参考")):
            active = True
            continue
        if not active or line.startswith(("[表格", "[/表格", "|")):
            continue
        if re.match(r"^(制定|审定|审订|批准|审批|校对|编写|负责人)", line) or re.match(r"^20\d{2}年", line):
            break
        if re.match(r"^(?:[（(][一二三四五六七八九十]+[）)]|\d+[.．、])", line) and any(k in line for k in ("书目", "文献", "资源", "资料")):
            continue
        line = re.sub(r"^\s*[\[［]?\d+[\]］]?\s*[．.、]?\s*", "", line).strip()
        if line and line != "无":
            (sites if re.search(r"https?://|www\.", line) else materials).append(line)
    if not materials and textbook:
        materials.append(textbook)
    return list(dict.fromkeys(materials))[:100], list(dict.fromkeys(sites))[:100]


def mapped_requirements(tables, count):
    direct, matrix = {}, {}
    for rows in tables:
        header = "".join(rows[0]) if rows else ""
        if "毕业要求" not in header:
            continue
        if len(rows[0]) <= 3 and "对应" in header:
            for row in rows[1:]:
                match = GOAL.match(row[0])
                if match and len(row) > 1:
                    strength = re.search(r"[(（]\s*([HML])\s*[）)]", row[1], re.I)
                    name = re.sub(r"[(（]\s*[HML]\s*[）)]", "", row[1], flags=re.I).strip()
                    if name:
                        direct[int(match[1])] = (name, strength[1].upper() if strength else "")
        elif "课程目标" in header and len(rows) > 2:
            strengths = rows[1]
            for row in rows[2:]:
                match = re.fullmatch(r"课程目标\s*(\d+)", row[0])
                if not match:
                    continue
                labels, levels = [], []
                for col, mark in enumerate(row[1:], 1):
                    if re.fullmatch(r"[√✓✔●1]|[HML]", compact(mark), re.I):
                        labels.append(rows[0][col])
                        level = mark.upper() if compact(mark).upper() in {"H", "M", "L"} else strengths[col].upper()
                        if level in {"H", "M", "L"}:
                            levels.append(level)
                if labels:
                    matrix[int(match[1])] = ("；".join(labels), min(levels, key=lambda s: "HML".index(s)) if levels else "")
    result, issues = [], []
    for number in range(1, count + 1):
        if number in direct and number in matrix and direct[number] != matrix[number]:
            issues.append(f"课程目标{number}的毕业要求对应关系前后不一致：{direct[number][0]} / {matrix[number][0]}")
        value, strength = direct.get(number, matrix.get(number, ("", "")))
        item = field(value, "已填" if value else "需手填", source="大纲·旧版对应关系表",
                     reason="此课程目标缺少毕业要求对应信息")
        if value:
            item["note"] = "原文未给出编号，按毕业要求名称填写"
        result.append({"objective": f"课程目标{number}", "requirement": copy.deepcopy(item),
                       "indicator": copy.deepcopy(item), "strength": strength})
    return result, issues


def score_components(tables):
    pattern = re.compile(r"([^+；;|。]{1,45}?)\s*[（(]\s*(\d+(?:\.\d+)?)\s*%[）)]")
    for rows in tables:
        for row in rows:
            if not any("成绩评定" in cell or "总成绩" in cell for cell in row):
                continue
            matches = []
            for name, ratio in pattern.findall(" ".join(row)):
                name = re.sub(r"^.*(?:成绩评定|总成绩)[：:]?\s*", "", name).strip()
                name = name.removesuffix("成绩").strip()
                if name:
                    matches.append({"name": name, "ratio": str(Decimal(ratio))})
            if matches:
                return matches
    return []


def relation_seed(tables, goals, components):
    count = len(goals)
    for rows in tables:
        if rows and "考核环节" in rows[0] and any(re.fullmatch(r"课程目标\s*\d+", cell) for cell in rows[0]):
            checked = validate_relation_grid(rows)
            if checked["ok"]:
                return rows, False
    if not count or not components:
        return [], False
    groups = {}
    for component in components:
        name = component["name"]
        link = "期中考核" if "期中" in name else "期末考核" if re.search(r"期末|期终|课程作业|课程设计|结课|考试", name) else "平时考核"
        groups.setdefault(link, []).append(component)
    grid = [["考核环节", "占比", "考核方式"] + [f"课程目标{i+1}" for i in range(count)] + ["小计"]]
    for link, methods in groups.items():
        ratio = sum(Decimal(item["ratio"]) for item in methods)
        if ratio <= 0:
            continue
        for index, method in enumerate(methods):
            share = Decimal(method["ratio"]) / ratio
            # Seed for the AI to refine. It is explicitly labelled as an inferred allocation.
            weights = [share / count] * count
            grid.append([link if index == 0 else "", str(ratio) + "%" if index == 0 else "", method["name"]]
                        + [f"{weight*100:.6f}%" for weight in weights] + [f"{share*100:.6f}%"])
    return grid, True


def load(filename, data):
    if len(data) > MAX_BYTES:
        raise ServiceError("教学大纲不能超过 5MB。", status=413)
    suffix = Path(filename).suffix.lower()
    if suffix not in {".docx", ".pdf"}:
        raise ServiceError("请上传 .docx 大纲或带文字的 PDF。")
    is_teaching_plan = False
    if suffix == ".docx":
        validate_document(data, filename)
        with computation_slot():
            try:
                document = Document(io.BytesIO(data))
            except Exception:
                raise ServiceError("无法读取这份 Word 大纲，请确认文件没有损坏。") from None
        from .teaching_plan import to_markdown
        converted = to_markdown(document)
        is_teaching_plan = converted is not None
        markdown = converted if is_teaching_plan else docx_to_markdown(data, document=document)
    else:
        markdown = pdf_to_markdown(data)
    tables = iter_markdown_tables(markdown)
    info = {}
    for rows in tables:
        for row in rows:
            for index, cell in enumerate(row[:-1]):
                label = compact(cell)
                if label in {name for aliases in BASIC.values() for name in aliases} | {"教材选用", "课程负责人"}:
                    info.setdefault(label, row[index+1].strip())
    basics = {key: next((info[alias] for alias in aliases if info.get(alias)), "") for key, aliases in BASIC.items()}
    if not basics["course_name"]:
        match = re.search(r"《([^》]+)》", markdown[:1000])
        basics["course_name"] = match[1] if match else Path(filename).stem
    goals_by_id = {}
    for rows in tables:
        if not rows or "课程目标" not in "".join(rows[0]) or len(rows[0]) > 3:
            continue
        if any(k in "".join(rows[0]) for k in ("考核", "权重", "学时")):
            continue
        for row in rows[1:]:
            match = GOAL.match(row[0])
            if match and match[2].strip():
                goals_by_id.setdefault(int(match[1]), match[2].strip())
    for line in markdown.splitlines():
        match = GOAL.match(line)
        if match and match[2].strip():
            goals_by_id.setdefault(int(match[1]), match[2].strip())
    referenced = [int(value) for value in re.findall(r"课程目标\s*([1-9]\d*)", markdown)]
    original_count = max([0, *goals_by_id, *referenced])
    count = min(original_count, MAX_GOALS)
    goals = [goals_by_id.get(number, "") for number in range(1, count+1)]
    paragraphs = [line for line in markdown.splitlines() if not line.startswith(("[表格", "[/表格", "|"))]
    intro = []
    active = False
    for line in paragraphs:
        if re.match(r"^[一二三四五六七八九十0-9]+[、.．].*(课程简介|课程描述|课程概述)", line):
            active = True
            continue
        if active and re.match(r"^[一二三四五六七八九十0-9]+[、.．]", line):
            break
        if active:
            intro.append(line)
    lessons = []
    for rows in tables:
        header = [compact(cell) for cell in rows[0]] if rows else []
        if "学时" not in header or not any(label in header for label in ("课次", "序号", "次数")):
            continue
        for row in rows[1:]:
            item = {label: row[i] for i, label in enumerate(header) if i < len(row)}
            number = next((item[k] for k in ("课次", "序号", "次数") if item.get(k)), "")
            if not number.isdigit():
                continue
            def pick(*tokens):
                return next((value for token in tokens for label, value in item.items() if token in label), "")
            lessons.append({"课次": number, "学时": item.get("学时", ""), "章目名称": pick("章目", "章节", "题目"),
                "教学目标及教学内容": clean_detail(pick("教学目标", "教学要求", "教学内容")),
                "教学方式": pick("教学方式", "教学方法"), "支撑课程目标": pick("支撑", "达成课程目标"),
                "课程思政": pick("课程思政点", "思政")})
        break
    textbook = info.get("教材选用", "")
    materials, sites = references(markdown, textbook)
    components = score_components(tables)
    usual = [item for item in components if not re.search(r"期末|期终|课程作业|课程设计|结课|考试", item["name"])]
    final = [item for item in components if item not in usual]
    outline = {"课程名称": basics["course_name"], "课程负责人": info.get("课程负责人", ""),
        "课程简介": "\n".join(intro), "课程目标": goals, "lessons": lessons, "学习资料": materials, "学习网站": sites,
        "考核": {"平时比例": str(sum(Decimal(item["ratio"]) for item in usual)) if usual else "",
                 "期末比例": str(sum(Decimal(item["ratio"]) for item in final)) if final else "",
                 "平时项": [item["name"] for item in usual], "期末项": [item["name"] for item in final], "原表": ""}}
    ai = {"course_basic_info": {key: field(value, "已填", source="大纲·基本信息") for key, value in basics.items()},
          "course_description": field(outline["课程简介"], "已填", source="大纲·课程简介"),
          "objectives_status": "已填", "objectives": [{"index": i+1, "text": goal} for i, goal in enumerate(goals)]}
    draft = build_draft(ai, markdown, "", None, matrix_text=markdown)
    draft["objectives"] = goals
    draft["grad_req_map"], issues = mapped_requirements(tables, count)
    if is_teaching_plan:
        issues.insert(0,"上传的是课程教案，已还原课次、教学内容与学时；缺失的课程目标、毕业要求和考核矩阵由 AI 推断，确认前请核对。")
        for key in ("course_name","hours"):
            draft["fields"][key] = field(basics[key],"已填",source="课程教案·封面" if key=="course_name" else "课程教案·课次学时合计")
    if original_count > MAX_GOALS:
        issues.append("原文目标数量或引用编号超出网页支持的 8 个目标，需要合并并修正课次引用")
    grid, inferred_grid = relation_seed(tables, goals, components)
    draft["relation"] = validate_relation_grid(grid)
    for number, goal in enumerate(goals, 1):
        if not goal:
            issues.append(f"后文引用课程目标{number}，但没有给出目标定义")
    if inferred_grid:
        issues.append("旧版考核表已转换为计算器结构，考核方式与各课程目标的分权重需要 AI 完善")
    if lessons and basics["hours"]:
        try:
            total = sum(hours(row["学时"]) for row in lessons)
            if total != hours(basics["hours"]):
                issues.append(f"课次学时合计 {total} 与总学时 {basics['hours']} 不一致")
        except ServiceError as exc:
            issues.append(str(exc))
    calendar = calendar_from_outline(filename, outline, textbook)
    draft.update(teaching_calendar=calendar, issues=[{"field": "compatibility", "problem": text} for text in issues],
                 sent_excerpt_has_roster=False)
    return {"draft": draft, "outline": outline, "textbook": textbook, "markdown": markdown, "components": components,
            "source_digest": hashlib.sha256(data).hexdigest(), "issues": issues}


def calendar_from_outline(filename, outline, textbook):
    if not outline["lessons"]:
        return {"version": 1, "source_name": filename, "error": "大纲缺少可用课次表，请开启 AI 纠错与补全。"}
    try:
        rows, book, warnings = calendar_engine.build_rows(outline["lessons"], textbook, False)
        return normalize_calendar_data({"version": 1, "source_name": filename, "lessons": rows, "教材": book,
            "textbook_source": textbook, "mode": "rule", "warnings": warnings})
    except ServiceError as exc:
        return {"version": 1, "source_name": filename, "error": str(exc)}


def normalize_outline(raw):
    if not isinstance(raw, dict):
        raise ServiceError("补全后的教学大纲格式不正确。")
    def text(value, limit=12000, *, numeric=False, label="教学大纲"):
        if numeric and type(value) in (int,float):
            number = Decimal(str(value))
            if not number.is_finite():
                raise ServiceError(label+"必须为有限数字。")
            value = format(number,"f").rstrip("0").rstrip(".") if "." in format(number,"f") else str(number)
        if not isinstance(value, str) or len(value) > limit or re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", value):
            raise ServiceError(label+"字段格式不正确或内容过长。")
        return value.strip()
    result = {key: text(raw.get(key, "")) for key in ("课程名称", "课程负责人", "课程简介")}
    for key in ("课程目标", "学习资料", "学习网站"):
        values = raw.get(key, [])
        if not isinstance(values, list) or len(values) > 100:
            raise ServiceError("教学大纲列表格式不正确。")
        result[key] = [text(value) for value in values]
    rows = raw.get("lessons", [])
    if not isinstance(rows, list) or not 1 <= len(rows) <= 100:
        raise ServiceError("教学大纲应包含 1–100 个课次。")
    seen = set()
    result["lessons"] = []
    for row in rows:
        if not isinstance(row, dict):
            raise ServiceError("教学大纲课次格式不正确。")
        item = {key: text(row.get(key, ""),numeric=key in {"课次","学时"},label=f"第 {len(seen)+1} 课次的{key}") for key in ("课次", "学时", "章目名称", "教学目标及教学内容", "教学方式", "支撑课程目标", "课程思政")}
        if not item["课次"].isdigit() or int(item["课次"]) < 1 or int(item["课次"]) in seen:
            raise ServiceError("课次应为不重复的正整数。")
        if hours(item["学时"]) <= 0:
            raise ServiceError("每课次学时必须大于 0。")
        seen.add(int(item["课次"]))
        result["lessons"].append(item)
    result["lessons"].sort(key=lambda row: int(row["课次"]))
    assess = raw.get("考核", {})
    if not isinstance(assess, dict):
        raise ServiceError("课程考核数据格式不正确。")
    result["考核"] = {key: text(assess.get(key, ""),numeric=key!="原表",label="课程考核的"+key) for key in ("平时比例", "期末比例", "原表")}
    for key in ("平时项", "期末项"):
        values = assess.get(key, [])
        if not isinstance(values, list) or len(values) > 50:
            raise ServiceError("课程考核项目格式不正确。")
        result["考核"][key] = [text(value, 200) for value in values]
    for key in ("source_digest", "completion_id"):
        if key in raw:
            result[key] = text(raw[key], 64)
    return result


def enrich_outline(original, compatible):
    """Fill missing legacy fields without changing the original script's parsing functions."""
    result = copy.deepcopy(original)
    for key in ("课程名称", "课程简介", "课程目标", "学习资料", "学习网站"):
        if not result.get(key) and compatible.get(key):
            result[key] = copy.deepcopy(compatible[key])
    by_id = {row["课次"]: row for row in compatible.get("lessons", [])}
    for row in result.get("lessons", []):
        number = calendar_engine.pick(row, "课次")
        if number in by_id:
            combined = {**by_id[number], **row}
            row.clear()
            row.update(combined)
    if not (result.get("考核") or {}).get("平时比例"):
        result["考核"] = copy.deepcopy(compatible.get("考核", result.get("考核", {})))
    return result
