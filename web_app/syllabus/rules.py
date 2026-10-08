"""把模型输出收成向导草稿，并用已确认的程序规则改写状态。"""
from __future__ import annotations

import json
import re
from pathlib import Path

from web_app.relation_grid import ratios_from_payload, relation_payload_from_grid
from web_app.service import ServiceError, _normalize_payload
from web_app.syllabus.grad_matrix import NO_MATRIX_REASON, matrix_grad_items
from web_app.syllabus.major import resolve_majors

PROMPT_PATH = Path(__file__).with_name("prompt_system_v2.txt")
COURSE_KEYS = (
    "course_name",
    "credits",
    "hours",
    "course_type",
    "course_code",
    "college",
)
TERM_KEYS = (
    "school_year_term",
    "year_start",
    "year_end",
    "semester",
    "teacher",
    "major",
    "class_name",
    "student_count",
    "exam_count",
)
TERM_LATER_REASON = "学期导入时填写"
_CONCRETE_TERM = re.compile(r"(\d{4})\s*[-–—]\s*(\d{4})\s*学年第\s*([12])\s*学期")
_RELATIVE_TERM = re.compile(r"第\s*\d+\s*学期|\d+\s*/\s*\d+")
_MODULE = re.compile(r"课程模块\s*[：:]\s*(\S+)")
_CATEGORY = re.compile(r"课程类别\s*[：:]\s*(\S+)")
_NATURE = re.compile(r"课程性质\s*[：:]\s*(\S+)")
_MAJOR = re.compile(r"适用专业\s*[：:]\s*(\S+)")
_LEADER = re.compile(r"课程负责人\s*[：:]\s*(\S+)")
_TEACHER = re.compile(r"任课教师\s*[：:]\s*(\S+)")
_CLASS_LINE = re.compile(r"班级取值：([^）\n]+)")


def load_prompt() -> str:
    return PROMPT_PATH.read_text(encoding="utf-8")


def field(value: str = "", status: str = "需手填", source: str = "", reason: str = "", candidates: list | None = None) -> dict:
    filled = status == "已填" and str(value or "").strip() not in {"", "需手填", "手填"}
    return {
        "value": str(value).strip() if filled else "",
        "status": "已填" if filled else "需手填",
        "source": source if filled else "",
        "reason": "" if filled else (reason or "需要老师填写"),
        "candidates": candidates or [],
    }


def _raw_field(item) -> dict:
    return item if isinstance(item, dict) else {}


def take_filled(item, *, source: str = "") -> dict:
    raw = _raw_field(item)
    value = str(raw.get("value") or "").strip()
    if value in {"需手填", "手填", "null", "None"}:
        value = ""
    if raw.get("status") == "已填" and value:
        return field(value, "已填", source or str(raw.get("source") or ""))
    return field("", "需手填", reason=str(raw.get("reason") or "需要老师填写"))


def _add_candidate(bucket: list[dict], label: str, value: str) -> None:
    value = (value or "").strip(" ：:，,；;|")
    if not value or value in {"需手填", "无"}:
        return
    label = label.strip()
    if any(item["label"] == label and item["value"] == value for item in bucket):
        return
    bucket.append({"label": label, "value": value})


def _scan(pattern: re.Pattern, text: str) -> list[str]:
    return [match.group(1).strip() for match in pattern.finditer(text or "")]


def _labeled_values(text: str, label: str) -> list[str]:
    """读“标签：值”，也读表格里紧跟在标签后面的一格。"""
    values: list[str] = []
    pattern = re.compile(rf"{re.escape(label)}\s*[：:]\s*(\S+)")
    values.extend(match.group(1).strip() for match in pattern.finditer(text or ""))
    for line in (text or "").splitlines():
        if label not in line or "|" not in line:
            continue
        cells = [cell.strip() for cell in line.split("|")]
        for index, cell in enumerate(cells[:-1]):
            if cell != label:
                continue
            value = cells[index + 1].strip()
            if value and value != label:
                values.append(value)
    unique: list[str] = []
    for value in values:
        if value not in unique:
            unique.append(value)
    return unique


def course_type_field(syllabus_text: str, register_text: str, register: dict | None) -> dict:
    candidates: list[dict] = []
    modules = _labeled_values(syllabus_text, "课程模块")
    categories = _labeled_values(syllabus_text, "课程类别")
    for value in modules + categories + _labeled_values(syllabus_text, "课程性质"):
        _add_candidate(candidates, f"大纲：{value}", value)
    if modules and categories:
        combined = f"{modules[0]}/{categories[0]}"
        _add_candidate(candidates, f"大纲：{combined}", combined)
    natures = []
    if register and register.get("course_type"):
        natures.append(register["course_type"])
    natures.extend(_scan(_NATURE, register_text))
    for value in natures:
        _add_candidate(candidates, f"成绩登记表：{value}", value)
    reason = "课程性质不自动填入，请点选大纲或成绩登记表中的写法"
    if not candidates:
        reason = "文档里没有课程模块、课程类别或课程性质"
    return field("", "需手填", reason=reason, candidates=candidates)


def teacher_field(syllabus_text: str, register_text: str, register: dict | None, ai_item) -> dict:
    teacher = ""
    if register and register.get("teacher"):
        teacher = str(register["teacher"]).strip()
    if not teacher:
        found = _scan(_TEACHER, register_text)
        if found:
            teacher = found[0]
    if teacher:
        return field(teacher, "已填", "成绩登记表")
    ai = take_filled(ai_item)
    source = str(_raw_field(ai_item).get("source") or "")
    leader = ""
    leaders = _scan(_LEADER, syllabus_text)
    if leaders:
        leader = leaders[0]
    if ai["status"] == "已填" and ("登记" in source or "文档B" in source) and ai["value"] != leader:
        return field(ai["value"], "已填", "成绩登记表")
    candidates = []
    if leader:
        _add_candidate(candidates, f"大纲·课程负责人：{leader}", leader)
    elif ai["status"] != "已填":
        candidate = str(_raw_field(ai_item).get("candidate") or "").strip()
        if candidate and "负责人" in str(_raw_field(ai_item).get("reason") or "") + candidate:
            _add_candidate(candidates, f"大纲·课程负责人：{candidate}", candidate.split("：")[-1])
    return field(
        "",
        "需手填",
        reason="没有成绩登记表上的任课教师。大纲的课程负责人只作为候选，不会自动填入。",
        candidates=candidates,
    )


def major_field(syllabus_text: str, register_text: str, register: dict | None, ai_item) -> dict:
    applicable = _labeled_values(syllabus_text, "适用专业")
    if not applicable:
        ai = take_filled(ai_item)
        source = str(_raw_field(ai_item).get("source") or "")
        if ai["status"] == "已填" and "适用专业" in source:
            applicable = [ai["value"]]
    if applicable:
        return field(applicable[0], "已填", "大纲·适用专业")
    classes: list[str] = []
    if register and register.get("classes"):
        classes.extend(register["classes"])
    for chunk in _CLASS_LINE.findall(register_text or ""):
        classes.extend(part.strip() for part in re.split(r"[、,，]", chunk) if part.strip())
    resolved = resolve_majors(classes)
    if resolved["status"] == "已填":
        return field(resolved["value"], "已填", resolved["source"])
    reason = resolved["reason"]
    if not classes:
        reason = "大纲没有适用专业，成绩登记表也没有可推导的班级"
    return field("", "需手填", reason=reason, candidates=resolved["candidates"])


def term_parts(school_year_term: str) -> tuple[str, str, str]:
    match = _CONCRETE_TERM.search(school_year_term or "")
    if not match:
        return "", "", ""
    return match.group(1), match.group(2), match.group(3)


def school_year_field(syllabus_text: str, register: dict | None, ai_item) -> dict:
    concrete = ""
    if register and register.get("school_year_term"):
        concrete = register["school_year_term"]
    if not concrete:
        ai = take_filled(ai_item)
        if ai["status"] == "已填" and _CONCRETE_TERM.search(ai["value"]):
            concrete = _CONCRETE_TERM.search(ai["value"]).group(0).replace(" ", "")
            year_start, year_end, semester = term_parts(ai["value"])
            concrete = f"{year_start}-{year_end}学年第{semester}学期"
    if concrete and _CONCRETE_TERM.search(concrete):
        return field(concrete, "已填", "成绩登记表标题" if register and register.get("school_year_term") else str(_raw_field(ai_item).get("source") or "成绩登记表标题"))
    relative = _RELATIVE_TERM.search(syllabus_text or "")
    original = relative.group(0) if relative else str(_raw_field(ai_item).get("value") or _raw_field(ai_item).get("candidate") or "")
    candidates = []
    if original.strip():
        _add_candidate(candidates, f"原文：{original.strip()}", original.strip())
    return field("", "需手填", reason="相对学期不换算，请填写具体学年学期", candidates=candidates)


def grad_items(ai_items, objective_count: int, syllabus_text: str = "") -> list[dict]:
    parsed = matrix_grad_items(syllabus_text or "", objective_count)
    if parsed is not None:
        return parsed
    items = list(ai_items or [])
    count = max(objective_count, len(items))
    result = []
    for index in range(count):
        item = items[index] if index < len(items) and isinstance(items[index], dict) else {}
        strength = str(item.get("strength") or "").strip().upper()
        if strength not in {"H", "M", "L"}:
            strength = ""
        result.append(
            {
                "objective": f"课程目标{index + 1}",
                "indicator": field("", "需手填", reason=NO_MATRIX_REASON),
                "requirement": field("", "需手填", reason=NO_MATRIX_REASON),
                "strength": strength,
            }
        )
    return result


def relation_from_ai(relation: dict, objective_count: int) -> list[list[str]] | None:
    links = (relation or {}).get("links") or []
    count = int((relation or {}).get("objectives_count") or objective_count or 0)
    useful = [link for link in links if isinstance(link, dict)]
    if not useful or count <= 0:
        return None
    header = ["考核环节", "占比", "考核方式"] + [f"课程目标{i + 1}" for i in range(count)] + ["小计"]
    grid = [header]
    for link in useful:
        methods = link.get("methods") or []
        if not methods:
            grid.append([str(link.get("name") or ""), str(link.get("ratio_text") or ""), ""] + [""] * count + [""])
            continue
        for index, method in enumerate(methods):
            if not isinstance(method, dict):
                continue
            weights = method.get("weights") or {}
            grid.append(
                [
                    str(link.get("name") or "") if index == 0 else "",
                    str(link.get("ratio_text") or "") if index == 0 else "",
                    str(method.get("name") or ""),
                ]
                + [str(weights.get(f"课程目标{i + 1}", "") or "") for i in range(count)]
                + [str(method.get("subtotal_text") or "")]
            )
    return grid


def validate_relation_grid(grid) -> dict:
    errors: list[str] = []
    payload = None
    try:
        payload = relation_payload_from_grid(grid)
    except ValueError as exc:
        errors.append(str(exc))
    if payload:
        for link in payload.get("links") or []:
            name = str(link.get("name") or "")
            if not any(token in name for token in ("平时", "期中", "期末")):
                errors.append(f"{name or '考核环节'} 的名称需要包含“平时”“期中”或“期末”")
            total = 0.0
            for method in link.get("methods") or []:
                supports = method.get("supports") or {}
                total += sum(float(value) for value in supports.values())
            if abs(total - 1) > 0.02:
                errors.append(f"{name} 的目标权重合计应为 100%，当前约为 {round(total * 100)}%")
        try:
            _normalize_payload(payload)
        except ServiceError as exc:
            errors.append(str(exc))
    ratios = ratios_from_payload(payload) if payload else {"usual": 0, "midterm": 0, "final": 0}
    weights = {}
    if payload:
        for key, value in (payload.get("objectives_total_weights") or {}).items():
            weights[key] = f"{float(value) * 100:.1f}%"
    return {
        "ok": not errors and payload is not None,
        "errors": errors,
        "ratios": ratios,
        "weights": weights,
        "payload": payload,
        "grid": grid,
    }


def parse_model_json(text: str) -> dict:
    raw = (text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?", "", raw, count=1).strip()
        raw = re.sub(r"```$", "", raw).strip()
    start = raw.find("{")
    end = raw.rfind("}")
    if start >= 0 and end > start:
        raw = raw[start : end + 1]
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("模型没有返回 JSON 对象")
    return data


def build_draft(ai: dict, syllabus_text: str, register_text: str, register: dict | None, *, matrix_text: str | None = None) -> dict:
    basic = ai.get("course_basic_info") if isinstance(ai.get("course_basic_info"), dict) else {}
    fields = {}
    for key in ("course_name", "credits", "hours", "course_code", "college"):
        chosen = take_filled(basic.get(key))
        if key == "course_code" and register and register.get("course_code"):
            chosen = field(register["course_code"], "已填", "成绩登记表")
        if key == "college" and chosen["status"] != "已填" and register and register.get("college"):
            chosen = field(register["college"], "已填", "成绩登记表")
        if key == "course_name" and chosen["status"] != "已填" and register and register.get("course_name"):
            chosen = field(register["course_name"], "已填", "成绩登记表")
        if key == "credits" and chosen["status"] != "已填" and register and register.get("credits"):
            chosen = field(str(register["credits"]), "已填", "成绩登记表")
        fields[key] = chosen
    fields["course_type"] = course_type_field(syllabus_text, register_text, register)
    # 教师、专业、班级、学年学期和人数随每次登记表变化，不从大纲写进课程。
    for key in TERM_KEYS:
        fields[key] = field("", "需手填", reason=TERM_LATER_REASON)

    description = take_filled(ai.get("course_description"))
    objectives = []
    if ai.get("objectives_status") == "已填":
        for item in ai.get("objectives") or []:
            if not isinstance(item, dict):
                continue
            text = str(item.get("text") or "").strip()
            if text and text != "需手填":
                objectives.append(text)
    relation = relation_from_ai(ai.get("relation_table") or {}, len(objectives))
    checked = validate_relation_grid(relation) if relation else {
        "ok": False,
        "errors": ["没有可用的关系表"],
        "ratios": {"usual": 0, "midterm": 0, "final": 0},
        "weights": {},
        "payload": None,
        "grid": [],
    }
    return {
        "fields": fields,
        "groups": {"course": list(COURSE_KEYS), "term": list(TERM_KEYS)},
        "course_description": description,
        "objectives": objectives,
        "objectives_status": "已填" if objectives else "需手填",
        "grad_req_map": grad_items(ai.get("grad_req_map"), len(objectives), matrix_text if matrix_text is not None else syllabus_text),
        "relation": {
            "ok": checked["ok"],
            "errors": checked["errors"],
            "ratios": checked["ratios"],
            "weights": checked["weights"],
            "grid": checked["grid"] or [],
        },
    }


def plain_text(value, label: str = "字段") -> str:
    """确认页的值必须落成字符串。对象取 value，其他类型拒绝。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        raise ServiceError(f"{label}必须是字符串")
    if isinstance(value, (int, float)):
        return str(value).strip()
    if isinstance(value, str):
        text = value.strip()
        if text == "[object Object]":
            raise ServiceError(f"{label}格式不正确")
        return text
    if isinstance(value, dict):
        if "value" in value:
            return plain_text(value.get("value"), label)
        if "text" in value:
            return plain_text(value.get("text"), label)
        raise ServiceError(f"{label}必须是字符串")
    raise ServiceError(f"{label}必须是字符串")


def _description_text(body: dict) -> str:
    raw = body.get("description")
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        raw = body.get("course_description")
    return plain_text(raw, "课程简介")


def settings_from_wizard(body: dict) -> tuple[dict, dict]:
    """确认页提交的内容 → 课程设置 + 学期字段。关系表必须能通过校验。"""
    if not isinstance(body, dict):
        raise ServiceError("确认内容格式不正确")
    incoming = body.get("fields") if isinstance(body.get("fields"), dict) else {}

    labels = {
        "course_name": "课程名称",
        "credits": "学分",
        "hours": "学时",
        "course_type": "课程性质",
        "course_code": "课程代码",
        "college": "开课学院",
    }

    def text_of(key: str) -> str:
        return plain_text(incoming.get(key, ""), labels.get(key, "字段"))

    objectives = []
    for item in body.get("objectives") or []:
        text = plain_text(item, "课程目标")
        if text:
            objectives.append(text)
    grad = []
    for index, item in enumerate(body.get("grad_req_map") or []):
        if not isinstance(item, dict):
            raise ServiceError("毕业要求必须是字符串")
        grad.append(
            {
                "objective": f"课程目标{index + 1}",
                "requirement": plain_text(item.get("requirement"), "毕业要求"),
                "indicator": plain_text(item.get("indicator"), "指标点"),
                "strength": plain_text(item.get("strength"), "关联度").upper(),
            }
        )
    grid = body.get("relation_grid")
    if not grid:
        grid = (body.get("relation") or {}).get("grid")
    checked = validate_relation_grid(grid)
    if not checked["ok"]:
        raise ServiceError("；".join(checked["errors"]) or "关系表未通过校验")
    course_name = text_of("course_name") or plain_text(body.get("course_name"), "课程名称")
    if not course_name:
        raise ServiceError("请填写课程名称")
    # 学年、学期、教师、专业、班级和人数不进课程层，等导入成绩登记表时写入学期。
    basic = {
        "course_name": course_name,
        "credits": text_of("credits"),
        "hours": text_of("hours"),
        "course_type": text_of("course_type"),
        "course_code": text_of("course_code"),
        "college": text_of("college"),
    }
    settings = {
        "mode": "forward",
        "course_open_info": {
            "course_name": course_name,
            "department": text_of("college"),
        },
        "course_basic_info": basic,
        "course_description": _description_text(body),
        "objective_requirements": objectives,
        "grad_req_map": grad,
        "relation_grid": checked["grid"],
        "relation_payload": checked["payload"],
        "ratios": checked["ratios"],
        "report_style": "专业",
        "word_limit": 200,
    }
    return settings, {"course_name": course_name}
