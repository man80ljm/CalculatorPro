"""Bounded AI repair and validation, layered on top of original syllabus rules."""
import copy
import json
import re
from decimal import Decimal
from pathlib import Path

from web_app.service import ServiceError
from . import client, compatibility
from .rules import field, parse_model_json, validate_relation_grid
from .prefilter import prefilter, still_has_roster
from web_app.teaching_calendar.source import apply_calendar_ai

PROMPT = Path(__file__).with_name("prompt_completion.txt")
COURSE_FIELDS = ("course_name", "credits", "hours", "course_type", "course_code", "college")


def value(raw):
    return raw.get("value", "") if isinstance(raw, dict) else raw


def text(raw, label, maximum=12000, *, numeric=False):
    raw = value(raw)
    if numeric and type(raw) in (int,float):
        number = Decimal(str(raw))
        if not number.is_finite():
            raise ServiceError(label+"必须为有限数字。")
        raw = format(number,"f").rstrip("0").rstrip(".") if "." in format(number,"f") else str(number)
    if not isinstance(raw, str) or len(raw) > maximum or re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", raw):
        raise ServiceError(label + "格式不正确或内容过长。")
    return raw.strip()


def draft_errors(draft, outline):
    errors = []
    if not text(draft["fields"].get("course_name", ""), "课程名称", 200):
        errors.append("请填写课程名称")
    goals = draft.get("objectives")
    if not isinstance(goals,list) or not goals or len(goals)>compatibility.MAX_GOALS or any(not isinstance(item,str) or not item.strip() for item in goals):
        errors.append("课程目标存在缺项或格式不正确")
    goals = goals if isinstance(goals,list) else []
    grad = draft.get("grad_req_map") or []
    if not isinstance(grad,list) or len(grad) != len(goals) or any(not isinstance(item,dict) or not value(item.get("requirement")) or not value(item.get("indicator")) or value(item.get("strength")) not in ("H","M","L") for item in grad):
        errors.append("毕业要求、指标点或支撑强度存在缺项")
    checked = draft["relation"]
    errors.extend(checked.get("errors") or [])
    if not checked.get("ok"):
        errors.append("考核对应关系表未通过校验")
    grid = checked.get("grid") or []
    if grid and len(grid[0]) != len(draft.get("objectives", [])) + 4:
        errors.append("考核表目标列数与课程目标数量不一致")
    if checked.get("payload"):
        totals = checked["payload"].get("objectives_total_weights") or {}
        for number in range(1, len(draft.get("objectives", [])) + 1):
            if float(totals.get(f"课程目标{number}", 0)) <= 0:
                errors.append(f"课程目标{number}没有考核方式支撑")
    try:
        normalized = compatibility.normalize_outline(outline)
        expected = compatibility.hours(value(draft["fields"].get("hours", "")))
        actual = sum(compatibility.hours(row["学时"]) for row in normalized["lessons"])
        for row in normalized["lessons"]:
            refs = [int(number) for number in re.findall(r"课程目标\s*(\d+)",row["支撑课程目标"])]
            if any(number < 1 or number > len(goals) for number in refs):
                errors.append(f"第 {row['课次']} 课次引用了未定义的课程目标，请补齐目标或修正引用")
        if actual != expected:
            errors.append(f"课次学时合计 {actual} 与总学时 {expected} 不一致")
    except ServiceError as exc:
        errors.append(str(exc))
    return list(dict.fromkeys(errors))


def merge(analysis, result, *, checkpoint=None):
    if not isinstance(result, dict):
        raise ServiceError("AI 补全内容必须为 JSON 对象。")
    draft = copy.deepcopy(analysis["draft"])
    fields = result.get("fields")
    if not isinstance(fields, dict):
        raise ServiceError("AI 未返回课程基本信息。")
    inferred = []
    def preserve(outline=None):
        if checkpoint:
            draft["completion"] = {"ok":False,"errors":[],"changes":analysis["issues"],"warnings":[],"inferred_fields":list(inferred)}
            checkpoint(draft,analysis["outline"] if outline is None else outline)
    for key in COURSE_FIELDS:
        new = text(fields.get(key, ""), key, 200,numeric=key in {"hours","credits"})
        old = value(draft["fields"].get(key)) or ""
        same = old == new
        if old and key in {"hours","credits"}:
            same = compatibility.hours(old) == compatibility.hours(new)
        if old and not same:
            raise ServiceError(f"AI 改动了已明确的课程信息：{key}，应保留 {old}")
        if same: new = old
        if new != old:
            inferred.append("fields." + key)
        draft["fields"][key] = field(new, "已填" if new else "需手填", source="AI 补全" if new != old else draft["fields"].get(key,{}).get("source","大纲·基本信息"))
    preserve()
    description = text(result.get("course_description"), "课程简介")
    if not description:
        raise ServiceError("课程简介尚未补全。")
    draft["course_description"] = field(description, "已填", source="AI 整理")
    preserve()
    goals = result.get("objectives")
    if not isinstance(goals, list) or not 1 <= len(goals) <= compatibility.MAX_GOALS:
        raise ServiceError("应返回 1–8 个课程目标，与现有计算器编辑器一致。")
    draft["objectives"] = [text(goal, "课程目标") for goal in goals]
    preserve()
    grad = result.get("grad_req_map")
    if not isinstance(grad, list) or len(grad) != len(goals):
        raise ServiceError("毕业要求对应关系数量应与课程目标一致。")
    draft["grad_req_map"] = []
    for index, item in enumerate(grad):
        if not isinstance(item, dict):
            raise ServiceError("毕业要求对应关系格式不正确。")
        strength = text(item.get("strength"), "支撑强度", 1).upper()
        row = {"objective": f"课程目标{index+1}", "strength": strength}
        for key in ("requirement","indicator"):
            new = text(item.get(key), "毕业要求或指标点", 1000,numeric=True)
            old = analysis["draft"]["grad_req_map"][index][key] if index < len(analysis["draft"]["grad_req_map"]) else ""
            generated = new != value(old)
            row[key] = field(new, "已填" if new else "需手填", source="AI 补全" if generated else "大纲·对应关系表")
            if generated:
                inferred.append(f"grad_req_map.{index}.{key}")
        draft["grad_req_map"].append(row)
    for index, goal in enumerate(draft["objectives"]):
        if index >= len(analysis["draft"]["objectives"]) or goal != analysis["draft"]["objectives"][index]:
            inferred.append(f"objectives.{index}")
    preserve()
    grid = result.get("relation_grid")
    if not isinstance(grid, list) or not 2 <= len(grid) <= 101:
        raise ServiceError("AI 未返回有效的考核对应关系表。")
    checked_grid = []
    for row in grid:
        if not isinstance(row, list) or len(row) != len(goals) + 4:
            raise ServiceError("考核对应关系表的列数不正确。")
        cells = [text(cell, "考核表", 500,numeric=bool(checked_grid) and index not in (0,2)) for index,cell in enumerate(row)]
        if checked_grid:
            for cell in [cells[1], *cells[3:]]:
                if cell and not re.fullmatch(r"\d+(?:\.\d+)?\s*[%％]?", cell):
                    raise ServiceError("成绩占比和目标权重必须为非负数字或百分数。")
        checked_grid.append(cells)
    draft["relation"] = validate_relation_grid(checked_grid)
    components = analysis["components"]
    if components and sum(Decimal(item["ratio"]) for item in components) == 100 and draft["relation"]["ok"]:
        original = {"usual":Decimal(0), "midterm":Decimal(0), "final":Decimal(0)}
        for item in components:
            name = item["name"]
            key = "midterm" if "期中" in name else "final" if re.search(r"期末|期终|课程作业|课程设计|结课|考试", name) else "usual"
            original[key] += Decimal(item["ratio"]) / 100
        for key, ratio in original.items():
            if abs(Decimal(str(draft["relation"]["ratios"][key])) - ratio) > Decimal("0.000001"):
                raise ServiceError("AI 改动了原文已经明确且合计正确的成绩占比。")
    preserve()
    outline = copy.deepcopy(result.get("outline", analysis["outline"]))
    normalized = compatibility.normalize_outline(outline)
    old_rows = analysis["outline"]["lessons"]
    if old_rows:
        if len(old_rows) != len(normalized["lessons"]):
            raise ServiceError("AI 改动了原文已有课次。")
        try:
            stable_hours = sum(compatibility.hours(row["学时"]) for row in old_rows) == compatibility.hours(value(analysis["draft"]["fields"].get("hours","")))
        except ServiceError:
            stable_hours = False
        for old, new in zip(old_rows, normalized["lessons"]):
            for key in ("课次","章目名称"):
                if old[key] != new[key]:
                    raise ServiceError("AI 改动了原文已有课次或章目。")
            if stable_hours and old["学时"] and compatibility.hours(old["学时"]) != compatibility.hours(new["学时"]):
                raise ServiceError("AI 改动了原文已有课次学时。")
    normalized["课程名称"] = value(draft["fields"]["course_name"])
    normalized["课程简介"] = description
    normalized["课程目标"] = draft["objectives"]
    if draft["relation"]["ok"]:
        payload = draft["relation"]["payload"]
        assess = {"平时项": [], "期末项": [], "原表": "\n".join(" | ".join(row) for row in checked_grid)}
        for link in payload["links"]:
            final = "期末" in link["name"]
            assess["期末项" if final else "平时项"].extend(method["name"] for method in link["methods"])
        assess["平时比例"] = str(Decimal(str(draft["relation"]["ratios"]["usual"])) * 100)
        assess["期末比例"] = str(Decimal(str(draft["relation"]["ratios"]["final"])) * 100)
        normalized["考核"] = assess
    preserve(normalized)
    calendar = compatibility.calendar_from_outline(analysis["draft"]["teaching_calendar"]["source_name"], normalized, analysis["textbook"])
    if not calendar.get("error") and isinstance(result.get("teaching_calendar"), dict):
        calendar = apply_calendar_ai(calendar, result["teaching_calendar"])
    draft["teaching_calendar"] = calendar
    changes = result.get("changes", [])
    if not isinstance(changes, list) or len(changes) > 100:
        raise ServiceError("补全说明格式不正确。")
    changes = [text(change, "补全说明", 1000) for change in changes]
    source_notice = next((issue for issue in analysis["issues"] if issue.startswith("上传的是课程教案")),"")
    if source_notice: changes.insert(0,source_notice)
    errors = draft_errors(draft, normalized)
    draft["completion"] = {"ok": not errors, "errors": errors, "changes": changes, "warnings": [], "inferred_fields": inferred}
    return draft, normalized


def run(analysis, *, allow_completion=True, progress=None, call_offset=0):
    progress = progress or (lambda **_: None)
    draft, outline = copy.deepcopy(analysis["draft"]), copy.deepcopy(analysis["outline"])
    errors = draft_errors(draft, outline)
    draft["completion"] = {"ok": not errors, "errors": errors, "changes": analysis["issues"], "warnings": [], "inferred_fields":[]}
    if not allow_completion:
        return draft, outline
    excerpt = prefilter(analysis["markdown"])
    if still_has_roster(excerpt) or still_has_roster(json.dumps(outline, ensure_ascii=False)):
        draft["completion"]["warnings"].append("大纲中仍含学生名单，未发送给 AI；请修改草稿或移除名单后重试。")
        return draft, outline
    request = {"草稿": {"fields": {key: value(draft["fields"].get(key, "")) for key in COURSE_FIELDS},
        "course_description": value(draft["course_description"]), "objectives": draft["objectives"],
        "grad_req_map": draft["grad_req_map"], "relation_grid": draft["relation"]["grid"]},
        "教学内容": outline, "已识别成绩占比": analysis["components"], "问题": analysis["issues"], "原文摘录":excerpt[:100000]}
    messages = [{"role":"system","content":PROMPT.read_text(encoding="utf-8")},{"role":"user","content":json.dumps(request, ensure_ascii=False)}]
    calls = max(0,min(2,int(call_offset)))
    remaining = 2 - calls
    if not remaining:
        draft["completion"]["warnings"].append("本次任务的两次 AI 请求已用完，已保留草稿；可手动修改或重新发起补全。")
    for attempt in range(remaining):
        calls += 1
        progress(percent=45 if attempt == 0 else 80, api_calls=calls,
                 stage_label="AI 正在检查矛盾并补全必要内容…" if attempt == 0 else "正在修正校验发现的问题（仅重写一次）…")
        try:
            response = client.complete(messages, max_tokens=12000, timeout=120)
        except Exception:
            draft["completion"]["warnings"].append("AI 调用失败或超时，已保留原文和兼容解析草稿；可以手动修改，或稍后重新补全。")
            break
        partial = None
        def keep_valid_sections(valid_draft, valid_outline):
            nonlocal partial
            partial = copy.deepcopy(valid_draft),copy.deepcopy(valid_outline)
        try:
            parsed = parse_model_json(response.get("content",""))
            candidate, candidate_outline = merge(analysis, parsed,checkpoint=keep_valid_sections)
            progress(percent=75 if attempt == 0 else 90, api_calls=calls, stage_label="正在核验目标、成绩占比和课次学时…")
            if candidate["completion"]["ok"]:
                draft, outline = candidate, candidate_outline
                break
            reason = "；".join(candidate["completion"]["errors"])
            draft, outline = candidate, candidate_outline
        except (ServiceError, ValueError, TypeError, KeyError, IndexError, OverflowError) as exc:
            reason = str(exc) if isinstance(exc, ServiceError) else "AI JSON 结构或字段格式不正确"
            if partial is not None:
                draft,outline = partial
                draft["completion"]["errors"] = list(dict.fromkeys([*draft_errors(draft,outline),reason]))
                draft["completion"]["ok"] = False
        if attempt < remaining-1:
            messages.extend([{"role":"assistant","content":response.get("content","")[:100000]},
                             {"role":"user","content":"请只修正以下校验错误，再完整返回同样结构的 JSON。保留已有原文信息：\n"+reason}])
        else:
            draft["completion"]["warnings"].append("AI 修正一次后仍有问题：" + reason + "。已保留可编辑草稿。")
    draft["meta"] = {"calls":calls, "completion_used":True}
    return draft, outline
