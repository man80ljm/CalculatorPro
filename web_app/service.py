"""把桌面端的计算、导出和 AI 报告跑在每次请求独立的临时目录里。"""
from __future__ import annotations

import json
import os
import shutil
import tempfile
import zipfile
from io import BytesIO
from typing import Any

import pandas as pd

from core import GradeProcessor
from core_app.course_docs import export_course_basic_docx, export_grad_req_docx
from core_app.relation_export import write_relation_docx_from_payload
from core_app.report_builder import ReportBuilder
from io_app.excel_templates import create_forward_template, create_reverse_template
from utils import get_resource_path, override_outputs_dir
from web_app.ai_runner import generate_answers

SPREAD_MAP = {
    "大跨度（14-23分）": "large",
    "中跨度（7-13分）": "medium",
    "小跨度（2-6分）": "small",
    "large": "large",
    "medium": "medium",
    "small": "small",
}
DIST_MAP = {
    "标准正态": "normal",
    "高分倾向": "left_skewed",
    "低分倾向": "right_skewed",
    "两极分化": "bimodal",
    "档位打分": "discrete",
    "完全随机": "uniform",
    "normal": "normal",
    "left_skewed": "left_skewed",
    "right_skewed": "right_skewed",
    "bimodal": "bimodal",
    "discrete": "discrete",
    "uniform": "uniform",
}

AI_DISABLED_MESSAGE = (
    "服务器未配置 DEEPSEEK_API_KEY，无法生成 AI 分析报告。"
    "成绩计算、模板下载和 xlsx/docx 导出不受影响。"
)


class ServiceError(ValueError):
    pass


class TextValue:
    def __init__(self, value: Any):
        self._value = "" if value is None else str(value)

    def text(self) -> str:
        return self._value


class StatusLabel:
    def __init__(self):
        self.text = ""

    def setText(self, text: str) -> None:
        self.text = "" if text is None else str(text)


def deepseek_api_key() -> str:
    return os.environ.get("DEEPSEEK_API_KEY", "").strip()


def ai_status() -> dict:
    if deepseek_api_key():
        return {"enabled": True, "message": "AI 分析报告可用（DeepSeek / deepseek-chat）。密钥只保存在服务器。"}
    return {"enabled": False, "message": AI_DISABLED_MESSAGE}


def _course_name(settings: dict) -> str:
    open_info = settings.get("course_open_info") or {}
    basic = settings.get("course_basic_info") or {}
    name = ""
    if isinstance(open_info, dict):
        name = (open_info.get("course_name") or "").strip()
    if not name and isinstance(basic, dict):
        name = (basic.get("course_name") or "").strip()
    return name or "未命名"


def _num_objectives(payload: dict, settings: dict) -> int:
    count = int(payload.get("objectives_count") or 0)
    if count <= 0:
        count = len(settings.get("objective_requirements") or [])
    if count <= 0:
        raise ServiceError("请先设置课程目标数量，并填写课程考核与课程目标对应关系")
    return count


def _normalize_payload(payload: dict) -> dict:
    if not isinstance(payload, dict):
        raise ServiceError("课程考核与课程目标对应关系格式不正确")
    links = payload.get("links")
    if not isinstance(links, list) or not links:
        raise ServiceError("请先填写课程考核与课程目标对应关系")
    cleaned_links = []
    ratio_sum = 0.0
    for link in links:
        if not isinstance(link, dict):
            raise ServiceError("考核环节格式不正确")
        name = str(link.get("name") or "").strip()
        if not name:
            raise ServiceError("考核环节名称不能为空")
        try:
            ratio = float(link.get("ratio"))
        except (TypeError, ValueError):
            raise ServiceError(f"{name} 的占比不是有效数字")
        if ratio < -1e-9 or ratio > 1 + 1e-9:
            raise ServiceError(f"{name} 的占比需要在 0 到 1 之间")
        methods_in = link.get("methods") or []
        if not isinstance(methods_in, list):
            raise ServiceError(f"{name} 的考核方式格式不正确")
        methods = []
        for method in methods_in:
            if not isinstance(method, dict):
                raise ServiceError(f"{name} 的考核方式格式不正确")
            method_name = str(method.get("name") or "").strip()
            if not method_name:
                raise ServiceError(f"{name} 存在未命名的考核方式")
            supports = method.get("supports") or {}
            if not isinstance(supports, dict):
                raise ServiceError(f"{method_name} 的目标权重格式不正确")
            try:
                subtotal = float(method.get("subtotal"))
            except (TypeError, ValueError):
                raise ServiceError(f"{method_name} 的小计不是有效数字")
            methods.append(
                {
                    "name": method_name,
                    "supports": supports,
                    "subtotal": subtotal,
                }
            )
        if ratio > 1e-9 and not methods:
            raise ServiceError(f"{name} 的占比大于 0，请至少填写一种考核方式")
        ratio_sum += ratio
        cleaned_links.append({"name": name, "ratio": ratio, "methods": methods})
    if abs(ratio_sum - 1.0) > 1e-3:
        raise ServiceError("各考核环节占比之和必须等于 1")
    try:
        objectives_count = int(payload.get("objectives_count") or 0)
    except (TypeError, ValueError):
        objectives_count = 0
    if objectives_count <= 0:
        raise ServiceError("课程目标数量至少为 1")
    normalized = dict(payload)
    normalized["objectives_count"] = objectives_count
    normalized["links"] = cleaned_links
    return normalized


def _ratios(settings: dict) -> tuple[float, float, float]:
    raw = settings.get("ratios") or {}
    if not isinstance(raw, dict):
        raw = {}

    def pick(key: str, default: float) -> float:
        value = raw.get(key, settings.get(key, default))
        try:
            return float(value)
        except (TypeError, ValueError):
            return default

    return pick("usual", 0.3), pick("midterm", 0.0), pick("final", 0.7)


def _mapped_mode(settings: dict, key: str, mapping: dict, default: str) -> str:
    return mapping.get(str(settings.get(key) or "").strip(), default)


def _count_students(path: str, mode: str) -> int:
    header = 1 if mode == "forward" else 0
    frame = pd.read_excel(path, header=header).fillna("")
    columns = list(frame.columns)
    if columns:
        first = str(columns[0]) if columns[0] is not None else ""
        if first.startswith("Unnamed") or first.strip() in ("", "nan"):
            columns[0] = "姓名"
            frame.columns = columns
    if "姓名" not in frame.columns:
        return 0
    count = 0
    for value in frame["姓名"]:
        text = str(value).strip()
        if text and text != "姓名" and text.lower() != "nan":
            count += 1
    return count


def _build_processor(settings: dict, input_path: str, payload: dict) -> GradeProcessor:
    usual, midterm, final = _ratios(settings)
    count = _num_objectives(payload, settings)
    weights = [TextValue(f"{1.0 / count:.6f}") for _ in range(count)]
    requirements = [str(item) for item in (settings.get("objective_requirements") or [])]
    processor = GradeProcessor(
        TextValue(_course_name(settings)),
        TextValue(str(count)),
        weights,
        TextValue(usual),
        TextValue(midterm),
        TextValue(final),
        StatusLabel(),
        input_path,
        course_description=str(settings.get("course_description") or ""),
        objective_requirements=requirements,
        relation_payload=payload,
    )
    noise = settings.get("noise_config")
    if isinstance(noise, dict) and noise:
        processor.set_noise_config(noise)
    return processor


def _write_supporting_docs(settings: dict, payload: dict) -> None:
    from utils import get_outputs_dir

    basic = settings.get("course_basic_info") or {}
    if not isinstance(basic, dict):
        basic = {}
    if not basic.get("course_name"):
        basic = dict(basic)
        basic["course_name"] = _course_name(settings)
    export_course_basic_docx(basic)
    grad_map = settings.get("grad_req_map") or []
    if not isinstance(grad_map, list):
        grad_map = []
    export_grad_req_docx(grad_map)
    write_relation_docx_from_payload(
        os.path.join(get_outputs_dir(), "4.课程考核与课程目标对应关系表.docx"),
        payload,
    )


def _run_grades(processor: GradeProcessor, settings: dict):
    mode = settings.get("mode") or "forward"
    if mode not in ("forward", "reverse"):
        raise ServiceError("模式只能是正向或逆向")
    spread = _mapped_mode(settings, "spread_mode", SPREAD_MAP, "medium")
    distribution = _mapped_mode(settings, "distribution", DIST_MAP, "normal")
    if mode == "forward":
        average = processor.process_forward_grades(spread_mode=spread, distribution=distribution)
    else:
        average = processor.process_reverse_grades(spread_mode=spread, distribution=distribution)
    achievement = getattr(processor, "current_achievement", {}) or {}
    return float(average or 0), {str(key): float(value) for key, value in achievement.items()}


def _load_previous(processor: GradeProcessor, previous_path: str | None) -> None:
    try:
        processor.load_previous_achievement(previous_path or "")
    except Exception as exc:
        raise ServiceError(f"上一学年达成度表读取失败：{exc}") from exc


def _zip_outputs(directory: str) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        names = sorted(os.listdir(directory))
        for name in names:
            if not name.lower().endswith((".xlsx", ".docx")):
                continue
            path = os.path.join(directory, name)
            if os.path.isfile(path):
                archive.write(path, arcname=name)
    return buffer.getvalue()


def _output_names(directory: str) -> list[str]:
    return sorted(
        name
        for name in os.listdir(directory)
        if name.lower().endswith((".xlsx", ".docx")) and os.path.isfile(os.path.join(directory, name))
    )


def _capture_outputs(directory: str) -> list[tuple[str, bytes]]:
    captured = []
    for name in _output_names(directory):
        with open(os.path.join(directory, name), "rb") as handle:
            captured.append((name, handle.read()))
    return captured


def _open_info_for_report(settings: dict) -> dict:
    info = dict(settings.get("course_open_info") or {})
    if not info.get("term"):
        info["term"] = info.get("semester") or "1"
    if not info.get("course_name"):
        info["course_name"] = _course_name(settings)
    return info


def _prepare_workspace(excel_bytes: bytes, previous_bytes: bytes | None, settings: dict):
    """Yields (work_dir, outputs_dir, input_path, previous_path, payload). Caller cleans up."""
    payload = _normalize_payload(settings.get("relation_payload") or {})
    settings = dict(settings)
    settings["relation_payload"] = payload
    work_dir = tempfile.mkdtemp(prefix="calculatorpro-")
    outputs_dir = os.path.join(work_dir, "outputs")
    os.makedirs(outputs_dir, exist_ok=True)
    input_path = os.path.join(work_dir, "grades.xlsx")
    with open(input_path, "wb") as handle:
        handle.write(excel_bytes)
    previous_path = None
    if previous_bytes:
        previous_path = os.path.join(work_dir, "previous.xlsx")
        with open(previous_path, "wb") as handle:
            handle.write(previous_bytes)
    return work_dir, outputs_dir, input_path, previous_path, payload, settings


def build_template(settings: dict) -> tuple[str, bytes]:
    payload = _normalize_payload(settings.get("relation_payload") or {})
    mode = settings.get("mode") or "forward"
    if mode not in ("forward", "reverse"):
        raise ServiceError("模式只能是正向或逆向")
    try:
        student_count = int(settings.get("student_count") or 0)
    except (TypeError, ValueError):
        student_count = 0
    if student_count < 1 or student_count > 500:
        raise ServiceError("学生人数需要在 1 到 500 之间")

    work_dir = tempfile.mkdtemp(prefix="calculatorpro-tpl-")
    try:
        with override_outputs_dir(os.path.join(work_dir, "outputs")):
            relation_path = os.path.join(work_dir, "relation_table.json")
            with open(relation_path, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False)
            if mode == "forward":
                output = create_forward_template(work_dir, student_count, relation_path)
                filename = "正向成绩模板.xlsx"
            else:
                output = create_reverse_template(work_dir, student_count, relation_path)
                filename = "逆向成绩模板.xlsx"
            with open(output, "rb") as handle:
                return filename, handle.read()
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def run_calculation(excel_bytes: bytes, previous_bytes: bytes | None, settings: dict, captured: list | None = None) -> dict:
    work_dir, outputs_dir, input_path, previous_path, payload, settings = _prepare_workspace(
        excel_bytes, previous_bytes, settings
    )
    try:
        with override_outputs_dir(outputs_dir):
            student_count = _count_students(input_path, settings.get("mode") or "forward")
            processor = _build_processor(settings, input_path, payload)
            _load_previous(processor, previous_path)
            _write_supporting_docs(settings, payload)
            average, achievement = _run_grades(processor, settings)
            if captured is not None:
                captured.extend(_capture_outputs(outputs_dir))
            return {
                "mode": settings.get("mode") or "forward",
                "course_name": _course_name(settings),
                "student_count": student_count,
                "average_score": average,
                "achievement": achievement,
                "files": _output_names(outputs_dir),
            }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def run_export(excel_bytes: bytes, previous_bytes: bytes | None, settings: dict, captured: list | None = None) -> tuple[str, bytes, dict]:
    work_dir, outputs_dir, input_path, previous_path, payload, settings = _prepare_workspace(
        excel_bytes, previous_bytes, settings
    )
    try:
        with override_outputs_dir(outputs_dir):
            student_count = _count_students(input_path, settings.get("mode") or "forward")
            processor = _build_processor(settings, input_path, payload)
            _load_previous(processor, previous_path)
            _write_supporting_docs(settings, payload)
            average, achievement = _run_grades(processor, settings)
            summary = {
                "mode": settings.get("mode") or "forward",
                "course_name": _course_name(settings),
                "student_count": student_count,
                "average_score": average,
                "achievement": achievement,
                "files": _output_names(outputs_dir),
            }
            if captured is not None:
                captured.extend(_capture_outputs(outputs_dir))
            filename = f"{_course_name(settings)}统计表.zip"
            return filename, _zip_outputs(outputs_dir), summary
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def run_ai_report(
    excel_bytes: bytes, previous_bytes: bytes | None, settings: dict, captured: list | None = None
) -> tuple[str, bytes]:
    api_key = deepseek_api_key()
    if not api_key:
        raise ServiceError(AI_DISABLED_MESSAGE)
    work_dir, outputs_dir, input_path, previous_path, payload, settings = _prepare_workspace(
        excel_bytes, previous_bytes, settings
    )
    try:
        with override_outputs_dir(outputs_dir):
            processor = _build_processor(settings, input_path, payload)
            _load_previous(processor, previous_path)
            _write_supporting_docs(settings, payload)
            _run_grades(processor, settings)
            processor.store_api_key(api_key)
            count = _num_objectives(payload, settings)
            try:
                word_limit = int(settings.get("word_limit") or 200)
            except (TypeError, ValueError):
                word_limit = 200
            if word_limit < 1:
                word_limit = 200
            style = str(settings.get("report_style") or "专业")
            answers = generate_answers(processor, count, style, word_limit)
            processor.generate_improvement_report(answers=answers, output_dir=outputs_dir)
            template_path = get_resource_path("report_template.docx")
            if not os.path.exists(template_path):
                raise ServiceError("未找到报告模板 report_template.docx")
            basic = dict(settings.get("course_basic_info") or {})
            if not basic.get("course_name"):
                basic["course_name"] = _course_name(settings)
            ReportBuilder(template_path, outputs_dir).build(
                _open_info_for_report(settings),
                basic,
                {},
            )
            if captured is not None:
                captured.extend(_capture_outputs(outputs_dir))
            filename = f"{_course_name(settings)}AI分析报告.zip"
            return filename, _zip_outputs(outputs_dir)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
