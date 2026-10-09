"""把桌面端的计算、导出和 AI 报告跑在每次请求独立的临时目录里。"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
import zipfile
from io import BytesIO
from typing import Any
from pathlib import Path

import pandas as pd

from core import GradeProcessor
from core_app.course_docs import export_course_basic_docx, export_grad_req_docx
from core_app.relation_export import write_relation_docx_from_payload
from core_app.report_builder import ReportBuilder
from io_app.excel_templates import create_forward_template, create_reverse_template
from utils import get_resource_path, override_outputs_dir
from web_app.ai_runner import generate_answers
from web_app.download_names import settings_archive_filename
from web_app.resource_limits import limited_computation

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
    def __init__(self, message: str, status: int = 400, **extra):
        super().__init__(message)
        self.status = status
        self.extra = extra


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
    from web_app.deepseek_pool import PoolError, primary_key

    try:
        return primary_key()
    except PoolError:
        return ""


def ai_status() -> dict:
    from web_app.deepseek_pool import PoolError, configured

    try:
        enabled = configured()
    except PoolError as exc:
        return {"enabled": False, "message": str(exc)}
    if enabled:
        from core_app.ai_report import deepseek_model

        return {"enabled": True, "message": f"AI 分析报告可用（DeepSeek / {deepseek_model()}）。密钥只保存在服务器。"}
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
    if len(links) > 30 or sum(len(link["methods"]) for link in links if isinstance(link, dict) and isinstance(link.get("methods"), list)) > 150:
        raise ServiceError("考核项目过多，请将考核环节控制在 30 个、考核方式控制在 150 个以内。")
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
    if objectives_count > 30:
        raise ServiceError("课程目标数量过多，最多支持 30 个。")
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


def attainment_seed(grade_path: str, settings: dict) -> int | None:
    """学期、成绩内容和拆分设置决定种子。没有学期 id 时返回 None，桌面版保持原来的随机。"""
    term_id = settings.get("term_id")
    if term_id in (None, ""):
        return None
    digest = hashlib.sha256()
    digest.update(str(int(term_id)).encode())
    with open(grade_path, "rb") as handle:
        digest.update(handle.read())
    material = {
        "relation_payload": settings.get("relation_payload"),
        "spread_mode": settings.get("spread_mode"),
        "distribution": settings.get("distribution"),
        "noise_config": settings.get("noise_config"),
    }
    digest.update(json.dumps(material, ensure_ascii=False, sort_keys=True, default=str).encode())
    return int.from_bytes(digest.digest()[:4], "big")


def _bind_score_seed(processor: GradeProcessor, settings: dict) -> None:
    seed = attainment_seed(processor.input_file, settings)
    if seed is not None:
        processor.set_score_seed(seed)


def _run_grades(processor: GradeProcessor, settings: dict):
    _bind_score_seed(processor, settings)
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
    from web_app.resource_limits import validate_document
    validate_document(excel_bytes, "grade.xlsx")
    if previous_bytes:
        validate_document(previous_bytes, "previous.xlsx")
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


@limited_computation
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


@limited_computation
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
            from web_app.grade_register import headcount_warnings

            basic = settings.get("course_basic_info") if isinstance(settings.get("course_basic_info"), dict) else {}
            warnings = headcount_warnings(
                settings.get("student_count"), basic.get("exam_count"), student_count, student_count, source="成绩文件"
            )
            return {
                "mode": settings.get("mode") or "forward",
                "course_name": _course_name(settings),
                "student_count": student_count,
                "average_score": average,
                "achievement": achievement,
                "files": _output_names(outputs_dir),
                "warnings": warnings,
            }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@limited_computation
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
            filename = settings_archive_filename(_course_name(settings), settings)
            return filename, _zip_outputs(outputs_dir), summary
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


class ReportBundle:
    """一次成绩处理产生的统计表、分析稿和压缩包。"""

    def __init__(self, filename: str, content: bytes, summary: dict, files: list[tuple[str, bytes]]):
        self.filename = filename
        self.content = content
        self.summary = summary
        self.files = files


def _ai_answer_error(answers) -> str:
    if not answers:
        return "AI 没有返回分析内容"
    head = str(answers[0] or "").strip()
    if head.startswith("API 调用失败"):
        return head
    return ""


def run_report_pipeline(
    excel_bytes: bytes, previous_bytes: bytes | None, settings: dict, on_stage=None, checkpoint_dir=None
) -> ReportBundle:
    """同一个临时目录里只跑一次成绩处理，再写分析稿并打包。on_stage 在进入每个阶段时调用。"""

    from web_app.resource_limits import computation_slot
    compute = None
    def enter(stage: str) -> None:
        nonlocal compute
        if stage == "ai" and compute is not None:
            compute.__exit__(None, None, None)
            compute = None
        elif stage != "ai" and compute is None:
            compute = computation_slot()
            compute.__enter__()
        if on_stage:
            on_stage(stage)

    work_dir = None
    try:
        enter("calculate")
        work_dir, outputs_dir, input_path, previous_path, payload, settings = _prepare_workspace(
            excel_bytes, previous_bytes, settings
        )
        with override_outputs_dir(outputs_dir):
            processor = _build_processor(settings, input_path, payload)
            _load_previous(processor, previous_path)
            processor.on_table_stage = lambda: enter("tables")
            student_count = _count_students(input_path, settings.get("mode") or "forward")
            average, achievement = _run_grades(processor, settings)
            if not getattr(processor, "_table_stage_entered", False):
                enter("tables")
            _write_supporting_docs(settings, payload)
            enter("ai")
            from web_app.deepseek_pool import PoolError, bound_account, configured, primary_key, redact_text, reserve

            cache = Path(checkpoint_dir) / "answers.json" if checkpoint_dir else None
            answers = None
            if cache and cache.is_file():
                try:
                    saved = json.loads(cache.read_text(encoding="utf-8"))
                    if isinstance(saved, list) and saved and all(isinstance(item, str) for item in saved):
                        answers = saved
                except (OSError, ValueError):
                    pass
            if answers is None:
                try:
                    ready = configured()
                except PoolError as exc:
                    raise ServiceError(str(exc)) from None
                if not ready:
                    raise ServiceError(AI_DISABLED_MESSAGE)
            extra = None
            if answers is None and bound_account() is None:
                extra = reserve("report").wait()
                extra.bind()
            try:
                processor.store_api_key(primary_key())
                count = _num_objectives(payload, settings)
                try:
                    word_limit = int(settings.get("word_limit") or 200)
                except (TypeError, ValueError):
                    word_limit = 200
                if word_limit < 1:
                    word_limit = 200
                style = str(settings.get("report_style") or "专业")
                if answers is None:
                    answers = generate_answers(processor, count, style, word_limit)
                failure = _ai_answer_error(answers)
                if failure:
                    raise ServiceError(redact_text(f"AI 分析撰写失败：{failure}"))
                if cache and not cache.is_file():
                    temporary = cache.with_suffix(".tmp")
                    with temporary.open("w", encoding="utf-8") as handle:
                        json.dump(answers, handle, ensure_ascii=False)
                        handle.flush()
                        os.fsync(handle.fileno())
                    os.replace(temporary, cache)
            finally:
                if extra is not None:
                    extra.unbind()
                    extra.release()
            enter("package")
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
            from web_app.grade_register import headcount_warnings

            warnings = headcount_warnings(
                settings.get("student_count"),
                basic.get("exam_count"),
                student_count,
                student_count,
                source="成绩文件",
            )
            files = _capture_outputs(outputs_dir)
            filename = settings_archive_filename(_course_name(settings), settings)
            content = _zip_outputs(outputs_dir)
            summary = {
                "mode": settings.get("mode") or "forward",
                "course_name": _course_name(settings),
                "student_count": student_count,
                "average_score": average,
                "achievement": achievement,
                "files": _output_names(outputs_dir),
                "warnings": warnings,
            }
            return ReportBundle(filename, content, summary, files)
    finally:
        if compute is not None:
            compute.__exit__(None, None, None)
        if work_dir:
            shutil.rmtree(work_dir, ignore_errors=True)


def run_ai_report(
    excel_bytes: bytes, previous_bytes: bytes | None, settings: dict, captured: list | None = None
) -> tuple[str, bytes]:
    from web_app.deepseek_pool import PoolError, configured, reserve

    try:
        ready = configured()
    except PoolError as exc:
        raise ServiceError(str(exc)) from None
    if not ready:
        raise ServiceError(AI_DISABLED_MESSAGE)
    with reserve("report").wait():
        bundle = run_report_pipeline(excel_bytes, previous_bytes, settings)
    if captured is not None:
        captured.extend(bundle.files)
    return bundle.filename, bundle.content
