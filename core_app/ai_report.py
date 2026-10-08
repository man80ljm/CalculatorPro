import os
import numpy as np
import pandas as pd
import json
import re
import requests
from typing import List, Dict,Callable, Optional
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from apply_noise import GradeReverseEngine
from utils import normalize_score, get_grade_level, calculate_final_score, calculate_achievement_level, adjust_column_widths, get_outputs_dir
import time
import random
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path

def deepseek_model() -> str:
    return os.environ.get("DEEPSEEK_MODEL", "").strip() or "deepseek-flash"


def deepseek_base_url() -> str:
    return (os.environ.get("DEEPSEEK_BASE_URL", "").strip() or "https://api.deepseek.com").rstrip("/")


def deepseek_chat_url() -> str:
    return deepseek_base_url() + "/v1/chat/completions"


def log_deepseek_call(kind: str, model: str, usage: dict | None, elapsed_seconds: float, ok: bool) -> None:
    """可选的调用记录。只写次数、模型和 token，不写密钥、提示词或学生名单。"""
    path = os.environ.get("CALC_DEEPSEEK_LOG", "").strip()
    if not path:
        return
    record = {
        "kind": kind,
        "model": model,
        "ok": bool(ok),
        "elapsed_seconds": round(float(elapsed_seconds), 3),
        "prompt_tokens": (usage or {}).get("prompt_tokens"),
        "completion_tokens": (usage or {}).get("completion_tokens"),
        "total_tokens": (usage or {}).get("total_tokens"),
    }
    run = os.environ.get("CALC_DEEPSEEK_RUN", "").strip()
    if run:
        record["run"] = run
    try:
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")
    except OSError:
        return


def report_max_tokens(num_objectives: int, word_limit: int) -> int:
    """一次要回全部段落。段数是 1 + 2N，上限按段数 × 字数放宽，封顶 8000。"""
    try:
        segments = 1 + 2 * max(0, int(num_objectives))
    except (TypeError, ValueError):
        segments = 1
    try:
        limit = int(word_limit)
    except (TypeError, ValueError):
        limit = 200
    if limit < 1:
        limit = 200
    return min(8000, max(1500, segments * limit * 3))


def call_deepseek_once(api_key: str, messages: list, max_tokens: int, timeout: int = 120) -> str:
    """单次请求，不在这里重试。报告路径最多调用两次。网页任务如果已经占了账号槽，就在该账号的密钥上做故障转移。"""
    from web_app.deepseek_pool import bound_account, report_text

    account = bound_account()
    if account is not None:
        return report_text(account, messages, max_tokens, timeout)
    key = (api_key or "").strip().strip("<").strip(">")
    if not key:
        raise RuntimeError("请先设置API Key")
    payload = {
        "model": deepseek_model(),
        "messages": messages,
        "temperature": 0.7,
        "max_tokens": int(max_tokens),
        "stream": False,
    }
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    started = time.perf_counter()
    try:
        response = requests.post(deepseek_chat_url(), headers=headers, json=payload, timeout=timeout)
        response.raise_for_status()
        body = response.json()
        usage = body.get("usage") or {}
        log_deepseek_call("report", deepseek_model(), usage, time.perf_counter() - started, True)
        return body["choices"][0]["message"]["content"].strip()
    except Exception:
        log_deepseek_call("report", deepseek_model(), {}, time.perf_counter() - started, False)
        raise


def parse_report_json(text: str, num_objectives: int) -> list[str] | None:
    raw = (text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw, count=1, flags=re.IGNORECASE)
        raw = re.sub(r"\s*```$", "", raw)
    data = None
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find("{")
        end = raw.rfind("}")
        if start >= 0 and end > start:
            try:
                data = json.loads(raw[start : end + 1])
            except json.JSONDecodeError:
                data = None
    if not isinstance(data, dict):
        return None
    overall = data.get("overall")
    objectives = data.get("objectives")
    if not isinstance(overall, str) or not overall.strip():
        return None
    if not isinstance(objectives, list):
        return None
    try:
        count = max(0, int(num_objectives))
    except (TypeError, ValueError):
        count = 0
    answers = [overall.strip()]
    for index in range(count):
        item = objectives[index] if index < len(objectives) and isinstance(objectives[index], dict) else {}
        analysis = item.get("analysis") if isinstance(item, dict) else ""
        improvement = item.get("improvement") if isinstance(item, dict) else ""
        answers.append("" if analysis is None else str(analysis).strip())
        answers.append("" if improvement is None else str(improvement).strip())
    return answers


PREVIOUS_MISSING = "—"
_PREVIOUS_TOTAL_KEYS = ("课程目标达成值", "课程总目标", "课程总达成值", "total_value")


def previous_year_known(data) -> bool:
    """没有上一学年文件时是 None。读到文件后才是字典，里面的 0 是真实数值。"""
    return isinstance(data, dict)


def previous_year_cell(data, key: str, missing: str = PREVIOUS_MISSING):
    if not previous_year_known(data) or key not in data or data.get(key) is None:
        return missing
    return data.get(key)


def previous_year_total(data, missing: str = PREVIOUS_MISSING):
    if not previous_year_known(data):
        return missing
    for key in _PREVIOUS_TOTAL_KEYS:
        if key in data and data.get(key) is not None:
            return data.get(key)
    return missing


def _attainment_rng(processor):
    """跟成绩拆分用同一个随机源，这样期望值落在已经消耗过的序列后面。"""
    getter = getattr(processor, "_score_rng", None)
    if callable(getter):
        return getter()
    rng = getattr(processor, "score_rng", None)
    if rng is not None and hasattr(rng, "uniform"):
        return rng
    from apply_noise import ScoreRng

    return ScoreRng(None)


def compute_expected_attainment(processor, total_attainment, prev_data=None):
    """桌面版期望值，只在这里计算，并写到 processor.expected_attainment。

    有上一学年总达成值、且它和本学年总达成度都不为 0 时，在两者之间取一次均匀随机数（三位小数）。
    没有上一学年文件时不把 0 当成上一轮达成值，期望值等于本学年总达成度。
    报告提示词只读 processor.expected_attainment，不再抽一次随机数。
    """
    if prev_data is None:
        prev_data = getattr(processor, "previous_achievement_data", None)
    raw_prev_total = previous_year_total(prev_data)
    prev_total = raw_prev_total if isinstance(raw_prev_total, (int, float)) else 0
    if prev_total and total_attainment:
        low = min(prev_total, total_attainment)
        high = max(prev_total, total_attainment)
        expected = round(_attainment_rng(processor).uniform(low, high), 3)
    else:
        expected = total_attainment
    try:
        expected = float(expected)
    except (TypeError, ValueError):
        expected = 0.0
    processor.expected_attainment = expected
    return expected


def _expected_for_prompt(processor, total_attainment, prev_data):
    """成绩计算已经写入 expected_attainment 时直接用，避免再调用随机数。"""
    stored = getattr(processor, "expected_attainment", None)
    if isinstance(stored, (int, float)) and not isinstance(stored, bool):
        return float(stored)
    return compute_expected_attainment(processor, total_attainment, prev_data)


def _report_context(processor, num_objectives: int) -> tuple[str, bool]:
    prev_data = getattr(processor, "previous_achievement_data", None)
    known = previous_year_known(prev_data)
    current_data = getattr(processor, "current_achievement", None) or {}
    expected = _expected_for_prompt(processor, current_data.get("总达成度", 0), prev_data)
    expected_text = f"{float(expected):.3f}"
    context = f"课程简介: {getattr(processor, 'course_description', '') or ''}\n"
    for index, requirement in enumerate(getattr(processor, "objective_requirements", None) or [], 1):
        context += f"课程目标{index}要求: {requirement}\n"
    if known:
        for index in range(1, num_objectives + 1):
            context += f"课程目标{index}上一学年达成度: {prev_data.get(f'课程目标{index}', 0)}\n"
            context += f"课程目标{index}本学年达成度: {current_data.get(f'课程目标{index}', 0)}\n"
        context += f"课程目标达成值（本学年）: {current_data.get('总达成度', 0)}\n"
        context += f"课程目标达成期望值: {expected_text}\n"
        context += f"上一轮教学课程目标达成值: {prev_data.get('课程总目标', 0)}\n"
    else:
        context += "上一学年达成度：无数据\n"
        for index in range(1, num_objectives + 1):
            context += f"课程目标{index}本学年达成度: {current_data.get(f'课程目标{index}', 0)}\n"
        context += f"课程目标达成值（本学年）: {current_data.get('总达成度', 0)}\n"
        context += f"课程目标达成期望值: {expected_text}\n"
    return context, known


def generate_report_answers(processor, num_objectives: int, report_style: str, word_limit: int) -> list[str]:
    """一次请求返回全部段落。解析失败最多再请求一次，然后把原文当作总体情况。"""
    try:
        count = max(0, int(num_objectives))
    except (TypeError, ValueError):
        count = 0
    safe_limit = word_limit if isinstance(word_limit, int) and word_limit > 0 else 200
    try:
        safe_limit = int(word_limit)
    except (TypeError, ValueError):
        safe_limit = 200
    if safe_limit < 1:
        safe_limit = 200
    min_chars = max(50, int(safe_limit * 0.8))
    style = report_style or "专业"
    example = {
        "overall": "总体情况，一段话",
        "objectives": [
            {"analysis": "课程目标1达成情况分析", "improvement": "课程目标1存在问题及改进措施"}
        ],
    }
    context, known = _report_context(processor, count)
    compare_rule = ""
    if not known:
        compare_rule = "无上学年数据，不要做同比/环比比较，不要提及上一学年数值。\n"
    prompt = (
        f"{context}\n"
        f"{compare_rule}"
        f"请以{style}风格，用一段话写总体情况，再为每一个课程目标各写两段："
        f"达成情况分析、存在问题及改进措施。课程目标共 {count} 个，objectives 数组长度必须是 {count}。\n"
        f"每段尽量接近{safe_limit}字，不少于{min_chars}字。不分点，不要 Markdown，不要标题符号。\n"
        "不要出现学生姓名或学号。只返回 JSON，不要代码围栏。格式如下：\n"
        f"{json.dumps(example, ensure_ascii=False)}"
    )
    system = "你撰写课程目标达成情况分析。只输出 JSON 对象，键为 overall 和 objectives。不要输出学生姓名或学号。"
    if not known:
        system += "无上学年数据，不要做同比或环比，不要提及上一学年数值。"
    messages = [
        {"role": "system", "content": system},
        {"role": "user", "content": prompt},
    ]
    api_key = getattr(processor, "api_key", "") or ""
    last_text = ""
    for _attempt in range(2):
        try:
            last_text = call_deepseek_once(api_key, messages, report_max_tokens(count, safe_limit))
        except Exception as exc:
            last_text = f"API 调用失败: {exc}"
            continue
        parsed = parse_report_json(last_text, count)
        if parsed is not None:
            return parsed
    answers = [last_text.strip()]
    answers.extend([""] * (2 * count))
    return answers


class AIReportMixin:
        def _set_cell_border(self, cell, size=4, color='000000'):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = tcPr.find(qn('w:tcBorders'))
            if tcBorders is None:
                tcBorders = OxmlElement('w:tcBorders')
                tcPr.append(tcBorders)
            for edge in ('top', 'left', 'bottom', 'right'):
                elem = tcBorders.find(qn(f'w:{edge}'))
                if elem is None:
                    elem = OxmlElement(f'w:{edge}')
                    tcBorders.append(elem)
                if size:
                    elem.set(qn('w:val'), 'single')
                    elem.set(qn('w:sz'), str(size))
                    elem.set(qn('w:space'), '0')
                    elem.set(qn('w:color'), color)
                else:
                    elem.set(qn('w:val'), 'nil')


        def test_deepseek_api(self, api_key: str) -> str:
            """测试 DeepSeek API 连接"""
            url = deepseek_chat_url()
            api_key = api_key.strip().strip('<').strip('>')
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": deepseek_model(),
                "messages": [
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": "测试连接"}
                ],
                "temperature": 0.7,
                "top_p": 1,
                "max_tokens": 10,
                "stream": False
            }
            
            try:
                response = requests.post(url, headers=headers, json=payload, timeout=10)
                response.raise_for_status()
                return "连接成功"
            except requests.RequestException as e:
                error_message = f"连接失败: {str(e)}"
                if hasattr(e, 'response') and e.response is not None:
                    error_message += f"\n服务器返回: {e.response.text}"
                return error_message

        def call_deepseek_api(self, prompt: str) -> str:
            """调用 DeepSeek API 获取答案，包含重试与超时处理。"""
            if not self.api_key:
                return "请先设置API Key"

            url = deepseek_chat_url()
            api_key = self.api_key.strip().strip("<").strip(">")
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            }

            max_tokens = 600
            match = re.search(r"接近(\d+)字", prompt)
            if match:
                try:
                    word_limit = int(match.group(1))
                    max_tokens = max(200, min(1500, word_limit * 4))
                except ValueError:
                    max_tokens = 600

            payload = {
                "model": deepseek_model(),
                "messages": [
                    {"role": "system", "content": "You are a helpful assistant specializing in course analysis and improvement."},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.7,
                "top_p": 1,
                "max_tokens": max_tokens,
                "stream": False,
            }

            max_retries = 3
            for attempt in range(max_retries):
                try:
                    response = requests.post(url, headers=headers, json=payload, timeout=30)
                    response.raise_for_status()
                    return response.json()["choices"][0]["message"]["content"].strip()
                except requests.Timeout:
                    if attempt < max_retries - 1:
                        print(f"API 调用超时，正在重试（第 {attempt + 1}/{max_retries} 次）...")
                        time.sleep(2)
                        continue
                    return "API 调用超时，请检查网络连接或稍后重试（可能需要使用VPN或代理访问 api.deepseek.com）"
                except requests.RequestException as e:
                    error_message = f"API 调用失败: {str(e)}"
                    if hasattr(e, "response") and e.response is not None:
                        error_message += f"\n服务端返回: {e.response.text}"
                    if attempt < max_retries - 1:
                        print(f"API 调用失败，正在重试（第 {attempt + 1}/{max_retries} 次）...")
                        time.sleep(2)
                        continue
                    return error_message
                except (KeyError, IndexError):
                    return "API 返回格式错误，无法解析结果"

        def load_previous_achievement(self, file_path: str) -> None:
            """\u52a0\u8f7d\u4e0a\u4e00\u5b66\u5e74\u8fbe\u6210\u5ea6\u8868\uff0c\u63d0\u53d6\u5404\u8bfe\u7a0b\u76ee\u6807\u7684\u5206\u76ee\u6807\u8fbe\u6210\u503c\u4ee5\u53ca\u8bfe\u7a0b\u76ee\u6807\u8fbe\u6210\u503c\u3002"""
            def _objective_count():
                payload = self.relation_payload or {}
                objectives = payload.get("objectives") if isinstance(payload, dict) else None
                if objectives:
                    return len(objectives)
                if self.objective_requirements:
                    return len(self.objective_requirements)
                return 5

            def _init_defaults(n):
                data = {f'\u8bfe\u7a0b\u76ee\u6807{i}': 0 for i in range(1, n + 1)}
                data['\u8bfe\u7a0b\u603b\u76ee\u6807'] = 0
                return data

            obj_count = _objective_count()
            if not file_path:
                self.previous_achievement_data = None
                return

            try:
                if not os.path.exists(file_path):
                    self.previous_achievement_data = None
                    if self.status_label:
                        self.status_label.setText("未找到上一学年达成度表，不做同比")
                    return

                xls = pd.ExcelFile(file_path)
                df = None
                for sheet in xls.sheet_names:
                    tmp = pd.read_excel(file_path, sheet_name=sheet)
                    cols = [str(c).strip() for c in tmp.columns]
                    if "\u8bfe\u7a0b\u5206\u76ee\u6807" in cols or "\u8bfe\u7a0b\u76ee\u6807" in cols:
                        df = tmp
                        break
                if df is None:
                    df = pd.read_excel(file_path)

                cols = [str(c).strip() for c in df.columns]
                data = _init_defaults(obj_count)

                if "\u8bfe\u7a0b\u5206\u76ee\u6807" in cols and "\u5206\u76ee\u6807\u8fbe\u6210\u503c" in cols:
                    for i in range(1, obj_count + 1):
                        key = f'\u8bfe\u7a0b\u76ee\u6807{i}'
                        rows = df[df["\u8bfe\u7a0b\u5206\u76ee\u6807"].astype(str).str.strip() == key]
                        if not rows.empty:
                            val = rows["\u5206\u76ee\u6807\u8fbe\u6210\u503c"].dropna().tolist()
                            if val:
                                data[key] = float(val[0])
                    total_rows = df[df["\u8bfe\u7a0b\u5206\u76ee\u6807"].astype(str).str.strip().isin([
                        "\u8bfe\u7a0b\u76ee\u6807\u8fbe\u6210\u503c",
                        "\u8bfe\u7a0b\u603b\u76ee\u6807\u8fbe\u6210\u503c",
                        "\u8bfe\u7a0b\u603b\u8fbe\u6210\u503c",
                    ])]
                    if not total_rows.empty:
                        val = total_rows["\u5206\u76ee\u6807\u8fbe\u6210\u503c"].dropna().tolist()
                        if val:
                            data["\u8bfe\u7a0b\u603b\u76ee\u6807"] = float(val[0])
                elif "\u8bfe\u7a0b\u76ee\u6807" in cols:
                    value_col = None
                    for cand in [
                        "\u4e0a\u4e00\u5e74\u5ea6\u8fbe\u6210\u5ea6",
                        "\u4e0a\u4e00\u8f6e\u6559\u5b66\u5206\u76ee\u6807\u8fbe\u6210\u503c",
                        "\u5206\u76ee\u6807\u8fbe\u6210\u503c",
                    ]:
                        if cand in cols:
                            value_col = cand
                            break
                    if value_col:
                        for _, row in df.iterrows():
                            target = str(row["\u8bfe\u7a0b\u76ee\u6807"]).strip()
                            if target in data and isinstance(row[value_col], (int, float)):
                                data[target] = float(row[value_col])
                        total_rows = df[df["\u8bfe\u7a0b\u76ee\u6807"].astype(str).str.strip().isin([
                            "\u8bfe\u7a0b\u76ee\u6807\u8fbe\u6210\u503c",
                            "\u8bfe\u7a0b\u603b\u76ee\u6807\u8fbe\u6210\u503c",
                            "\u8bfe\u7a0b\u603b\u8fbe\u6210\u503c",
                        ])]
                        if not total_rows.empty and isinstance(total_rows[value_col].iloc[0], (int, float)):
                            data["\u8bfe\u7a0b\u603b\u76ee\u6807"] = float(total_rows[value_col].iloc[0])
                self.previous_achievement_data = data
                if self.status_label:
                    self.status_label.setText(f"\u5df2\u52a0\u8f7d\u4e0a\u4e00\u5b66\u5e74\u8fbe\u6210\u5ea6\u8868: {os.path.basename(file_path)}")
            except Exception as e:
                if self.status_label:
                    self.status_label.setText("\u52a0\u8f7d\u4e0a\u4e00\u5b66\u5e74\u8fbe\u6210\u5ea6\u8868\u5931\u8d25")
                raise ValueError(f"\u52a0\u8f7d\u4e0a\u4e00\u5b66\u5e74\u8fbe\u6210\u5ea6\u8868\u5931\u8d25: {str(e)}")

        def generate_improvement_report(self, answers: list[str] | None, output_dir: str | None = None) -> str:
            base_dir = Path(output_dir) if output_dir else Path(get_outputs_dir())
            base_dir.mkdir(parents=True, exist_ok=True)

            course_name = getattr(self, 'course_name', '')
            if not course_name and hasattr(self, 'course_name_input'):
                try:
                    course_name = self.course_name_input.text().strip()
                except Exception:
                    course_name = ''
            course_name = course_name or '课程'
            safe_name = re.sub(r'[\/:*?"<>|]', '_', course_name)
            output_file = base_dir / f"6.课程目标达成情况分析、存在问题及改进措施表.docx"

            # 删除旧版 xlsx 兼容文件
            old_xlsx = base_dir / f"{safe_name}课程分目标达成情况分析、存在问题及改进措施.xlsx"
            if old_xlsx.exists():
                old_xlsx.unlink(missing_ok=True)

            obj_count = len(self.objective_requirements or [])
            total_questions = 1 + obj_count * 2
            answers = answers or []
            while len(answers) < total_questions:
                answers.append('')

            overall_answer = answers[0].strip() if answers else ''

            rows: list[tuple[str, str]] = []
            rows.append(('（一）总体情况', 'heading'))
            rows.append((overall_answer, 'answer'))
            rows.append(('（二）课程分目标达成情况分析、存在问题及改进措施', 'heading'))

            idx = 1
            for i in range(1, obj_count + 1):
                rows.append((f"{i}. 课程目标{i}", 'fixed'))
                rows.append(('（1）达成情况分析：', 'fixed'))
                rows.append((answers[idx].strip() if idx < len(answers) else '', 'answer'))
                idx += 1
                rows.append(('（2）存在问题及改进措施：', 'fixed'))
                rows.append((answers[idx].strip() if idx < len(answers) else '', 'answer'))
                idx += 1

            doc = Document()
            table = doc.add_table(rows=len(rows), cols=2)
            table.autofit = False
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            total_width_cm = 14.64
            left_col_cm = 1.0
            right_col_cm = total_width_cm - left_col_cm
            table.columns[0].width = Cm(left_col_cm)
            table.columns[1].width = Cm(right_col_cm)

            fixed_size = Pt(15)  # 标题字体
            answer_size = Pt(14)  # 正文字体

            for r_idx, (text, kind) in enumerate(rows):
                row = table.rows[r_idx]
                row.cells[0].text = ''
                p = row.cells[1].paragraphs[0]
                p.text = ''
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pf = p.paragraph_format
                pf.first_line_indent = None
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)

                run = p.add_run(text)
                run.font.name = '仿宋'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
                if kind == 'answer':
                    run.font.size = answer_size
                    run.bold = False
                elif kind == 'heading':
                    run.font.size = fixed_size
                    run.bold = True
                else:
                    run.font.size = fixed_size
                    run.bold = False

                row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP
                row.cells[1].vertical_alignment = WD_ALIGN_VERTICAL.TOP

            for row in table.rows:
                for cell in row.cells:
                    self._set_cell_border(cell, size=0)

            doc.save(output_file)
            return str(output_file)

        def store_api_key(self, api_key: str) -> None:
            """存储 API Key"""
            self.api_key = api_key
            if hasattr(self, 'status_label') and self.status_label:
                self.status_label.setText("已存储API Key")

        def generate_ai_report(self, *args, **kwargs) -> None:
            """适配旧接口调用的 generate_improvement_report 包装器"""
            answers = kwargs.get('answers') if isinstance(kwargs, dict) else None
            self.generate_improvement_report(answers=answers)
