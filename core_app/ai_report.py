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
from utils import normalize_score, get_grade_level, calculate_final_score, calculate_achievement_level, adjust_column_widths
import time
import random
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH

class AIReportMixin:
        def test_deepseek_api(self, api_key: str) -> str:
            """测试 DeepSeek API 连接"""
            url = "https://api.deepseek.com/v1/chat/completions"
            api_key = api_key.strip().strip('<').strip('>')
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": "deepseek-chat",
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

            url = "https://api.deepseek.com/v1/chat/completions"
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
                "model": "deepseek-chat",
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
                self.previous_achievement_data = _init_defaults(obj_count)
                return

            try:
                if not os.path.exists(file_path):
                    self.previous_achievement_data = _init_defaults(obj_count)
                    if self.status_label:
                        self.status_label.setText("\u672a\u627e\u5230\u4e0a\u4e00\u5b66\u5e74\u8fbe\u6210\u5ea6\u8868\uff0c\u5df2\u4f7f\u7528\u9ed8\u8ba4\u503c")
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

        def generate_improvement_report(self, current_achievement: Dict[str, float], course_name: str, num_objectives: int, answers=None) -> None:
            """Generate single-column improvement report."""
            output_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "outputs")
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(
                output_dir,
                f"{self._safe_filename(course_name)}\u8bfe\u7a0b\u5206\u76ee\u6807\u8fbe\u6210\u60c5\u51b5\u5206\u6790\u3001\u5b58\u5728\u95ee\u9898\u53ca\u6539\u8fdb\u63aa\u65bd.xlsx",
            )

            num_objectives = max(1, int(num_objectives))
            questions = []
            for i in range(1, num_objectives + 1):
                questions.append(f"\u8bfe\u7a0b\u76ee\u6807{i}\uff081\uff09\u8fbe\u6210\u60c5\u51b5\u5206\u6790")
                questions.append(f"\u8bfe\u7a0b\u76ee\u6807{i}\uff082\uff09\u5b58\u5728\u95ee\u9898\u53ca\u6539\u8fdb\u63aa\u65bd")

            if answers is None:
                prev_data = self.previous_achievement_data or {}
                current_data = current_achievement or {}
                prev_total = prev_data.get("\u8bfe\u7a0b\u603b\u76ee\u6807", 0)
                current_total = current_data.get("\u603b\u8fbe\u6210\u5ea6", 0)

                context = f"\u8bfe\u7a0b\u7b80\u4ecb: {self.course_description}\n"
                for i, req in enumerate(self.objective_requirements, 1):
                    context += f"\u8bfe\u7a0b\u76ee\u6807{i}\u8981\u6c42: {req}\n"
                for i in range(1, num_objectives + 1):
                    prev_score = prev_data.get(f"\u8bfe\u7a0b\u76ee\u6807{i}", 0)
                    current_score = current_data.get(f"\u8bfe\u7a0b\u76ee\u6807{i}", 0)
                    context += f"\u8bfe\u7a0b\u76ee\u6807{i}\u4e0a\u4e00\u5b66\u5e74\u8fbe\u6210\u5ea6: {prev_score}\n"
                    context += f"\u8bfe\u7a0b\u76ee\u6807{i}\u672c\u5b66\u5e74\u8fbe\u6210\u5ea6: {current_score}\n"
                context += f"\u8bfe\u7a0b\u603b\u76ee\u6807\u4e0a\u4e00\u5b66\u5e74\u8fbe\u6210\u5ea6: {prev_total}\n"
                context += f"\u8bfe\u7a0b\u603b\u76ee\u6807\u672c\u5b66\u5e74\u8fbe\u6210\u5ea6: {current_total}\n"

                cache_file = os.path.join(output_dir, "api_cache.json")
                cached_answers = {}
                if os.path.exists(cache_file):
                    try:
                        with open(cache_file, "r", encoding="utf-8") as f:
                            cached_answers = json.load(f)
                    except Exception:
                        cached_answers = {}

                answers = []
                total_questions = len(questions)
                for i, question in enumerate(questions):
                    if self.status_label:
                        self.status_label.setText(f"\u6b63\u5728\u5904\u7406\u7b2c {i + 1}/{total_questions} \u4e2a\u95ee\u9898...")
                    prompt = f"{context}\n\u95ee\u9898: {question}"
                    cache_key = f"{course_name}_{question}"
                    if cache_key in cached_answers:
                        answers.append(cached_answers[cache_key])
                    else:
                        answer = self.call_deepseek_api(prompt)
                        cached_answers[cache_key] = answer
                        answers.append(answer)
                        try:
                            with open(cache_file, "w", encoding="utf-8") as f:
                                json.dump(cached_answers, f, indent=4, ensure_ascii=False)
                        except Exception:
                            pass

            rows = []
            overall_answer = ""
            answer_idx = 0
            if answers and len(answers) == (num_objectives * 2 + 1):
                overall_answer = answers[0]
                answer_idx = 1
            rows.append("\uff08\u4e00\uff09\u603b\u4f53\u60c5\u51b5")
            rows.append(overall_answer or "")
            rows.append("\uff08\u4e8c\uff09\u8bfe\u7a0b\u5206\u76ee\u6807\u8fbe\u6210\u60c5\u51b5\u5206\u6790\u3001\u5b58\u5728\u95ee\u9898\u53ca\u6539\u8fdb\u63aa\u65bd")
            for i in range(1, num_objectives + 1):
                rows.append(f"{i}.\u8bfe\u7a0b\u76ee\u6807{i}")
                rows.append("\uff081\uff09\u8fbe\u6210\u60c5\u51b5\u5206\u6790\uff1a")
                rows.append(answers[answer_idx] if answers else "")
                answer_idx += 1
                rows.append("\uff082\uff09\u5b58\u5728\u95ee\u9898\u53ca\u6539\u8fdb\u63aa\u65bd\uff1a")
                rows.append(answers[answer_idx] if answers else "")
                answer_idx += 1

            heading_texts = {
                "\uff08\u4e00\uff09\u603b\u4f53\u60c5\u51b5",
                "\uff08\u4e8c\uff09\u8bfe\u7a0b\u5206\u76ee\u6807\u8fbe\u6210\u60c5\u51b5\u5206\u6790\u3001\u5b58\u5728\u95ee\u9898\u53ca\u6539\u8fdb\u63aa\u65bd",
            }
            heading_font = Font(name="\u4eff\u5b8b", size=16, bold=True)
            body_font = Font(name="\u4eff\u5b8b", size=16, bold=False)
            no_border = Border()

            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "\u8bfe\u7a0b\u5206\u76ee\u6807\u8fbe\u6210\u60c5\u51b5\u5206\u6790\u3001\u5b58\u5728\u95ee\u9898\u53ca\u6539\u8fdb\u63aa\u65bd"
                ws.column_dimensions["A"].width = 8
                ws.column_dimensions["B"].width = 67
                ws.sheet_view.showGridLines = False

                for row_idx, text in enumerate(rows, start=1):
                    cell = ws.cell(row=row_idx, column=2)
                    if text in heading_texts:
                        cell.value = text
                        cell.font = heading_font
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    else:
                        display_text = text if text else ""
                        cell.value = display_text
                        cell.font = body_font
                        cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                    cell.border = no_border
                wb.save(output_file)
            except Exception as e:
                print(f"Error writing to Excel: {str(e)}")
                raise

        def store_api_key(self, api_key: str) -> None:
            """\u5b58\u50a8 API Key\u3002"""
            self.api_key = api_key
            if self.status_label:
                self.status_label.setText("\u5df2\u5b58\u50a8API Key")

        def generate_ai_report(self, num_objectives: int, current_achievement: Dict[str, float]) -> None:
            """\u751f\u6210AI\u5206\u6790\u62a5\u544a\u3002"""
            if not self.api_key:
                raise ValueError("\u8bf7\u5148\u8bbe\u7f6eAPI Key")

            course_name = self.course_name_input.text()
            if not course_name:
                if self.status_label:
                    self.status_label.setText("\u8bf7\u5148\u8f93\u5165\u8bfe\u7a0b\u540d\u79f0")
                return

            try:
                self.generate_improvement_report(current_achievement, course_name, num_objectives)
                if self.status_label:
                    self.status_label.setText("AI\u5206\u6790\u62a5\u544a\u5df2\u751f\u6210")
            except Exception as e:
                if self.status_label:
                    self.status_label.setText("\u751f\u6210AI\u5206\u6790\u62a5\u544a\u5931\u8d25")
                raise ValueError(f"\u751f\u6210AI\u5206\u6790\u62a5\u544a\u5931\u8d25: {str(e)}")
