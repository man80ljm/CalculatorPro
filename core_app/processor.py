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

from core_app.ai_report import AIReportMixin
from core_app.word_exports import WordExportMixin
from core_app.excel_calc import ExcelCalcMixin

class GradeProcessor(AIReportMixin, WordExportMixin, ExcelCalcMixin):
        def __init__(self, course_name_input, num_objectives_input, weight_inputs, usual_ratio_input,
                     midterm_ratio_input, final_ratio_input, status_label, input_file,
                     course_description="", objective_requirements=None, relation_payload=None):
            self.course_name_input = course_name_input
            self.num_objectives_input = num_objectives_input
            self.weight_inputs = weight_inputs
            self.usual_ratio_input = usual_ratio_input
            self.midterm_ratio_input = midterm_ratio_input
            self.final_ratio_input = final_ratio_input
            self.status_label = status_label
            self.input_file = input_file
            self.course_description = course_description
            self.objective_requirements = objective_requirements or []
            self.previous_achievement_data = None
            self.api_key = None
            self.relation_payload = relation_payload or {}
            self.noise_config = None
            self.reverse_engine = GradeReverseEngine()

        def set_noise_config(self, config: dict):
            """\u8bbe\u7f6e\u566a\u58f0\u914d\u7f6e"""
            self.noise_config = config or None

        def set_relation_payload(self, payload: dict):
            """\u8bbe\u7f6e\u8bfe\u7a0b\u8003\u6838\u4e0e\u76ee\u6807\u5bf9\u5e94\u5173\u7cfb"""
            self.relation_payload = payload or {}

        def _safe_filename(self, name: str) -> str:
            """Sanitize filename for Windows paths."""
            if not name:
                return "\u672a\u547d\u540d"
            safe = re.sub(r'[\\/:*"<>|?\\r\\n\\t]', "_", str(name)).strip()
            return safe or "\u672a\u547d\u540d"
