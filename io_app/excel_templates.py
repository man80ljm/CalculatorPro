import json
import os
from typing import Dict, List, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Protection


def _ensure_outputs_dir(base_dir: str) -> str:
    outputs_dir = os.path.join(base_dir, "outputs")
    os.makedirs(outputs_dir, exist_ok=True)
    return outputs_dir


def _load_relation_json(path: str) -> Dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def create_reverse_template(base_dir: str, student_count: int) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "逆向成绩模板"

    headers = ["姓名", "平时考核", "期中考核", "期末考核"]
    ws.append(headers)
    _apply_header_style(ws, 1, len(headers))
    _append_blank_rows(ws, student_count, len(headers))
    _protect_sheet(ws, editable_cols=list(range(1, len(headers) + 1)))

    output_path = os.path.join(_ensure_outputs_dir(base_dir), "逆向成绩模板.xlsx")
    wb.save(output_path)
    return output_path


def create_forward_template(base_dir: str, student_count: int, relation_json_path: str) -> str:
    data = _load_relation_json(relation_json_path)
    links = data.get("links", [])

    wb = Workbook()
    ws = wb.active
    ws.title = "正向成绩模板"

    # Row 1: link names, Row 2: method names
    ws.cell(row=1, column=1, value="姓名")
    ws.cell(row=2, column=1, value="姓名")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    col = 2
    method_columns = []
    for link in links:
        methods = link.get("methods", [])
        if not methods:
            methods = [{"name": "无"}]
        start_col = col
        for method in methods:
            ws.cell(row=2, column=col, value=method.get("name", ""))
            method_columns.append(col)
            col += 1
        end_col = col - 1
        ws.cell(row=1, column=start_col, value=link.get("name", ""))
        if end_col >= start_col:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)

    _apply_header_style(ws, 1, col - 1)
    _apply_header_style(ws, 2, col - 1)

    _append_blank_rows(ws, student_count, col - 1, start_row=3)
    _protect_sheet(ws, editable_cols=list(range(1, col)))

    output_path = os.path.join(_ensure_outputs_dir(base_dir), "正向成绩模板.xlsx")
    wb.save(output_path)
    return output_path


def _append_blank_rows(ws, student_count: int, col_count: int, start_row: int = 2):
    for i in range(student_count):
        row_idx = start_row + i
        for col in range(1, col_count + 1):
            ws.cell(row=row_idx, column=col, value="")


def _apply_header_style(ws, row: int, col_count: int):
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _protect_sheet(ws, editable_cols: List[int]):
    ws.protection.sheet = True
    ws.protection.enable()
    max_row = ws.max_row
    for row in range(1, max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if row <= 2:
                cell.protection = Protection(locked=True)
                continue
            cell.protection = Protection(locked=(col not in editable_cols))
