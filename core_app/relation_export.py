"""导出课程考核与课程目标对应关系表（Word / JSON）。

这段逻辑从 relation_table.py 抽出，避免网页路径导入 PyQt。
桌面端仍通过 relation_table 调用同一实现。
"""
import json
from typing import List

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def _format_percent(value: float) -> str:
    text = f"{value:.2f}".rstrip("0").rstrip(".")
    return f"{text}%"


def _set_cell_shading(cell, fill: str):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shading)


def _set_paragraph_center(cell):
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

def _set_cell_border(cell, **kwargs):
    """
    设置单元格边框
    用法:
    _set_cell_border(cell, top={"sz": 12, "val": "single", "color": "#FF0000"}, bottom={...})
    如果不传参数，默认给四周加上黑色细实线
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # 检查 tcBorders 标签，没有就创建
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # 默认样式 (0.5磅黑色实线)
    default_border = {"sz": "4", "val": "single", "color": "auto", "space": "0"}

    # 遍历上下左右四个边
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        
        # 如果调用时没传具体参数，就使用默认样式
        if edge_data is None and not kwargs:
            edge_data = default_border
        elif edge_data is None:
            continue # 如果传了参数但没传这个边的，就跳过（保持原样）

        # 查找并删除旧的边框设置
        tag = 'w:{}'.format(edge)
        element = tcBorders.find(qn(tag))
        if element is not None:
            tcBorders.remove(element)

        # 创建新的边框
        new_element = OxmlElement(tag)
        for key in ["sz", "val", "color", "space", "shadow"]:
            if key in edge_data:
                new_element.set(qn('w:{}'.format(key)), str(edge_data[key]))
            elif key in default_border: # 补全默认值
                new_element.set(qn('w:{}'.format(key)), str(default_border[key]))
                
        tcBorders.append(new_element)


from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_table_borders(table):
    """
    【核弹级】强制给整个表格加上边框（包括内部和外部）
    这比给每个单元格加边框更稳定，不会因为合并单元格而断线。
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    
    # 确保 tblPr 存在
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    
    # 查找或创建 tblBorders 节点
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)
    
    # 定义我们要设置的6条边：上下左右 + 内部横线 + 内部竖线
    borders = {
        'top': {"val": "single", "sz": "4", "color": "auto"},
        'bottom': {"val": "single", "sz": "4", "color": "auto"},
        'left': {"val": "single", "sz": "4", "color": "auto"},
        'right': {"val": "single", "sz": "4", "color": "auto"},
        'insideH': {"val": "single", "sz": "4", "color": "auto"}, # 内部横线
        'insideV': {"val": "single", "sz": "4", "color": "auto"}  # 内部竖线
    }
    
    # 循环应用每一个边框设置
    for border_name, attrs in borders.items():
        # 先删除旧的设置（防止冲突）
        existing = tblBorders.find(qn(f'w:{border_name}'))
        if existing is not None:
            tblBorders.remove(existing)
            
        # 创建新的 XML 节点
        border = OxmlElement(f'w:{border_name}')
        for key, value in attrs.items():
            border.set(qn(f'w:{key}'), value)
        tblBorders.append(border)

def set_first_column_bold(table):
    """
    将表格的第一列所有文字设置为加粗
    """
    for row in table.rows:
        # 获取每一行的第1个单元格 (索引为0)
        cell = row.cells[0]
        
        # 遍历单元格里的所有段落和文字块，把它们变粗
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True

def export_relation_table(
    output_path: str,
    objectives_count: int,
    link_names: List[str],
    link_ratios: List[float],
    link_counts: List[int],
    methods_data: List[dict],
    obj_totals: List[float],
    total_sum: float,
):
    doc = Document()
    cols = 2 + objectives_count + 2
    method_rows = sum(1 if c <= 0 else c for c in link_counts)
    rows = 2 + method_rows + 1
    table = doc.add_table(rows=rows, cols=cols)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    total_width = 14.64
    col_width = total_width / cols
    for c in range(cols):
        for r in range(rows):
            table.cell(r, c).width = Cm(col_width)

    def _mark_header(row):
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        tbl_header = OxmlElement('w:tblHeader')
        tbl_header.set(qn('w:val'), 'true')
        trPr.append(tbl_header)

    def _set_cell_text(cell, text, bold=False):
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.font.name = '\u4eff\u5b8b'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '\u4eff\u5b8b')
        run.font.size = Pt(12)
        run.bold = bold
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Header row 1
    _set_cell_text(table.cell(0, 0), '\u8003\u6838\u73af\u8282', bold=True)
    _set_cell_text(table.cell(0, 1), '\u8003\u6838\u65b9\u5f0f', bold=True)
    _set_cell_text(table.cell(0, 2), '\u8bfe\u7a0b\u76ee\u6807\u5206\u6743\u91cd', bold=True)
    table.cell(0, 2).merge(table.cell(0, 1 + objectives_count))
    _set_cell_text(table.cell(0, 2 + objectives_count), '\u5c0f\u8ba1', bold=True)
    _set_cell_text(table.cell(0, 3 + objectives_count), '\u5408\u8ba1', bold=True)

    # Header row 2
    for i in range(objectives_count):
        _set_cell_text(table.cell(1, 2 + i), f"\u8bfe\u7a0b\u76ee\u6807{i+1}", bold=True)

    # Merge header vertical cells
    for col in [0, 1, 2 + objectives_count, 3 + objectives_count]:
        table.cell(0, col).merge(table.cell(1, col))

    _mark_header(table.rows[0])
    _mark_header(table.rows[1])

    # Fill rows
    data_row = 2
    method_idx = 0
    for link_idx, link_name in enumerate(link_names):
        rows_for_link = 1 if link_counts[link_idx] <= 0 else link_counts[link_idx]
        link_label = f"{link_name}\n{_format_percent(link_ratios[link_idx] * 100)}"
        _set_cell_text(table.cell(data_row, 0), link_label, bold=False)
        if rows_for_link > 1:
            table.cell(data_row, 0).merge(table.cell(data_row + rows_for_link - 1, 0))
            table.cell(data_row, 3 + objectives_count).merge(
                table.cell(data_row + rows_for_link - 1, 3 + objectives_count)
            )
        link_total = 0.0
        for _ in range(rows_for_link):
            row = data_row
            method = methods_data[method_idx]
            _set_cell_text(table.cell(row, 1), method["method_name"] or " ", bold=False)
            for obj_idx, value in enumerate(method["weights"]):
                _set_cell_text(table.cell(row, 2 + obj_idx), _format_percent(value), bold=False)
            _set_cell_text(table.cell(row, 2 + objectives_count), _format_percent(method["subtotal"]), bold=False)
            link_total += method["subtotal"]
            data_row += 1
            method_idx += 1
        _set_cell_text(
            table.cell(data_row - rows_for_link, 3 + objectives_count),
            _format_percent(link_total),
            bold=False,
        )

    # Total row
    total_row = rows - 1
    _set_cell_text(table.cell(total_row, 0), "100%", bold=True)
    _set_cell_text(table.cell(total_row, 1), "\u8bfe\u7a0b\u76ee\u6807\u603b\u6743\u91cd", bold=True)
    for obj_idx, value in enumerate(obj_totals):
        _set_cell_text(table.cell(total_row, 2 + obj_idx), _format_percent(value), bold=False)
    _set_cell_text(table.cell(total_row, 2 + objectives_count), _format_percent(total_sum), bold=False)
    _set_cell_text(table.cell(total_row, 3 + objectives_count), _format_percent(100.0), bold=False)

    # === 【修改这里】 ===
    # 不需要遍历单元格了，直接给表格下达“全局显示边框”的指令
    set_table_borders(table)
    # ==================
    # 2. 【新增】把第一列加粗
    set_first_column_bold(table)
    doc.save(output_path)

def export_relation_json(
    output_path: str,
    objectives_count: int,
    link_names: List[str],
    link_ratios: List[float],
    link_counts: List[int],
    methods_data: List[dict],
    obj_totals: List[float],
    total_sum: float,
):
    links = []
    for link_idx, link_name in enumerate(link_names):
        methods = []
        for method in methods_data:
            if method["link_idx"] != link_idx:
                continue
            supports = {}
            for obj_idx, value in enumerate(method["weights"]):
                supports[f"课程目标{obj_idx + 1}"] = round(value / 100.0, 6)
            methods.append(
                {
                    "name": method["method_name"],
                    "supports": supports,
                    "subtotal": round(method["subtotal"] / 100.0, 6),
                }
            )
        links.append(
            {
                "name": link_name,
                "ratio": round(link_ratios[link_idx], 6),
                "methods": methods,
            }
        )

    objectives = {}
    for obj_idx, total in enumerate(obj_totals):
        objectives[f"课程目标{obj_idx + 1}"] = round(total / 100.0, 6)

    payload = {
        "objectives_count": objectives_count,
        "links": links,
        "objectives_total_weights": objectives,
        "total_sum": round(total_sum / 100.0, 6),
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return payload


def write_relation_docx_from_payload(output_path: str, payload: dict) -> None:
    """根据网页/桌面共用的关系表 JSON 生成表4。"""
    objectives_count = int(payload.get("objectives_count") or 0)
    links = payload.get("links") or []
    link_names = []
    link_ratios = []
    link_counts = []
    methods_data = []

    for link_idx, link in enumerate(links):
        methods = link.get("methods") or []
        link_names.append((link.get("name") or "").strip())
        link_ratios.append(float(link.get("ratio") or 0))
        if not methods:
            link_counts.append(0)
            methods_data.append(
                {
                    "link_idx": link_idx,
                    "method_name": "无",
                    "weights": [0.0] * objectives_count,
                    "subtotal": 0.0,
                }
            )
            continue
        link_counts.append(len(methods))
        for method in methods:
            supports = method.get("supports") or {}
            weights = [
                float(supports.get(f"课程目标{i + 1}", 0) or 0) * 100
                for i in range(objectives_count)
            ]
            subtotal = method.get("subtotal")
            if subtotal is None:
                subtotal = sum(weights) / 100.0
            methods_data.append(
                {
                    "link_idx": link_idx,
                    "method_name": method.get("name") or "",
                    "weights": weights,
                    "subtotal": float(subtotal) * 100,
                }
            )

    stored = payload.get("objectives_total_weights") or {}
    obj_totals = []
    for obj_idx in range(objectives_count):
        key = f"课程目标{obj_idx + 1}"
        if key in stored:
            obj_totals.append(float(stored[key]) * 100)
            continue
        total = 0.0
        for link_idx, ratio in enumerate(link_ratios):
            for method in methods_data:
                if method["link_idx"] == link_idx and obj_idx < len(method["weights"]):
                    total += method["weights"][obj_idx] * ratio
        obj_totals.append(total)

    if payload.get("total_sum") is not None:
        total_sum = float(payload.get("total_sum") or 0) * 100
    else:
        total_sum = sum(obj_totals)

    export_relation_table(
        output_path,
        objectives_count,
        link_names,
        link_ratios,
        link_counts,
        methods_data,
        obj_totals,
        total_sum,
    )
