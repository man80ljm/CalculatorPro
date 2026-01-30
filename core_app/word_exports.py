import os
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH


class WordExportMixin:
    def _set_docx_cell_margins(self, cell, top=0, bottom=0, left=0, right=0):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcMar = tcPr.find(qn('w:tcMar'))
        if tcMar is None:
            tcMar = OxmlElement('w:tcMar')
            tcPr.append(tcMar)
        for tag, val in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
            node = tcMar.find(qn(f'w:{tag}'))
            if node is None:
                node = OxmlElement(f'w:{tag}')
                tcMar.append(node)
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')

    def _set_docx_cell_border(self, cell, size=4):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        borders = tcPr.find(qn('w:tcBorders'))
        if borders is None:
            borders = OxmlElement('w:tcBorders')
            tcPr.append(borders)
        for edge in ('top', 'left', 'bottom', 'right'):
            edge_tag = qn(f'w:{edge}')
            element = borders.find(edge_tag)
            if element is None:
                element = OxmlElement(f'w:{edge}')
                borders.append(element)
            element.set(qn('w:val'), 'single')
            element.set(qn('w:sz'), str(size))
            element.set(qn('w:color'), '000000')

    def _export_stats_docx(self, composition_text, max_score, min_score, avg_score, counts, ratios):
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        output_dir = os.path.join(root, 'outputs')
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, '2.课程成绩统计表.docx')

        doc = Document()
        table = doc.add_table(rows=5, cols=6)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        total_cm = 14.64
        first_cm = 3.75
        other_cm = (total_cm - first_cm) / 5

        data = [
            ['成绩构成', composition_text.strip(), '', '', '', ''],
            ['最高成绩', max_score, '最低成绩', min_score, '平均成绩', avg_score],
            ['成绩等级', '90-100\n(优秀)', '80-89\n(良好)', '70-79\n(中等)', '60-69\n(及格)', '<60\n(不及格)'],
            ['人数'] + list(counts),
            ['占考核人数的比例'] + [f"{r*100:.2f}%" for r in ratios],
        ]

        bold_coords = {
            (0, 0),
            (1, 0), (1, 2), (1, 4),
            (2, 0), (2, 1), (2, 2), (2, 3), (2, 4), (2, 5),
            (3, 0),
            (4, 0),
        }

        for r_idx, row_vals in enumerate(data):
            row = table.rows[r_idx]
            for c_idx in range(6):
                if r_idx == 0 and c_idx > 1:
                    continue
                cell = row.cells[c_idx]
                cell.width = Cm(first_cm if c_idx == 0 else other_cm)
                cell.text = ''
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                text_content = str(row_vals[c_idx]) if c_idx < len(row_vals) else ''
                run = p.add_run(text_content)
                run.font.name = '仿宋'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
                run.font.size = Pt(12)
                if (r_idx, c_idx) in bold_coords:
                    run.bold = True
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                self._set_docx_cell_border(cell, size=4)
                self._set_docx_cell_margins(cell, top=0, bottom=0, left=0, right=0)

        table.cell(0, 1).merge(table.cell(0, 5))

        doc.save(output_path)
        return output_path

    def _clear_cell_content(self, cell):
        """彻底清空单元格内容"""
        # 删除所有段落的内容
        for paragraph in cell.paragraphs:
            p = paragraph._element
            # 删除段落中的所有子元素（除了pPr）
            for child in list(p):
                if child.tag != qn('w:pPr'):
                    p.remove(child)
        # 确保只保留一个段落
        tc = cell._tc
        for p in list(tc.findall(qn('w:p')))[1:]:
            tc.remove(p)

    def _set_cell_text_with_format(self, cell, text, bold=False, font_name='仿宋', font_size=12):
        """设置单元格文本并格式化"""
        self._clear_cell_content(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text)
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(font_size)
        run.bold = bold
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def _set_row_height(self, row, height_cm=1.0):
        """设置行高"""
        row.height = Cm(height_cm)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        # 同时通过XML确保行高生效
        trPr = row._tr.get_or_add_trPr()
        trHeight = trPr.find(qn('w:trHeight'))
        if trHeight is None:
            trHeight = OxmlElement('w:trHeight')
            trPr.append(trHeight)
        # 1cm = 567 twips (1 twip = 1/20 pt, 1 cm ≈ 28.35 pt)
        trHeight.set(qn('w:val'), str(int(height_cm * 567)))
        trHeight.set(qn('w:hRule'), 'exact')

    def _export_eval_result_docx(self, links, obj_keys, method_avgs, prev_data,
                                 total_attainment, expected_attainment, prev_total):
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        output_dir = os.path.join(root, 'outputs')
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(
            output_dir,
            '5.基于考核结果的课程目标达成情况评价结果表.docx'
        )

        doc = Document()
        table = doc.add_table(rows=1, cols=7)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        total_cm = 14.64
        col_w = Cm(total_cm / 7)

        headers = [
            '课程分目标',
            '考核环节',
            '分权重',
            '分值/满分',
            '学生实际得分平均分',
            '分目标达成值',
            '上一轮教学分目标达成值',
        ]

        header_row = table.rows[0]
        # 表头行高设置为自适应内容
        header_row.height_rule = WD_ROW_HEIGHT_RULE.AUTO
        
        for c_idx, text in enumerate(headers):
            cell = header_row.cells[c_idx]
            cell.width = col_w
            self._set_cell_text_with_format(cell, text, bold=True)
            self._set_docx_cell_border(cell, size=4)
            self._set_docx_cell_margins(cell, top=0, bottom=0, left=0, right=0)

        trPr = header_row._tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), '1')
        trPr.append(tblHeader)

        # 先收集所有数据行信息
        all_data_rows = []  # [(obj_name, obj_rows, obj_attainment, prev_val), ...]
        
        for idx, obj_key in enumerate(obj_keys):
            obj_name = f"课程目标{idx + 1}"
            obj_rows = []  # 该课程目标下的所有行

            obj_weight_sum = 0.0
            obj_actual_sum = 0.0

            for link in links:
                link_name = link.get('name', '')
                if '平时' in link_name:
                    display_link = '平时成绩'
                elif '期中' in link_name:
                    display_link = '期中考核'
                elif '期末' in link_name:
                    display_link = '期末考核'
                else:
                    display_link = link_name

                link_ratio = float(link.get('ratio', 0))
                methods = link.get('methods', []) or []

                support_sum = 0.0
                actual_sum = 0.0
                for m in methods:
                    supports = m.get('supports', {}) or {}
                    weight = float(supports.get(obj_key, 0))
                    support_sum += weight
                    method_avg = float(method_avgs.get(m.get('name', ''), 0))
                    actual_sum += (method_avg / 100.0) * weight * 100.0

                obj_weight_sum += support_sum * link_ratio * 100.0
                obj_actual_sum += (actual_sum / 100.0) * link_ratio * 100.0

                weight_text = f"{support_sum * link_ratio * 100.0:.1f}"
                actual_text = f"{(actual_sum / 100.0) * link_ratio * 100.0:.2f}"
                obj_rows.append((display_link, weight_text, actual_text))

            # 计算达成值
            obj_attainment = (obj_actual_sum / obj_weight_sum) if obj_weight_sum > 0 else 0
            
            # 获取上一轮数据
            prev_val = 0
            if prev_data:
                raw_prev = prev_data.get(obj_name, 0)
                if isinstance(raw_prev, dict):
                    prev_val = raw_prev.get('value', 0) or 0
                elif isinstance(raw_prev, (int, float)):
                    prev_val = raw_prev
                else:
                    try:
                        prev_val = float(raw_prev)
                    except Exception:
                        prev_val = 0

            all_data_rows.append((obj_name, obj_rows, obj_attainment, prev_val))

        # 现在创建所有数据行
        for obj_name, obj_rows, obj_attainment, prev_val in all_data_rows:
            obj_start = len(table.rows)
            
            for row_idx, (display_link, weight_text, actual_text) in enumerate(obj_rows):
                new_row = table.add_row()
                self._set_row_height(new_row, 1.0)
                row_cells = new_row.cells
                
                for c_idx in range(7):
                    cell = row_cells[c_idx]
                    cell.width = col_w
                    self._set_docx_cell_border(cell, size=4)
                    self._set_docx_cell_margins(cell, top=0, bottom=0, left=0, right=0)
                
                # 只在每组的第一行填充第0、5、6列
                if row_idx == 0:
                    # 第0列：课程分目标（加粗）
                    self._set_cell_text_with_format(row_cells[0], obj_name, bold=True)
                    # 第5列：分目标达成值
                    self._set_cell_text_with_format(row_cells[5], f"{obj_attainment:.3f}", bold=False)
                    # 第6列：上一轮达成值
                    prev_text = f"{prev_val:.3f}" if isinstance(prev_val, (int, float)) else str(prev_val)
                    self._set_cell_text_with_format(row_cells[6], prev_text, bold=False)
                
                # 第1列：考核环节（加粗）
                self._set_cell_text_with_format(row_cells[1], display_link, bold=True)
                # 第2列：分权重
                self._set_cell_text_with_format(row_cells[2], weight_text, bold=False)
                # 第3列：分值/满分
                self._set_cell_text_with_format(row_cells[3], "100", bold=False)
                # 第4列：学生实际得分平均分
                self._set_cell_text_with_format(row_cells[4], actual_text, bold=False)

            obj_end = len(table.rows) - 1
            
            # 合并单元格（如果有多行）
            if obj_end > obj_start:
                # 合并第0列
                table.cell(obj_start, 0).merge(table.cell(obj_end, 0))
                # 合并第5列
                table.cell(obj_start, 5).merge(table.cell(obj_end, 5))
                # 合并第6列
                table.cell(obj_start, 6).merge(table.cell(obj_end, 6))

        def add_total_row(label, value):
            new_row = table.add_row()
            self._set_row_height(new_row, 1.0)
            
            row_cells = new_row.cells
            
            for c_idx in range(7):
                cell = row_cells[c_idx]
                cell.width = col_w
                self._set_docx_cell_border(cell, size=4)
                self._set_docx_cell_margins(cell, top=0, bottom=0, left=0, right=0)
            
            # 先填充内容
            self._set_cell_text_with_format(row_cells[0], label, bold=True)
            value_text = f"{value:.3f}" if isinstance(value, (int, float)) else str(value)
            self._set_cell_text_with_format(row_cells[6], value_text, bold=False)
            
            # 再合并单元格
            table.cell(len(table.rows) - 1, 0).merge(table.cell(len(table.rows) - 1, 5))

        add_total_row('课程目标达成值', total_attainment)
        add_total_row('课程目标达成期望值', expected_attainment)
        add_total_row('上一轮教学课程目标达成值', prev_total)

        doc.save(output_path)
        return output_path