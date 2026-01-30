
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
        output_path = os.path.join(output_dir, '\u0032.\u8bfe\u7a0b\u6210\u7ee9\u7edf\u8ba1\u8868.docx')

        doc = Document()
        table = doc.add_table(rows=5, cols=6)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        total_cm = 14.64
        first_cm = 3.75
        other_cm = (total_cm - first_cm) / 5

        data = [
            ['\u6210\u7ee9\u6784\u6210', composition_text.strip(), '', '', '', ''],
            ['\u6700\u9ad8\u6210\u7ee9', max_score, '\u6700\u4f4e\u6210\u7ee9', min_score, '\u5e73\u5747\u6210\u7ee9', avg_score],
            ['\u6210\u7ee9\u7b49\u7ea7', '90-100\n(\u4f18\u79c0)', '80-89\n(\u826f\u597d)', '70-79\n(\u4e2d\u7b49)', '60-69\n(\u53ca\u683c)', '<60\n(\u4e0d\u53ca\u683c)'],
            ['\u4eba\u6570'] + list(counts),
            ['\u5360\u8003\u6838\u4eba\u6570\u7684\u6bd4\u4f8b'] + [f"{r*100:.2f}%" for r in ratios],
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
                run.font.name = '\u4eff\u5b8b'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '\u4eff\u5b8b')
                run.font.size = Pt(12)
                if (r_idx, c_idx) in bold_coords:
                    run.bold = True
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                self._set_docx_cell_border(cell, size=4)
                self._set_docx_cell_margins(cell, top=0, bottom=0, left=0, right=0)

        table.cell(0, 1).merge(table.cell(0, 5))

        doc.save(output_path)
        return output_path

    def _export_eval_result_docx(self, links, obj_keys, method_avgs, prev_data,
                                 total_attainment, expected_attainment, prev_total):
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        output_dir = os.path.join(root, 'outputs')
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(
            output_dir,
            '\u0035.\u57fa\u4e8e\u8003\u6838\u7ed3\u679c\u7684\u8bfe\u7a0b\u76ee\u6807\u8fbe\u6210\u60c5\u51b5\u8bc4\u4ef7\u7ed3\u679c\u8868.docx'
        )

        doc = Document()
        table = doc.add_table(rows=1, cols=7)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        total_cm = 14.64
        col_w = Cm(total_cm / 7)

        headers = [
            '\u8bfe\u7a0b\u5206\u76ee\u6807',
            '\u8003\u6838\u73af\u8282',
            '\u5206\u6743\u91cd',
            '\u5206\u503c/\u6ee1\u5206',
            '\u5b66\u751f\u5b9e\u9645\u5f97\u5206\u5e73\u5747\u5206',
            '\u5206\u76ee\u6807\u8fbe\u6210\u503c',
            '\u4e0a\u4e00\u8f6e\u6559\u5b66\u5206\u76ee\u6807\u8fbe\u6210\u503c',
        ]

        header_row = table.rows[0]
        for c_idx, text in enumerate(headers):
            cell = header_row.cells[c_idx]
            cell.width = col_w
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            run.font.name = '\u4eff\u5b8b'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '\u4eff\u5b8b')
            run.font.size = Pt(12)
            run.bold = True
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            self._set_docx_cell_border(cell, size=4)
            self._set_docx_cell_margins(cell, top=0, bottom=0, left=0, right=0)

        trPr = header_row._tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), '1')
        trPr.append(tblHeader)

        row_map = []
        for idx, obj_key in enumerate(obj_keys):
            obj_name = f"\u8bfe\u7a0b\u76ee\u6807{idx + 1}"
            obj_start = len(table.rows)

            obj_weight_sum = 0.0
            obj_actual_sum = 0.0

            for link in links:
                link_name = link.get('name', '')
                if '\u5e73\u65f6' in link_name:
                    display_link = '\u5e73\u65f6\u6210\u7ee9'
                elif '\u671f\u4e2d' in link_name:
                    display_link = '\u671f\u4e2d\u8003\u6838'
                elif '\u671f\u672b' in link_name:
                    display_link = '\u671f\u672b\u8003\u6838'
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

                row_cells = table.add_row().cells
                row_cells[0].text = obj_name
                row_cells[1].text = display_link
                row_cells[2].text = f"{support_sum * link_ratio * 100.0:.1f}"
                row_cells[3].text = "100"
                row_cells[4].text = f"{(actual_sum / 100.0) * link_ratio * 100.0:.2f}"
                row_cells[5].text = f"{(obj_actual_sum / obj_weight_sum) if obj_weight_sum > 0 else 0:.3f}"
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
                row_cells[6].text = f"{prev_val:.3f}" if isinstance(prev_val, (int, float)) else str(prev_val)

                for c_idx in range(7):
                    cell = row_cells[c_idx]
                    cell.width = col_w
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    run = p.runs[0]
                    run.font.name = '\u4eff\u5b8b'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '\u4eff\u5b8b')
                    run.font.size = Pt(12)
                    run.bold = False
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    self._set_docx_cell_border(cell, size=4)
                    self._set_docx_cell_margins(cell, top=0, bottom=0, left=0, right=0)

            obj_end = len(table.rows) - 1
            if obj_end > obj_start:
                table.cell(obj_start, 0).merge(table.cell(obj_end, 0))
                table.cell(obj_start, 5).merge(table.cell(obj_end, 5))
                table.cell(obj_start, 6).merge(table.cell(obj_end, 6))

            row_map.append((obj_start, obj_end))

        def add_total_row(label, value):
            row_cells = table.add_row().cells
            row_cells[0].text = label
            row_cells[6].text = f"{value:.3f}" if isinstance(value, (int, float)) else str(value)
            table.cell(len(table.rows) - 1, 0).merge(table.cell(len(table.rows) - 1, 5))
            for c_idx in range(7):
                cell = row_cells[c_idx]
                cell.width = col_w
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                if p.runs:
                    run = p.runs[0]
                else:
                    run = p.add_run('')
                run.font.name = '\u4eff\u5b8b'
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '\u4eff\u5b8b')
                run.font.size = Pt(12)
                run.bold = True if c_idx == 0 else False
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                self._set_docx_cell_border(cell, size=4)
                self._set_docx_cell_margins(cell, top=0, bottom=0, left=0, right=0)

        add_total_row('\u8bfe\u7a0b\u76ee\u6807\u8fbe\u6210\u503c', total_attainment)
        add_total_row('\u8bfe\u7a0b\u76ee\u6807\u8fbe\u6210\u671f\u671b\u503c', expected_attainment)
        add_total_row('\u4e0a\u4e00\u8f6e\u6559\u5b66\u8bfe\u7a0b\u76ee\u6807\u8fbe\u6210\u503c', prev_total)

        doc.save(output_path)
        return output_path
