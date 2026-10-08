"""合成登记表：学生状态、名单间隔、跨页表头与人数文字。"""
import io

import pytest
from openpyxl import Workbook

from web_app.grade_register import parse_register


HEADERS = ["班级", "学号", "姓名", "平时(30%)", "期末(70%)", "总评", "备注"]
STUDENTS = [
    ["合成班", "10000001", "合成甲", 80, 90, 87, ""],
    ["合成班", "10000002", "合成乙", 0, 0, 0, ""],
    ["合成班", "10000003", "合成丙", 75, 88, 84.1, ""],
]


def register_bytes(rows, footer=None):
    book = Workbook()
    sheet = book.active
    sheet.append(HEADERS)
    for row in rows:
        sheet.append(row)
    if footer is not None:
        sheet.append(footer)
    blob = io.BytesIO()
    book.save(blob)
    return blob.getvalue()


@pytest.mark.parametrize("status", ["缓考", "缺考", "优秀", "不及格"])
@pytest.mark.parametrize("column", [4, 6])
def test_student_status_does_not_end_roster(status, column):
    rows = [list(student) for student in STUDENTS]
    rows[1][column] = status
    parsed = parse_register("synthetic.xlsx", register_bytes(rows, ["实考 2人 总人数 3人"]))
    assert [student["name"] for student in parsed["students"]] == ["合成甲", "合成乙", "合成丙"]
    assert parsed["student_count"] == 3
    assert parsed["exam_count"] == 2
    assert parsed["students"][1]["final"] == (None if column == 4 else 0)


@pytest.mark.parametrize("label", ["成绩统计", "平均成绩", "实考 3人", "缓考 1人", "分数段", "优秀", "不及格"])
def test_statistics_section_is_not_counted_as_students(label):
    footer = [label, "", "统计数字", 1, 2, 3]
    rows = STUDENTS + [footer, ["", "", "非学生行", 1, 2, 3]]
    parsed = parse_register("synthetic.xlsx", register_bytes(rows))
    assert parsed["student_count"] == 3
    assert [student["name"] for student in parsed["students"]] == ["合成甲", "合成乙", "合成丙"]


@pytest.mark.parametrize("gap", [[], ["", "", "", None, 80], HEADERS, ["", "", "姓名"]])
def test_roster_continues_after_empty_rows_and_repeated_headers(gap):
    rows = [STUDENTS[0], gap, STUDENTS[1], HEADERS, STUDENTS[2]]
    parsed = parse_register("synthetic.xlsx", register_bytes(rows, ["成绩统计"]))
    assert parsed["student_count"] == 3
    assert [student["name"] for student in parsed["students"]] == ["合成甲", "合成乙", "合成丙"]
    assert parsed["students"][1]["final"] == 0


def test_pdf_repeated_header_on_second_page_keeps_all_students():
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import PageBreak, SimpleDocTemplate, Table, TableStyle

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    blob = io.BytesIO()
    document = SimpleDocTemplate(blob, pagesize=A4)
    style = TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ])
    first = Table([HEADERS] + STUDENTS[:1], colWidths=[55, 70, 50, 65, 65, 50, 40])
    second = Table([HEADERS] + STUDENTS[1:], colWidths=[55, 70, 50, 65, 65, 50, 40])
    first.setStyle(style)
    second.setStyle(style)
    document.build([first, PageBreak(), second])
    parsed = parse_register("synthetic.pdf", blob.getvalue())
    assert parsed["student_count"] == 3
    assert [student["name"] for student in parsed["students"]] == ["合成甲", "合成乙", "合成丙"]
