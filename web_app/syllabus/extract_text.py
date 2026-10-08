"""把 docx / 文字 PDF 抽成按阅读顺序排列的段落和表格。"""
from __future__ import annotations

import io

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


class DocumentError(ValueError):
    pass


DOC_MESSAGE = "暂不支持 .doc，请在 Word 中另存为 .docx 后上传。"
SCAN_MESSAGE = "这是扫描件/图片 PDF，请上传 Word 版大纲或带文字的 PDF。"
TYPE_MESSAGE = "只接受 .docx，或能复制出文字的 PDF。"


def _cell_text(cell) -> str:
    return " ".join(cell.text.split())


def docx_to_markdown(data: bytes) -> str:
    try:
        document = Document(io.BytesIO(data))
    except Exception as exc:
        raise DocumentError("无法读取这份 Word 文件，请确认它是 .docx。") from exc
    lines: list[str] = []
    table_index = 0
    for child in document.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            text = Paragraph(child, document).text.strip()
            if text:
                lines.append(text)
        elif tag == "tbl":
            table_index += 1
            table = Table(child, document)
            lines.append(f"[表格{table_index}]")
            for row in table.rows:
                cells = []
                previous = None
                for cell in row.cells:
                    cells.append("" if cell._tc is previous else _cell_text(cell))
                    previous = cell._tc
                lines.append("| " + " | ".join(cells) + " |")
            lines.append(f"[/表格{table_index}]")
    return "\n".join(lines)


def pdf_to_markdown(data: bytes) -> str:
    import pdfplumber

    try:
        pdf = pdfplumber.open(io.BytesIO(data))
    except Exception as exc:
        raise DocumentError(SCAN_MESSAGE) from exc
    lines: list[str] = []
    try:
        for page_index, page in enumerate(pdf.pages, 1):
            tables = page.find_tables() or []
            boxes = [table.bbox for table in tables]
            items: list[tuple[float, str]] = []

            def outside(obj):
                return not any(
                    box[0] <= obj["x0"] and obj["x1"] <= box[2] and box[1] <= obj["top"] and obj["bottom"] <= box[3]
                    for box in boxes
                )

            try:
                text_lines = page.filter(outside).extract_text_lines() or []
            except Exception:
                text_lines = []
            for line in text_lines:
                items.append((float(line.get("top") or 0), str(line.get("text") or "")))
            for index, table in enumerate(tables, 1):
                rows = table.extract() or []
                block = [f"[表格p{page_index}-{index}]"]
                for row in rows:
                    cells = [" ".join(str(cell or "").split()) for cell in row]
                    block.append("| " + " | ".join(cells) + " |")
                block.append(f"[/表格p{page_index}-{index}]")
                items.append((float(table.bbox[1]), "\n".join(block)))
            if not items:
                plain = page.extract_text() or ""
                if plain.strip():
                    lines.append(plain.strip())
            else:
                lines.extend(text for _, text in sorted(items, key=lambda item: item[0]) if text)
    finally:
        pdf.close()
    return "\n".join(lines)


def pdf_char_count(data: bytes) -> int:
    return len("".join(pdf_to_markdown(data).split()))


def suffix_of(name: str) -> str:
    lower = (name or "").lower().strip()
    if lower.endswith(".docx"):
        return ".docx"
    if lower.endswith(".doc"):
        return ".doc"
    if lower.endswith(".pdf"):
        return ".pdf"
    if lower.endswith(".xlsx"):
        return ".xlsx"
    if lower.endswith(".xls"):
        return ".xls"
    dot = lower.rfind(".")
    return lower[dot:] if dot >= 0 else ""
