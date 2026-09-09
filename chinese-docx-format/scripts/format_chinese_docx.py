#!/usr/bin/env python3
"""把已有 DOCX 迁移到本技能的样式基线并输出新文件。

脚本只在内存中读取输入，绝不覆盖输入文件；正文文字、表格、图片和 OMML
对象保留，手工标题编号作为非内容格式被移除后由样式编号重新生成。
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from build_base_template import (
    add_baseline_marker,
    configure_document,
    configure_styles,
    find,
    remove_all,
    tag,
)


M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def clean(text: str | None) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def style_text(paragraph) -> str:
    style = paragraph.style
    return f"{style.style_id} {style.name}".lower() if style is not None else ""


def in_table(paragraph) -> bool:
    node = paragraph._p
    while node is not None:
        if node.tag == tag("tc"):
            return True
        node = node.getparent()
    return False


def has_omml(paragraph) -> bool:
    return any(node.tag.startswith("{" + M + "}") for node in paragraph._p.iter())


def is_toc_entry(paragraph) -> bool:
    """目录缓存通常由超链接组成，不能当作正文标题重新编号。"""
    return any(node.tag == tag("hyperlink") for node in paragraph._p.iter())


def toc_role(paragraph) -> str | None:
    if not is_toc_entry(paragraph):
        return None
    text = clean(paragraph.text)
    if re.match(r"^\d+\.\d+\.\d+", text):
        return "TOC 3"
    if re.match(r"^\d+\.\d+", text):
        return "TOC 2"
    return "TOC 1"


def all_table_paragraphs(tables, result: list) -> None:
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph not in result:
                        result.append(paragraph)
                all_table_paragraphs(cell.tables, result)


def all_paragraphs(document: Document) -> list:
    result = list(document.paragraphs)
    all_table_paragraphs(document.tables, result)
    unique = []
    seen = set()
    for paragraph in result:
        key = id(paragraph._p)
        if key not in seen:
            seen.add(key)
            unique.append(paragraph)
    return unique


def looks_like_heading(text: str) -> int | None:
    if not text or len(text) > 100 or text.endswith(("。", "；", ";", "！", "!", "？", "?")):
        return None
    if re.match(r"^第[一二三四五六七八九十百千万零〇]+章(?:\s|$)", text):
        return 1
    if re.match(r"^[一二三四五六七八九十百千万零〇]+[、.．]\s*", text):
        return 1
    if re.match(r"^\d+\.\d+\.\d+(?:\s|$)", text):
        return 3
    if re.match(r"^\d+\.\d+(?:\s|$)", text):
        return 2
    return None


def is_technical(text: str, raw_text: str | None = None) -> bool:
    if not text or len(text) > 240:
        return False
    raw_text = raw_text or text
    technical_markers = (
        "::",
        "=>",
        "->",
        "↓",
        "{",
        "}",
        "GET ",
        "POST ",
        "PUT ",
        "DELETE ",
        "SELECT ",
        "import ",
        "from ",
        "def ",
        "class ",
    )
    if any(text.startswith(marker) for marker in ("GET ", "POST ", "PUT ", "DELETE ", "SELECT ", "import ", "from ", "def ", "class ")):
        return True
    marker_count = sum(text.count(marker) for marker in technical_markers)
    if marker_count >= 2:
        return True
    if "\n" in raw_text and (marker_count or text.count("/") >= 2):
        return True
    return False


def classify(paragraph, is_first_nonempty: bool) -> str:
    text = clean(paragraph.text)
    raw_text = paragraph.text or ""
    style = style_text(paragraph)
    toc = toc_role(paragraph)
    if toc:
        return toc
    if has_omml(paragraph):
        return "Equation"
    if "code block" in style or "codeblock" in style or "代码" in style or "源码" in style:
        return "Code Block"
    if "figure caption" in style or "图题" in style or re.match(r"^图\s*[0-9A-Za-z]", text):
        return "Figure Caption"
    if "table caption" in style or "表题" in style or re.match(r"^表\s*[0-9A-Za-z]", text):
        return "Table Caption"
    if "reference" in style or "endnote bibliography" in style:
        return "Reference"
    if "appendix heading" in style or re.match(r"^附录\s*[0-9A-Za-z]", text):
        return "Appendix Heading"
    if text in {"摘要", "摘 要"}:
        return "Abstract Heading"
    if text == "ABSTRACT":
        return "English Abstract Heading"
    if text in {"目录", "目 录"}:
        return "TOC Title"
    if text == "参考文献":
        return "Heading 1"
    if "abstract body" in style and "english" in style:
        return "English Abstract Body"
    if "abstract body" in style:
        return "Abstract Body"
    if "heading 1" in style or "heading1" in style or "标题 1" in style:
        return "Heading 1"
    if "heading 2" in style or "heading2" in style or "标题 2" in style:
        return "Heading 2"
    if "heading 3" in style or "heading3" in style or "标题 3" in style:
        return "Heading 3"
    if "title" in style and is_first_nonempty:
        return "Title"
    if is_first_nonempty and not in_table(paragraph) and text and len(text) <= 100:
        return "Title"
    if not in_table(paragraph) and looks_like_heading(text):
        return f"Heading {looks_like_heading(text)}"
    if is_technical(text, raw_text):
        return "Code Block"
    if in_table(paragraph):
        return "Table Text"
    return "Normal"


def strip_heading_prefix(paragraph, role: str) -> None:
    if role not in {"Heading 1", "Heading 2", "Heading 3", "Appendix Heading"}:
        return
    text = clean(paragraph.text)
    patterns = {
        "Heading 1": r"^\s*(?:第[一二三四五六七八九十百千万零〇]+章|[一二三四五六七八九十百千万零〇]+[、.．]|\d+\.)\s*",
        "Heading 2": r"^\s*\d+\.\d+\s*",
        "Heading 3": r"^\s*\d+\.\d+\.\d+\s*",
        "Appendix Heading": r"^\s*附录\s*[0-9A-Za-z]+\s*",
    }
    match = re.match(patterns[role], text)
    if not match or not paragraph.runs:
        return
    remaining = len(match.group(0))
    for run in paragraph.runs:
        if remaining <= 0:
            break
        value = run.text or ""
        if not value:
            continue
        remove_count = min(remaining, len(value))
        run.text = value[remove_count:]
        remaining -= remove_count


def clean_paragraph_properties(paragraph, style_name: str, style_id: str) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    for child in list(ppr):
        if child.tag != tag("sectPr"):
            ppr.remove(child)
    ppr = paragraph._p.get_or_add_pPr()
    pstyle = OxmlElement("w:pStyle")
    pstyle.set(tag("val"), style_id)
    ppr.insert(0, pstyle)


def clean_run_properties(paragraph) -> None:
    for marker in list(paragraph._p.iter(tag("proofErr"))):
        parent = marker.getparent()
        if parent is not None:
            parent.remove(marker)
    for run in paragraph._p.iter(tag("r")):
        rpr = run.find(tag("rPr"))
        if rpr is None:
            continue
        # 保留上标、下标和语言信息，清除会覆盖样式的视觉属性。
        for child in list(rpr):
            if child.tag not in {tag("vertAlign"), tag("lang"), tag("rtl")}:
                rpr.remove(child)
        if len(rpr) == 0:
            run.remove(rpr)


def set_cell_width(cell, width_twips: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = find(tc_pr, "tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.insert(0, tc_w)
    tc_w.set(tag("w"), str(width_twips))
    tc_w.set(tag("type"), "dxa")


def set_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = find(tbl_pr, "tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = find(borders, side)
        if node is None:
            node = OxmlElement("w:" + side)
            borders.append(node)
        node.set(tag("val"), "single")
        node.set(tag("sz"), "4")
        node.set(tag("space"), "0")
        node.set(tag("color"), "D9D9D9")


def set_table_margins(table) -> None:
    tbl_pr = table._tbl.tblPr
    margins = find(tbl_pr, "tblCellMar")
    if margins is None:
        margins = OxmlElement("w:tblCellMar")
        tbl_pr.append(margins)
    for side, value in (("top", 80), ("left", 100), ("bottom", 80), ("right", 100)):
        node = find(margins, side)
        if node is None:
            node = OxmlElement("w:" + side)
            margins.append(node)
        node.set(tag("w"), str(value))
        node.set(tag("type"), "dxa")


def set_table_widths(table, available_twips: int) -> None:
    columns = len(table.columns)
    if columns == 0:
        return
    scores = []
    for column_index in range(columns):
        max_chars = 1
        for row in table.rows:
            if column_index < len(row.cells):
                value = max((len(clean(p.text)) for p in row.cells[column_index].paragraphs), default=1)
                max_chars = max(max_chars, min(value, 36))
        scores.append(max_chars)
    minimum = max(600, min(1000, available_twips // (columns * 2)))
    remaining = max(0, available_twips - minimum * columns)
    total_score = sum(scores) or columns
    widths = [minimum + round(remaining * score / total_score) for score in scores]
    widths[-1] += available_twips - sum(widths)

    tbl_pr = table._tbl.tblPr
    tbl_w = find(tbl_pr, "tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.insert(0, tbl_w)
    tbl_w.set(tag("w"), str(available_twips))
    tbl_w.set(tag("type"), "dxa")
    layout = find(tbl_pr, "tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(tag("type"), "fixed")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        column = OxmlElement("w:gridCol")
        column.set(tag("w"), str(width))
        grid.append(column)
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            if index < len(widths):
                set_cell_width(cell, widths[index])


def mark_header_row(table) -> None:
    if len(table.rows) < 2:
        return
    row = table.rows[0]._tr
    tr_pr = row.get_or_add_trPr()
    remove_all(tr_pr, "tblHeader")
    tr_pr.append(OxmlElement("w:tblHeader"))


def remove_fixed_row_heights(table) -> None:
    for row in table.rows:
        tr_pr = row._tr.find(tag("trPr"))
        if tr_pr is not None:
            remove_all(tr_pr, "trHeight")


def all_tables(document: Document) -> list:
    result = []

    def collect(tables) -> None:
        for table in tables:
            result.append(table)
            for row in table.rows:
                for cell in row.cells:
                    collect(cell.tables)

    collect(document.tables)
    return result


def format_tables(document: Document) -> None:
    if not document.sections:
        return
    section = document.sections[0]
    available = int(round((section.page_width - section.left_margin - section.right_margin) / 635))
    for table in all_tables(document):
        tbl_pr = table._tbl.tblPr
        for child_name in ("tblStyle", "tblLook", "tblInd", "tblCellSpacing", "tblpPr"):
            remove_all(tbl_pr, child_name)
        table.autofit = False
        set_table_widths(table, available)
        set_table_borders(table)
        set_table_margins(table)
        mark_header_row(table)
        remove_fixed_row_heights(table)
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    role = "Code Block" if is_technical(clean(paragraph.text), paragraph.text) else "Table Text"
                    style = document.styles[role]
                    clean_paragraph_properties(paragraph, role, style.style_id)
                    clean_run_properties(paragraph)


def scale_inline_images(document: Document) -> None:
    for section in document.sections:
        available = section.page_width - section.left_margin - section.right_margin
        for shape in document.inline_shapes:
            if shape.width <= available:
                continue
            ratio = available / shape.width
            shape.width = int(shape.width * ratio)
            shape.height = int(shape.height * ratio)


def format_document(input_path: Path, output_path: Path, profile: str) -> None:
    if input_path.resolve() == output_path.resolve():
        raise ValueError("输出文件必须与输入文件不同，原文档不会被覆盖")
    document = Document(str(input_path))
    original_paragraphs = all_paragraphs(document)
    first_nonempty = True
    roles = []
    for paragraph in original_paragraphs:
        role = classify(paragraph, first_nonempty)
        if clean(paragraph.text):
            first_nonempty = False
        roles.append((paragraph, role))

    configure_styles(document, reset_numbering=False)
    configure_document(document)

    for paragraph, role in roles:
        if role not in {style.name for style in document.styles}:
            role = "Normal"
        strip_heading_prefix(paragraph, role)
        style = document.styles[role]
        clean_paragraph_properties(paragraph, role, style.style_id)
        clean_run_properties(paragraph)

    format_tables(document)
    scale_inline_images(document)
    add_baseline_marker(document)
    document.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="将 DOCX 迁移到标准中文文档格式并输出新文件")
    parser.add_argument("input", type=Path, help="输入 DOCX，只读")
    parser.add_argument("--output", required=True, type=Path, help="输出 DOCX，不得与输入相同")
    parser.add_argument("--profile", choices=("auto", "thesis", "general"), default="auto")
    args = parser.parse_args()
    if args.input.suffix.lower() != ".docx" or not args.input.is_file():
        raise SystemExit("输入必须是存在的 .docx 文件")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    format_document(args.input.resolve(), args.output.resolve(), args.profile)
    print(f"已生成：{args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
