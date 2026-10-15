from __future__ import annotations

import argparse
import datetime as dt
import tempfile
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont


DEFAULT_OUTPUT = Path(__file__).resolve().parents[3] / "examples" / "output" / "zz_coco_native_word_stress_fixture.docx"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a deterministic DOCX fixture for Coco native Word rendering.")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    args.out.parent.mkdir(parents=True, exist_ok=True)
    image_path = create_fixture_image(args.out.parent / "zz_coco_native_fixture_chart.png")

    doc = Document()
    doc.settings.odd_and_even_pages_header_footer = True
    configure_styles(doc)
    configure_section(doc.sections[0], width=Inches(8.27), height=Inches(11.69), margins=Inches(0.82))
    configure_header_footer(doc.sections[0], "Coco Native Stress Fixture | Portrait section", "Footer A", image_path)

    build_cover(doc)
    build_front_matter(doc, image_path)
    build_body_section(doc, image_path)
    build_landscape_section(doc, image_path)
    build_final_section(doc, image_path)

    add_comments(doc)
    doc.save(args.out)
    patch_anchored_images(args.out)
    patch_tracked_change(args.out)
    print(args.out)


def configure_styles(doc: Document) -> None:
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.space_after = Pt(6)
    set_east_asia_font(normal, "宋体")

    for name, size, color, before, after in [
        ("Title", 24, "0B2545", 0, 10),
        ("Heading 1", 16, "2E74B5", 14, 8),
        ("Heading 2", 13, "2E74B5", 10, 6),
        ("Heading 3", 12, "1F4D78", 8, 4),
    ]:
        style = styles[name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
        set_east_asia_font(style, "黑体" if name != "Normal" else "宋体")
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)


def set_east_asia_font(style_or_run, font_name: str) -> None:
    rpr = style_or_run.element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)


def configure_section(section, width, height, margins) -> None:
    section.orientation = WD_ORIENT.LANDSCAPE if width > height else WD_ORIENT.PORTRAIT
    section.page_width = width
    section.page_height = height
    section.top_margin = margins
    section.right_margin = margins
    section.bottom_margin = margins
    section.left_margin = margins
    section.header_distance = Inches(0.45)
    section.footer_distance = Inches(0.45)


def configure_header_footer(section, header_text: str, footer_prefix: str, image_path: Path) -> None:
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    section.different_first_page_header_footer = True
    section.first_page_header.is_linked_to_previous = False
    section.first_page_footer.is_linked_to_previous = False
    section.even_page_header.is_linked_to_previous = False
    section.even_page_footer.is_linked_to_previous = False
    configure_header_story(section.header, f"{header_text} | DEFAULT", image_path)
    configure_header_story(section.first_page_header, f"{header_text} | FIRST", image_path)
    configure_header_story(section.even_page_header, f"{header_text} | EVEN", image_path)
    configure_footer_story(section.footer, f"{footer_prefix} DEFAULT", image_path)
    configure_footer_story(section.first_page_footer, f"{footer_prefix} FIRST", image_path)
    configure_footer_story(section.even_page_footer, f"{footer_prefix} EVEN", image_path)


def configure_header_story(story, text: str, image_path: Path) -> None:
    paragraph = story.paragraphs[0]
    paragraph.text = text
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.add_run(" ")
    picture = paragraph.add_run().add_picture(str(image_path), width=Inches(0.32))
    drawing = picture._inline.getparent()
    doc_pr = drawing.find(".//" + qn("wp:docPr")) if drawing is not None else None
    if doc_pr is not None:
        anchor_marker = header_anchor_marker(text)
        doc_pr.set("name", anchor_marker)
        doc_pr.set("descr", anchor_marker)
    if paragraph.runs:
        paragraph.runs[0].font.size = Pt(9)
        paragraph.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)


def configure_footer_story(story, label: str, image_path: Path) -> None:
    paragraph = story.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_paragraph_bottom_border(paragraph)
    picture = paragraph.add_run().add_picture(str(image_path), width=Inches(0.2))
    drawing = picture._inline.getparent()
    doc_pr = drawing.find(".//" + qn("wp:docPr")) if drawing is not None else None
    if doc_pr is not None:
        anchor_marker = footer_anchor_marker(label)
        doc_pr.set("name", anchor_marker)
        doc_pr.set("descr", anchor_marker)
    paragraph.add_run(" ")
    run = paragraph.add_run(f"{label} | Page ")
    run.font.size = Pt(9)
    add_word_field(paragraph, " PAGE ")
    paragraph.add_run(" of ").font.size = Pt(9)
    add_word_field(paragraph, " NUMPAGES ")
    paragraph.add_run(" | Section pages ").font.size = Pt(9)
    add_word_field(paragraph, " SECTIONPAGES ")


def header_anchor_marker(text: str) -> str:
    if "FIRST" in text and "Portrait section" in text:
        return "COCO_HEADER_ANCHOR_FIRST"
    return "COCO_HEADER_INLINE"


def footer_anchor_marker(label: str) -> str:
    if "FIRST" in label and "Footer A" in label:
        return "COCO_FOOTER_ANCHOR_FIRST"
    return "COCO_FOOTER_INLINE"


def add_word_field(paragraph, instruction: str) -> None:
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def add_paragraph_bottom_border(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "9AA6B2")
    p_bdr.append(bottom)


def build_cover(doc: Document) -> None:
    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("Coco Native Word Editor 渲染压力样本").bold = True
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run("覆盖分节、页眉页脚、复杂表格、图片、批注与修订信号").italic = True
    doc.add_paragraph("生成目的：这个文档不是业务材料，而是用来暴露浏览器 native Word 渲染的边界。")
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = True
    p.add_run("目录页前分页段落：").bold = True
    p.add_run("用于验证 pageBreakBefore 是否能产生新页面。")


def build_front_matter(doc: Document, image_path: Path) -> None:
    doc.add_heading("1. 样式继承与段落格式", level=1)
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Inches(0.35)
    p.paragraph_format.left_indent = Inches(0.15)
    p.paragraph_format.right_indent = Inches(0.1)
    p.add_run("这一段包含首行缩进、左右缩进、")
    bold = p.add_run("加粗")
    bold.bold = True
    p.add_run("、")
    red = p.add_run("红色强调")
    red.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    p.add_run("、英文 Office XML token、以及中文字体回退。")

    for text in ["真实列表项一：检查编号和缩进", "真实列表项二：检查换行后的悬挂缩进", "真实列表项三：检查段后距"]:
        doc.add_paragraph(text, style="List Number")

    doc.add_heading("2. 图片与图注", level=1)
    p = doc.add_paragraph("图 2-1 使用内嵌图片，后续会补充浮动锚点图片 fixture。")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(str(image_path), width=Inches(4.8))
    caption = doc.add_paragraph("图 2-1 复杂样式 fixture 示意图")
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption.runs[0].font.size = Pt(9)
    caption.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    add_style_probe_paragraph(doc)

    doc.add_heading("2.1 Anchor image wrap fixture", level=2)
    doc.add_paragraph(
        "The next paragraphs exercise native floating image layout paths: square, tight, behind, inFront and top/bottom.",
    )
    add_anchor_image_placeholder(doc, image_path, "COCO_ANCHOR_square", width=1.45)
    doc.add_paragraph(
        "Square wrap anchor paragraph. This long text should flow beside the floating image instead of pushing the image into a normal inline slot. "
        "The renderer should keep the anchor marked in DOM and avoid horizontal page overflow.",
    )
    add_anchor_image_placeholder(doc, image_path, "COCO_ANCHOR_tight", width=1.25)
    doc.add_paragraph(
        "Tight wrap anchor paragraph. For browser rendering we treat tight as a float-like layout, preserving the key visual intent.",
    )
    add_anchor_image_placeholder(doc, image_path, "COCO_ANCHOR_behind", width=1.65)
    doc.add_paragraph("Behind-text anchor paragraph. The image should be promoted to the appropriate overlay-like layer.")
    add_anchor_image_placeholder(doc, image_path, "COCO_ANCHOR_behindOverlap", width=3.8)
    doc.add_paragraph(
        "COCO_LAYER_PROBE behind-overlap paragraph: this text should remain readable above a large behindDoc image.",
    )
    add_anchor_image_placeholder(doc, image_path, "COCO_ANCHOR_inFront", width=1.35)
    doc.add_paragraph("In-front anchor paragraph. The image should remain absolute and above text content.")
    add_anchor_image_placeholder(doc, image_path, "COCO_ANCHOR_topBottom", width=1.55)
    doc.add_paragraph("Top-and-bottom anchor paragraph. The image should clear flow instead of behaving like a side float.")


def add_style_probe_paragraph(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(0.25)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(18)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(21)
    run = p.add_run("COCO_STYLE_PROBE exact-line indent font-size font-family")
    run.font.name = "Arial"
    run.font.size = Pt(13.5)
    run.font.color.rgb = RGBColor(0x22, 0x44, 0x88)


def build_body_section(doc: Document, image_path: Path) -> None:
    doc.add_heading("3. 跨页表格与单元格样式", level=1)
    build_complex_merge_table(doc)
    doc.add_paragraph("下表包含重复表头、合并单元格、底色、不同列宽以及足够多的行，用来验证分页和表格 fragment。")
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False
    widths = [0.55, 1.2, 2.1, 1.15, 1.2]
    headers = ["序号", "模块", "检查点", "状态", "备注"]
    for i, cell in enumerate(table.rows[0].cells):
        cell.text = headers[i]
        cell.width = Inches(widths[i])
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade_cell(cell, "E8EEF5")
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.runs[0].bold = True
    mark_repeat_header(table.rows[0])

    for idx in range(1, 54):
        row = table.add_row()
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        values = [
            str(idx),
            ["段落", "表格", "图片", "页眉页脚", "分节"][idx % 5],
            f"检查第 {idx} 行的换行、行高、边框和中文文本是否稳定渲染。",
            ["PASS", "WARN", "TODO"][idx % 3],
            "长备注文本用于制造换行；包含 Latin ABC、数字 12345 和中文标点。",
        ]
        for col, value in enumerate(values):
            cell = row.cells[col]
            cell.width = Inches(widths[col])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.text = value
            if col in (0, 3):
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if idx % 7 == 0:
                shade_cell(cell, "F7FBFF")

    doc.add_paragraph("REVISION_TARGET_OLD")
    doc.add_paragraph("表后段落：用于观察跨页表格之后的段落间距。")
    doc.add_picture(str(image_path), width=Inches(2.2))


def build_complex_merge_table(doc: Document) -> None:
    doc.add_heading("3.1 Complex merged table fixture", level=2)
    table = doc.add_table(rows=5, cols=4)
    table.style = "Table Grid"
    table.autofit = False
    widths = [1.1, 1.45, 1.65, 1.25]
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths[index])
            set_cell_margin(cell, top=80, right=140, bottom=80, left=140)

    header = table.rows[0].cells
    header[0].text = "Phase"
    header[1].merge(header[3])
    header[1].text = "Merged KPI Header"
    for cell in table.rows[0].cells:
        shade_cell(cell, "D9EAF7")
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if paragraph.runs:
                paragraph.runs[0].bold = True
    mark_repeat_header(table.rows[0])

    phase = table.cell(1, 0).merge(table.cell(3, 0))
    phase.text = "Phase A vertical merge"
    phase.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    shade_cell(phase, "EAF2F8")
    set_no_wrap(phase)

    rows = [
        ("Metric 1", "long wrapped description in a merged-table row", "PASS"),
        ("Metric 2", "margin and shading should survive native rendering", "WARN"),
        ("Metric 3", "rowspan should hide vMerge continuation cells", "PASS"),
    ]
    for offset, values in enumerate(rows, start=1):
        cells = table.rows[offset].cells
        for col, value in enumerate(values, start=1):
            cells[col].text = value
            if col == 3:
                cells[col].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    footer = table.rows[4].cells
    footer[0].text = "Total"
    footer[1].merge(footer[2])
    footer[1].text = "Horizontal merge footer"
    footer[3].text = "Done"
    shade_cell(footer[0], "F2F2F2")
    shade_cell(footer[1], "F2F2F2")
    shade_cell(footer[3], "F2F2F2")


def build_landscape_section(doc: Document, image_path: Path) -> None:
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.72)
    section.right_margin = Inches(0.68)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.68)
    section.header_distance = Inches(0.35)
    section.footer_distance = Inches(0.35)
    configure_header_footer(section, "Landscape section header | wide table", "Footer B", image_path)

    doc.add_heading("4. 横向分节宽表", level=1)
    wide = doc.add_table(rows=1, cols=8)
    wide.style = "Table Grid"
    wide.autofit = False
    for i, cell in enumerate(wide.rows[0].cells):
        cell.text = f"列 {i + 1}"
        shade_cell(cell, "E2F0D9")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].runs[0].bold = True
    mark_repeat_header(wide.rows[0])
    for row_idx in range(1, 14):
        cells = wide.add_row().cells
        for col_idx, cell in enumerate(cells):
            cell.text = f"R{row_idx}C{col_idx + 1} 宽表内容"
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx < 2 else WD_ALIGN_PARAGRAPH.LEFT


def build_final_section(doc: Document, image_path: Path) -> None:
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section(section, width=Inches(8.27), height=Inches(11.69), margins=Inches(0.75))
    configure_header_footer(section, "Final portrait section | comments and revisions", "Footer C", image_path)

    doc.add_heading("5. 批注、修订和边界段落", level=1)
    doc.add_paragraph("COMMENT_TARGET 这一段会挂载批注，用于验证 comment range 的解析和渲染。")
    doc.add_paragraph("COMMENT_TARGET_CLOSE_A close comment anchor A for bubble collision testing.")
    doc.add_paragraph("COMMENT_TARGET_CLOSE_B close comment anchor B for bubble collision testing.")
    doc.add_paragraph("短段落。")
    doc.add_paragraph("    包含前导空格、制表意图和较长中文文本，观察空白是否被过度吞并。")
    p = doc.add_paragraph("显式分页符之前的段落")
    p.add_run().add_break(WD_BREAK.PAGE)
    doc.add_paragraph("显式分页符之后的段落")


def add_comments(doc: Document) -> None:
    comment_targets = {
        "COMMENT_TARGET_CLOSE_A": "Second comment: should stack below nearby comment bubbles.",
        "COMMENT_TARGET_CLOSE_B": "Third comment: verifies same-page bubble collision avoidance.",
        "COMMENT_TARGET": "这是用于 native 批注渲染的结构化评论。",
    }
    applied = set()
    for paragraph in doc.paragraphs:
        for marker, body in comment_targets.items():
            if marker in applied or marker not in paragraph.text or not paragraph.runs:
                continue
            doc.add_comment(paragraph.runs[0], body, author="Coco QA", initials="CQ")
            applied.add(marker)


def add_anchor_image_placeholder(doc: Document, image_path: Path, marker: str, width: float) -> None:
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    run.add_picture(str(image_path), width=Inches(width))
    drawing = run._r.find(qn("w:drawing"))
    if drawing is None:
        raise RuntimeError(f"Anchor fixture drawing was not created for {marker}")
    doc_pr = drawing.find(".//" + qn("wp:docPr"))
    if doc_pr is None:
        raise RuntimeError(f"Anchor fixture docPr was not created for {marker}")
    doc_pr.set("name", marker)
    doc_pr.set("descr", marker)


def create_fixture_image(path: Path) -> Path:
    image = Image.new("RGB", (960, 420), "white")
    draw = ImageDraw.Draw(image)
    colors = ["#2E74B5", "#70AD47", "#FFC000", "#C00000"]
    draw.rectangle((0, 0, 959, 419), outline="#CCD6E0", width=3)
    draw.text((36, 28), "Coco Native Fixture Chart", fill="#0B2545")
    for index, color in enumerate(colors):
        left = 80 + index * 200
        height = 80 + index * 48
        draw.rectangle((left, 330 - height, left + 110, 330), fill=color)
        draw.text((left, 342), f"S{index + 1}", fill="#333333")
    draw.line((60, 330, 880, 330), fill="#666666", width=2)
    draw.line((60, 80, 60, 330), fill="#666666", width=2)
    image.save(path)
    return path


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_margin(cell, top: int, right: int, bottom: int, left: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margin = tc_pr.find(qn("w:tcMar"))
    if margin is None:
        margin = OxmlElement("w:tcMar")
        tc_pr.append(margin)
    for side, value in {
        "top": top,
        "right": right,
        "bottom": bottom,
        "left": left,
    }.items():
        side_el = margin.find(qn(f"w:{side}"))
        if side_el is None:
            side_el = OxmlElement(f"w:{side}")
            margin.append(side_el)
        side_el.set(qn("w:w"), str(value))
        side_el.set(qn("w:type"), "dxa")


def set_no_wrap(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    if tc_pr.find(qn("w:noWrap")) is None:
        tc_pr.append(OxmlElement("w:noWrap"))


def mark_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def patch_anchored_images(path: Path) -> None:
    specs = {
        "COCO_ANCHOR_square": {
            "wrap": '<wp:wrapSquare wrapText="bothSides"/>',
            "behind": "0",
            "pos_h": '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>',
            "pos_v": '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
        },
        "COCO_ANCHOR_tight": {
            "wrap": '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/></wp:wrapPolygon></wp:wrapTight>',
            "behind": "0",
            "pos_h": '<wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>',
            "pos_v": '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
        },
        "COCO_ANCHOR_behind": {
            "wrap": "<wp:wrapNone/>",
            "behind": "1",
            "pos_h": '<wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>',
            "pos_v": '<wp:positionV relativeFrom="page"><wp:posOffset>1371600</wp:posOffset></wp:positionV>',
        },
        "COCO_ANCHOR_behindOverlap": {
            "wrap": "<wp:wrapNone/>",
            "behind": "1",
            "pos_h": '<wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>',
            "pos_v": '<wp:positionV relativeFrom="page"><wp:posOffset>2057400</wp:posOffset></wp:positionV>',
        },
        "COCO_ANCHOR_inFront": {
            "wrap": "<wp:wrapNone/>",
            "behind": "0",
            "pos_h": '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>',
            "pos_v": '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
        },
        "COCO_ANCHOR_topBottom": {
            "wrap": "<wp:wrapTopAndBottom/>",
            "behind": "0",
            "pos_h": '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
            "pos_v": '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
        },
        "COCO_HEADER_ANCHOR_FIRST": {
            "wrap": "<wp:wrapNone/>",
            "behind": "0",
            "pos_h": '<wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>',
            "pos_v": '<wp:positionV relativeFrom="page"><wp:posOffset>360000</wp:posOffset></wp:positionV>',
        },
        "COCO_FOOTER_ANCHOR_FIRST": {
            "wrap": "<wp:wrapNone/>",
            "behind": "0",
            "pos_h": '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>',
            "pos_v": '<wp:positionV relativeFrom="page"><wp:posOffset>10000000</wp:posOffset></wp:positionV>',
        },
    }
    with ZipFile(path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    for marker, spec in specs.items():
        replaced = False
        for name in list(entries):
            if not is_word_story_xml(name):
                continue
            xml = entries[name].decode("utf-8")
            if f'name="{marker}"' not in xml:
                continue
            entries[name] = replace_inline_with_anchor(xml, marker, spec).encode("utf-8")
            replaced = True
        if not replaced:
            raise RuntimeError(f"Inline drawing marker was not found in Word story XML: {marker}")
    rewrite_docx(path, entries)


def is_word_story_xml(name: str) -> bool:
    if name == "word/document.xml":
        return True
    return (
        name.startswith("word/header")
        or name.startswith("word/footer")
    ) and name.endswith(".xml")


def replace_inline_with_anchor(xml: str, marker: str, spec: dict[str, str]) -> str:
    marker_index = xml.find(f'name="{marker}"')
    if marker_index < 0:
        raise RuntimeError(f"Inline drawing marker was not found in document.xml: {marker}")
    start = xml.rfind("<wp:inline", 0, marker_index)
    end = xml.find("</wp:inline>", marker_index)
    if start < 0 or end < 0:
        raise RuntimeError(f"Inline drawing container was not found for marker: {marker}")
    end += len("</wp:inline>")
    inline_xml = xml[start:end]
    open_end = inline_xml.find(">")
    if open_end < 0:
        raise RuntimeError(f"Inline drawing opening tag was malformed for marker: {marker}")
    inner = inline_xml[open_end + 1 : -len("</wp:inline>")]
    extent_xml, remaining_inner = split_anchor_extent_prefix(inner)
    anchor_xml = (
        f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'relativeHeight="251658240" behindDoc="{spec["behind"]}" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'{spec["pos_h"]}'
        f'{spec["pos_v"]}'
        f"{extent_xml}"
        f'{spec["wrap"]}'
        f"{remaining_inner}"
        "</wp:anchor>"
    )
    return xml[:start] + anchor_xml + xml[end:]


def split_anchor_extent_prefix(inner: str) -> tuple[str, str]:
    prefix = ""
    remaining = inner
    for tag in ("extent", "effectExtent"):
        start_token = f"<wp:{tag}"
        if not remaining.startswith(start_token):
            continue
        end = remaining.find("/>")
        if end < 0:
            break
        end += len("/>")
        prefix += remaining[:end]
        remaining = remaining[end:]
    return prefix, remaining


def patch_tracked_change(path: Path) -> None:
    with ZipFile(path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    xml = entries["word/document.xml"].decode("utf-8")
    replacement = (
        '<w:del w:id="41" w:author="Coco QA" w:date="2026-08-11T09:00:00Z">'
        "<w:r><w:delText>旧版风险描述</w:delText></w:r>"
        "</w:del>"
        '<w:ins w:id="42" w:author="Coco QA" w:date="2026-08-11T09:05:00Z">'
        "<w:r><w:t>新版风险描述：修订插入文本应被高亮显示</w:t></w:r>"
        "</w:ins>"
    )
    source = "<w:p><w:r><w:t>REVISION_TARGET_OLD</w:t></w:r></w:p>"
    target = f"<w:p>{replacement}</w:p>"
    if source not in xml:
        raise RuntimeError("Revision placeholder paragraph was not found in document.xml")
    xml = xml.replace(source, target, 1)
    entries["word/document.xml"] = xml.encode("utf-8")
    rewrite_docx(path, entries)


def rewrite_docx(path: Path, entries: dict[str, bytes]) -> None:
    with tempfile.NamedTemporaryFile(delete=False, dir=path.parent, suffix=".docx") as tmp:
        temp_path = Path(tmp.name)
    with ZipFile(temp_path, "w", ZIP_DEFLATED) as zout:
        for name, content in entries.items():
            zout.writestr(name, content)
    temp_path.replace(path)


if __name__ == "__main__":
    main()
