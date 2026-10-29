# -*- coding: utf-8 -*-
"""生成一份【虚构的、可公开】的英文学位论文模板 docx（海外版测试 / 演示用）。

设计要点
--------
- 校名一律用 "Example University"，不含任何真实学校 / 人名 / 机构，可安全公开。
- **Word 批注（comments.xml）是海外「学校模板」模式的一手规范来源**：本模板把
  每条格式要求写成英文批注，覆盖 heading / abstract / keywords / body / margins /
  paper size / three-line table / table spacing / references / captions。
- 模板自身样式与排版**与批注自洽**（批注说 18pt，样式就是 18pt），既能验证解析、
  也不会造出"模板自相矛盾"的假样本。（「批注与实测冲突时批注优先」在单测另造变体。）
- 三线表：只保留 top / header-bottom / bottom 三条横线，无竖线、无 body 行间横线。
- 🔴 批注注入的 XML 闭合标签是 ``</Types>`` / ``</Relationships>``（**不带冒号**）。
  国内旧脚本写作 ``</Types:`` / ``</Relationships:``，那是空操作 —— 批注文件会写进
  zip 却不建立关系，Word 能打开但引擎按关系读不到。此处已修正。

用法
----
    python tools/make_sample_template_en.py [输出路径]

默认输出到仓库 ``_smoke/example_university_template_en.docx``（_smoke 已 gitignore）。
"""

from __future__ import annotations

import os
import sys
import zipfile

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

FONT = "Times New Roman"

# 批注顺序即 id（1 基）。措辞模拟真实英文模板："类别词开头 + 具体要求"，
# 类别词落在开头 20 字符内，便于解析器归类。
COMMENTS = [
    "Title: Times New Roman 24 pt, bold, centered, single-spaced.",
    "Heading 1: Times New Roman 18 pt, bold, centered, 12 pt before, 6 pt after.",
    "Heading 2: Times New Roman 14 pt, bold, left-aligned, 12 pt before, 6 pt after.",
    "Heading 3: Times New Roman 12 pt, bold, italic, left-aligned.",
    "Abstract heading: Times New Roman 12 pt, bold, centered. Abstract body: 12 pt, double-spaced, no first-line indent.",
    "Keywords: Times New Roman 12 pt, label in italic, left-aligned.",
    "Body text: Times New Roman 12 pt, double-spaced, first-line indent 0.5 inch, justified.",
    "Margins: top 1 inch, bottom 1 inch, left 1 inch, right 1 inch.",
    "Paper size: Letter, 8.5 x 11 inches.",
    "Tables: use three-line tables (top rule, header rule, bottom rule) with no vertical lines.",
    "Leave one blank line after each table before the following text or heading.",
    "References: APA 7th edition, hanging indent 0.5 inch, double-spaced.",
    "Figure captions: Times New Roman 12 pt, bold label, placed below the figure.",
    "Table captions: Times New Roman 12 pt, bold label, placed above the table.",
    "Layout: single column, one column per page.",
]

DEFAULT_OUT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "_smoke", "example_university_template_en.docx",
)


# ---------------------------------------------------------------- 样式
def _style(doc, name, size_pt, *, bold=False, italic=False, align=None,
           line_spacing=None, first_line=None, space_before=None, space_after=None):
    st = doc.styles[name]
    st.font.name = FONT
    st.font.size = Pt(size_pt)
    st.font.bold = bold
    st.font.italic = italic
    # 西文槽显式写死，避免 Word 回退默认字体导致实测值与批注不一致
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(attr), FONT)
    pf = st.paragraph_format
    if align is not None:
        pf.alignment = align
    if line_spacing is not None:
        pf.line_spacing = line_spacing
    if first_line is not None:
        pf.first_line_indent = first_line
    if space_before is not None:
        pf.space_before = space_before
    if space_after is not None:
        pf.space_after = space_after
    return st


# ---------------------------------------------------------------- 三线表
def _three_line_table(doc):
    """插入 3×3 三线表：top / header-bottom / bottom 三条横线，无竖线。"""
    t = doc.add_table(rows=3, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(("Dimension", "Method A", "Method B")):
        t.cell(0, i).text = h
    for r in range(1, 3):
        for c in range(3):
            t.cell(r, c).text = "Sample %d-%d" % (r, c)

    tbl = t._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for el in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(el)
    borders = OxmlElement("w:tblBorders")
    for side, val in (("top", "single"), ("bottom", "single"),
                      ("left", "nil"), ("right", "nil"),
                      ("insideH", "nil"), ("insideV", "nil")):
        e = OxmlElement("w:" + side)
        e.set(qn("w:val"), val)
        if val == "single":
            e.set(qn("w:sz"), "8")
            e.set(qn("w:space"), "0")
            e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)

    # 表头行下横线（三线表中间那条）
    for cell in t.rows[0].cells:
        tcPr = cell._tc.get_or_add_tcPr()
        for el in tcPr.findall(qn("w:tcBorders")):
            tcPr.remove(el)
        tcb = OxmlElement("w:tcBorders")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "8")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        tcb.append(b)
        tcPr.append(tcb)
    return t


# ---------------------------------------------------------------- 批注锚点
def _mark_comment(paragraph, cid):
    """在段落上挂一条批注锚点（同一段可挂多条，Word 支持）。"""
    p = paragraph._p
    start = OxmlElement("w:commentRangeStart")
    start.set(qn("w:id"), str(cid))
    ppr = p.find(qn("w:pPr"))
    if ppr is not None:
        ppr.addnext(start)
    else:
        p.insert(0, start)
    end = OxmlElement("w:commentRangeEnd")
    end.set(qn("w:id"), str(cid))
    p.append(end)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rs = OxmlElement("w:rStyle")
    rs.set(qn("w:val"), "CommentReference")
    rpr.append(rs)
    r.append(rpr)
    cr = OxmlElement("w:commentReference")
    cr.set(qn("w:id"), str(cid))
    r.append(cr)
    p.append(r)


def _format_runs(paragraph, size_pt, bold=True):
    """给段落里所有 run 写直接格式（字体/字号/加粗）。

    为什么不直接用 Word 的 "Title" 样式：``heading_level(Title)`` 返回 1，
    会让 ``detect_title_page`` 早退（docx_reader.py:3810），标题不被排除，
    结果 24pt 的论文大标题占掉 L1 字号档 —— 实测踩过。用 Normal + 直接格式
    才能让标题页（含作者块）被正确识别并跳过。
    """
    for r in paragraph.runs:
        r.font.name = FONT
        r.font.size = Pt(size_pt)
        r.font.bold = bold
        rpr = r._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(attr), FONT)


def build_with_anchors(tmp_path):
    """生成正文并在保存前挂好全部批注锚点，返回锚点条数。"""
    doc = Document()

    # 页面：Letter + 1 英寸四边（与批注 8/9 自洽）
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)

    # 样式（与批注自洽）
    _style(doc, "Normal", 12, line_spacing=2.0, first_line=Inches(0.5),
           align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    _style(doc, "Title", 24, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
           line_spacing=1.0)
    _style(doc, "Heading 1", 18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
           space_before=Pt(12), space_after=Pt(6))
    _style(doc, "Heading 2", 14, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
           space_before=Pt(12), space_after=Pt(6))
    _style(doc, "Heading 3", 12, bold=True, italic=True,
           align=WD_ALIGN_PARAGRAPH.LEFT)

    def add(text, style=None):
        return doc.add_paragraph(text, style=style)

    n = 0

    def mark(cid, para):
        nonlocal n
        _mark_comment(para, cid)
        n += 1

    # 论文大标题：Normal + 直接格式（24pt 粗体居中），不用 "Title" 样式 ——
    # 见 _format_runs 的说明（Title 样式会被 heading_level 判成 1 级，顶掉 L1 采样）。
    p_title = add("The Effect of Example Treatments on Sample Outcomes", "Normal")
    p_title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.first_line_indent = Inches(0)
    p_title.paragraph_format.line_spacing = 1.0
    _format_runs(p_title, 24, True)
    mark(1, p_title)

    # 作者 / 单位块：真实模板必有，且 detect_title_page 靠它认定标题页 ——
    # 缺了这段，24pt 的 Title 会被当成「伪标题」占掉 L1 字号档（实测）。
    for _line in ("Jane Doe",
                  "Department of Example Studies, Example University"):
        p_au = add(_line, "Normal")
        p_au.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_au.paragraph_format.first_line_indent = Inches(0)

    mark(5, add("Abstract", "Heading 1"))
    add("This sample abstract describes the background, method, and principal "
        "findings of a fictional study. It exists only to demonstrate the "
        "formatting rules captured in this template.", "Normal")

    mark(6, add("Keywords: example; formatting; template", "Normal"))

    mark(2, add("1  Introduction", "Heading 1"))
    mark(3, add("1.1  Background", "Heading 2"))
    mark(4, add("1.1.1  Scope of This Template", "Heading 3"))

    # 正文段：同时承载批注 7（body text）、8（margins）、9（paper size）、15（layout）
    p_body = add("Body paragraphs in this template use Times New Roman 12 pt, "
                 "double spacing, a 0.5 inch first-line indent, and justified "
                 "alignment.", "Normal")
    for cid in (7, 8, 9, 15):
        mark(cid, p_body)

    # 表题段：承载批注 14（table captions）+ 10（three-line tables）
    p_tcap = add("Table 1  Sample Three-Line Table", "Normal")
    for cid in (14, 10):
        mark(cid, p_tcap)

    _three_line_table(doc)

    # 表后段：承载批注 11（table spacing after）
    mark(11, add("The paragraph after a table leaves one blank line of spacing "
                 "before it resumes.", "Normal"))

    mark(12, add("References", "Heading 1"))

    for line in (
        "[1] Example, A. (2024). A fictional reference entry. "
        "Example Journal, 12(3), 45-52.",
        "[2] Sample, B. (2023). Another fictional reference entry. Example Press.",
    ):
        p_ref = add(line, "Normal")
        # 悬挂缩进 0.5 英寸（与批注 12 自洽）
        p_ref.paragraph_format.left_indent = Inches(0.5)
        p_ref.paragraph_format.first_line_indent = Inches(-0.5)

    mark(13, add("Figure 1  Sample Figure Caption", "Normal"))

    doc.save(tmp_path)
    return n


def _xml_escape(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def inject_comments(tmp_path, out_path):
    """把 comments.xml 注入 docx，并建立 Content-Type Override 与关系。"""
    body = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    ]
    for cid, text in enumerate(COMMENTS, start=1):
        body.append(
            '<w:comment w:id="%d" w:author="Example University" w:initials="E" '
            'w:date="2026-10-29T10:00:00Z"><w:p><w:r><w:t xml:space="preserve">'
            "%s</w:t></w:r></w:p></w:comment>" % (cid, _xml_escape(text))
        )
    body.append("</w:comments>")
    comments_xml = "".join(body)

    with zipfile.ZipFile(tmp_path, "r") as zin:
        items = {n: zin.read(n) for n in zin.namelist()}

    ct = items["[Content_Types].xml"].decode("utf-8")
    if "/word/comments.xml" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/comments.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
        )
    items["[Content_Types].xml"] = ct.encode("utf-8")

    rels = items["word/_rels/document.xml.rels"].decode("utf-8")
    if "comments.xml" not in rels:
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="rIdCmts900" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
        )
    items["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    items["word/comments.xml"] = comments_xml.encode("utf-8")

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in items.items():
            zout.writestr(n, data)


def make(out_path=None):
    """生成模板并返回输出路径。"""
    out_path = out_path or DEFAULT_OUT
    tmp = out_path + ".tmp.docx"
    anchored = build_with_anchors(tmp)
    inject_comments(tmp, out_path)
    try:
        os.remove(tmp)
    except OSError:
        pass
    return out_path, anchored


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT
    path, anchored = make(out)
    print("Template generated:", path)
    print("Comments defined:", len(COMMENTS), "| anchors attached:", anchored)
