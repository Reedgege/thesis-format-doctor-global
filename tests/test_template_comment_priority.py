# -*- coding: utf-8 -*-
"""模板批注优先于实测格式（海外「学校模板」模式）的回归测试。

守护点
------
1. **批注 > 实测**：模板自身是 16pt / Letter / 单栏 / 满格表，批注却要求
   18pt / A4 / 双栏 / 三线表 / APA → 目标画像必须取批注值；
2. **来源标注**：批注派生字段的 resolution 必须是 ``school-comment``
   （与实测的 ``school`` 区分开，报告才能说清「批注说的」还是「量出来的」）；
3. **参考文献流派**：显式下拉 > 批注推断；只从批注推断，不从实测猜；
4. **三线表接线**：模板模式必须把 table_lines 传进 TargetProfile
   （此前根本没传 → 勾了学校模板也永远不查三线表）；
5. **无批注退化**：完全退回实测，行为与改动前一致。
"""

from __future__ import annotations

import zipfile

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from src.engine.docx_reader import read_docx
from src.engine.questionnaire import target_from_template_docx

COMMENTS = [
    "Heading 1: Times New Roman 18 pt, bold, centered.",
    "Body text: Times New Roman 12 pt, double-spaced, justified.",
    "Paper size: A4",
    "Layout: two columns",
    "Tables: use three-line tables with no vertical lines.",
    "References: APA 7th edition.",
]


def _mark_comment(paragraph, cid):
    p = paragraph._p
    cs = OxmlElement("w:commentRangeStart")
    cs.set(qn("w:id"), str(cid))
    ppr = p.find(qn("w:pPr"))
    if ppr is not None:
        ppr.addnext(cs)
    else:
        p.insert(0, cs)
    ce = OxmlElement("w:commentRangeEnd")
    ce.set(qn("w:id"), str(cid))
    p.append(ce)
    r = OxmlElement("w:r")
    cref = OxmlElement("w:commentReference")
    cref.set(qn("w:id"), str(cid))
    r.append(cref)
    p.append(r)


def _grid_table(doc):
    """满格表（有竖线、有内部横线）—— 明确不是三线表。"""
    t = doc.add_table(rows=2, cols=2)
    for r in range(2):
        for c in range(2):
            t.cell(r, c).text = "x"
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + side)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "8")
        borders.append(e)
    tblPr.append(borders)


def _inject_comments(tmp_path, out_path, comments):
    body = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    ]
    for cid, text in enumerate(comments, start=1):
        body.append(
            '<w:comment w:id="%d" w:author="Test" w:initials="T" '
            'w:date="2026-10-29T10:00:00Z"><w:p><w:r><w:t xml:space="preserve">'
            "%s</w:t></w:r></w:p></w:comment>" % (cid, text))
    body.append("</w:comments>")

    with zipfile.ZipFile(tmp_path) as zin:
        items = {n: zin.read(n) for n in zin.namelist()}
    ct = items["[Content_Types].xml"].decode("utf-8")
    if "/word/comments.xml" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/comments.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')
    items["[Content_Types].xml"] = ct.encode("utf-8")
    rels = items["word/_rels/document.xml.rels"].decode("utf-8")
    if "comments.xml" not in rels:
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="rIdCmts900" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')
    items["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    items["word/comments.xml"] = "".join(body).encode("utf-8")
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in items.items():
            zout.writestr(n, data)


def _build(path, with_comments=True):
    """模板自身：Letter / 单栏 / 正文 11pt 单倍左对齐 / H1 16pt 左对齐 / 满格表。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)

    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.0
    doc.styles["Heading 1"].font.name = "Times New Roman"
    doc.styles["Heading 1"].font.size = Pt(16)
    doc.styles["Heading 1"].font.bold = True
    doc.styles["Heading 1"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    h1 = doc.add_paragraph("1  Introduction", style="Heading 1")
    body = doc.add_paragraph("Body text for the priority probe.", style="Normal")
    _grid_table(doc)

    if with_comments:
        for cid, para in ((1, h1), (2, body), (3, body), (4, body), (5, body), (6, body)):
            _mark_comment(para, cid)

    tmp = str(path) + ".tmp.docx"
    doc.save(tmp)
    if with_comments:
        _inject_comments(tmp, str(path), COMMENTS)
        import os
        os.remove(tmp)
    else:
        import os
        os.replace(tmp, str(path))
    return path


@pytest.fixture()
def conflict_target(tmp_path):
    p = _build(tmp_path / "conflict.docx", with_comments=True)
    prof = read_docx(str(p))
    return prof, target_from_template_docx(prof)


def test_comment_beats_measured_for_page_and_headings(conflict_target):
    prof, tgt = conflict_target

    # 前置：实测确实是 16pt / 11pt / 单倍 / 左对齐
    assert prof.heading_formats[1].size_pt == 16.0
    assert prof.font_size_pt == 11.0

    # 批注值必须压过实测
    assert tgt.heading_formats[1]["size_pt"] == 18.0
    assert tgt.heading_formats[1]["alignment"] == "center"
    assert tgt.page["font_size_pt"] == 12.0
    assert tgt.page["line_spacing"] == 2.0
    assert tgt.page["body_alignment"] == "justify"
    assert tgt.resolution["heading_formats"] == "school-comment"
    assert tgt.resolution["font_size_pt"] == "school-comment"


def test_comment_beats_measured_for_structural_dims(conflict_target):
    prof, tgt = conflict_target
    assert prof.paper_size == "Letter" and prof.columns == 1

    assert tgt.paper_size == "A4"
    assert tgt.columns == 2
    assert tgt.resolution["paper_size"] == "school-comment"
    assert tgt.resolution["columns"] == "school-comment"


def test_table_lines_wired_from_comment(conflict_target):
    _, tgt = conflict_target
    # 模板自己的表是满格表 → 实测推断不出三线；批注要求三线 → 必须以批注为准
    assert tgt.table_lines == "three_line"
    assert tgt.resolution["table_lines"] == "school-comment"


def test_reference_style_from_comment(conflict_target):
    _, tgt = conflict_target
    assert tgt.reference_style == "APA"      # REFERENCE_STYLE_KEYS 为大写
    assert tgt.resolution["reference_style"] == "school-comment"
    assert any("Word comments" in n for n in tgt.notes)


def test_explicit_dropdown_beats_comment(conflict_target):
    prof, _ = conflict_target
    tgt = target_from_template_docx(prof, reference_style="MLA")
    assert tgt.reference_style == "MLA", "显式下拉必须压过批注推断"


def test_no_comments_falls_back_to_measured(tmp_path):
    p = _build(tmp_path / "plain.docx", with_comments=False)
    prof = read_docx(str(p))
    tgt = target_from_template_docx(prof)

    assert tgt.heading_formats[1]["size_pt"] == 16.0
    assert tgt.page["font_size_pt"] == 11.0
    assert tgt.paper_size == "Letter"
    assert tgt.columns == 1
    assert tgt.table_lines is None, "满格表模板不该被判成三线表"
    assert tgt.reference_style is None
    assert tgt.resolution.get("heading_formats") == "school"
    assert not any("Word comments" in n for n in tgt.notes)
