# -*- coding: utf-8 -*-
"""海外学校模板模式：三线表 / 表后间隙 / 纸张 / 栏数 的「查 → 改 → 复检」回归。

背景
----
Step 3 把批注推出的 ``table_lines`` / ``table_gap_pt`` / ``paper_size`` / ``columns``
接进 ``TargetProfile`` 之后，checker 里早已存在的判定（三线表、表与文字间距、纸型、
栏数）才真正活过来。本文件钉死：

1. 四类问题都能被检查出来，且**每条都绑了 FixAction**（否则报告写"可自动修正"却修不动）；
2. ``fix_docx`` 之后复检必须全部通过，产物实物（纸型/栏数/表格线）也对；
3. 模板没写表后间隙（实测 0）时，按老板 2026-10-29 决策 A 兜底 12pt 并给提示，
   来源标 ``school-default``（归一后仍是 school，所以照样可修）；
4. 模板批注写了间隙值（1 个空行）时用批注值（12pt × 2.0 行距 = 24pt），不触发兜底。
"""

from __future__ import annotations

import zipfile

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from src.engine.checker import run_check
from src.engine.docx_reader import read_docx
from src.engine.fixer import fix_docx
from src.engine.questionnaire import target_from_template_docx

BASE_COMMENTS = [
    "Body text: Times New Roman 12 pt, double-spaced, justified.",
    "Paper size: A4",
    "Layout: two columns",
    "Tables: use three-line tables with no vertical lines.",
]
GAP_COMMENT = "Leave one blank line after each table before the following text."


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
    t = doc.add_table(rows=2, cols=2)
    for r in range(2):
        for c in range(2):
            t.cell(r, c).text = "x"
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + side)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "8")
        borders.append(e)
    t._tbl.tblPr.append(borders)


def _inject(src, out, comments, cids):
    body = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    ]
    for cid, text in zip(cids, comments):
        body.append(
            '<w:comment w:id="%d" w:author="Test" w:initials="T" '
            'w:date="2026-10-29T10:00:00Z"><w:p><w:r><w:t xml:space="preserve">'
            "%s</w:t></w:r></w:p></w:comment>" % (cid, text))
    body.append("</w:comments>")

    with zipfile.ZipFile(src) as zin:
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
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in items.items():
            zout.writestr(n, data)


def _build_template(path, gap_comment=False):
    """模板：Letter / 单栏 / 11pt / 网格表且表旁 0 间距；批注另有要求。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)

    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.0

    before = doc.add_paragraph("Template body text.", style="Normal")
    _grid_table(doc)
    after = doc.add_paragraph("Following text.", style="Normal")

    comments = list(BASE_COMMENTS) + ([GAP_COMMENT] if gap_comment else [])
    for cid, para in ((1, before), (2, before), (3, before), (4, before)):
        _mark_comment(para, cid)
    if gap_comment:
        _mark_comment(after, 5)

    tmp = str(path) + ".tmp.docx"
    doc.save(tmp)
    _inject(tmp, str(path), comments, list(range(1, len(comments) + 1)))
    import os
    os.remove(tmp)
    return path


def _build_paper(path):
    """问题论文：Letter / 单栏 / 11pt 单倍 / 网格表 / 表上下都紧贴正文。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.0
    doc.add_paragraph("Body text before the table.", style="Normal")
    _grid_table(doc)
    doc.add_paragraph("Body text immediately after the table.", style="Normal")
    doc.save(str(path))
    return path


@pytest.fixture()
def template_target(tmp_path):
    p = _build_template(tmp_path / "template.docx")
    return target_from_template_docx(read_docx(str(p)))


CHECK_DIMS = ("table_lines", "table_gap", "paper_size", "columns", "font_size")


def test_all_structural_problems_detected_and_fixable(template_target, tmp_path):
    paper = _build_paper(tmp_path / "paper.docx")
    _, findings = run_check(str(paper), template_target)

    got = {f.dimension: f for f in findings if f.severity in ("warn", "fail")}
    for dim in CHECK_DIMS:
        assert dim in got, "缺少维度告警: %s（现有 %s）" % (dim, sorted(got))
        assert got[dim].fix is not None, \
            "维度 %s 没绑 FixAction —— 报告会说可修但实际修不动" % dim


def test_fix_then_recheck_passes_and_output_is_correct(template_target, tmp_path):
    paper = _build_paper(tmp_path / "paper2.docx")
    out = tmp_path / "paper2_fixed.docx"
    fix_docx(str(paper), str(out), template_target)

    _, findings2 = run_check(str(out), template_target)
    still_bad = [f for f in findings2
                 if f.dimension in CHECK_DIMS and f.severity in ("warn", "fail")]
    assert not still_bad, "修完复检仍报: %s" % [(f.dimension, f.actual) for f in still_bad]

    prof = read_docx(str(out))
    assert prof.paper_size == "A4"
    assert prof.columns == 2
    assert all(s.get("three_line") for s in (prof.table_line_states or []))


def test_table_gap_fallback_when_template_silent(tmp_path):
    """决策 A：模板没写表后间隙且实测 0 → 兜底 12pt，来源 school-default，并给提示。"""
    p = _build_template(tmp_path / "tpl_silent.docx", gap_comment=False)
    tgt = target_from_template_docx(read_docx(str(p)))

    assert tgt.table_gap_pt == 12.0
    assert tgt.resolution["table_gap_pt"] == "school-default"
    assert any("Table spacing" in n for n in tgt.notes)


def test_table_gap_from_comment_wins_over_fallback(tmp_path):
    """批注写了「一个空行」→ 12pt × 2.0 行距 = 24pt，且来源是 school-comment。"""
    p = _build_template(tmp_path / "tpl_gap.docx", gap_comment=True)
    tgt = target_from_template_docx(read_docx(str(p)))

    assert tgt.table_gap_pt == pytest.approx(24.0)
    assert tgt.resolution["table_gap_pt"] == "school-comment"
    assert not any("Table spacing: your template does not state" in n for n in tgt.notes)


def _build_template_with_caption(path, gap_comment=False):
    """模板：Letter/单栏/11pt/网格表，表上方有 caption 段落（继承正文、段前=0）。

    用于回归 caption 段前实测为 0 是否会被误判成显式「0pt 要求」——
    那样会与 table_gap 修正器写进表题段前的空隙互相打架。
    """
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.0

    before = doc.add_paragraph("Template body text.", style="Normal")
    doc.add_paragraph("Table 1 Sample three-line table", style="Normal")  # caption，继承 0 段前
    _grid_table(doc)
    after = doc.add_paragraph("Following text.", style="Normal")

    comments = list(BASE_COMMENTS) + ([GAP_COMMENT] if gap_comment else [])
    for cid in range(1, len(comments) + 1):
        _mark_comment(before, cid)
    if gap_comment:
        _mark_comment(after, len(comments))

    tmp = str(path) + ".tmp.docx"
    doc.save(tmp)
    _inject(tmp, str(path), comments, list(range(1, len(comments) + 1)))
    import os
    os.remove(tmp)
    return path


def test_caption_space_before_measured_zero_not_a_requirement(tmp_path):
    """Regression（Issue B）：模板 caption 段前实测为 0（继承正文样式）时，
    target 不应写成 caption_space_before_pt=0.0，否则会与 table_gap 修正器写进
    表题段前的空隙互相打架，点完「全部修正」报告仍挂红项。"""
    p = _build_template_with_caption(tmp_path / "tpl_cap.docx", gap_comment=True)
    tgt = target_from_template_docx(read_docx(str(p)))

    assert tgt.caption_space_before_pt is None, \
        "caption 段前实测 0 被误当成显式 0pt 要求"
    assert tgt.caption_space_after_pt is None
    # 冲突的另一方（table_gap）仍正常带出
    assert tgt.table_gap_pt == pytest.approx(24.0)


def test_caption_gap_no_recheck_failure(tmp_path):
    """端到端：带 caption + 表格的论文按模板修正后，复检不得出现 caption_space_before 红项。"""
    _build_template_with_caption(tmp_path / "tpl_cap2.docx", gap_comment=True)
    tgt = target_from_template_docx(read_docx(str(tmp_path / "tpl_cap2.docx")))

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)
    doc.add_paragraph("Body before table.", style="Normal")
    doc.add_paragraph("Table 1 My table", style="Normal")
    _grid_table(doc)
    doc.add_paragraph("Body right after table.", style="Normal")
    paper = tmp_path / "paper_cap.docx"
    doc.save(str(paper))

    out = tmp_path / "paper_cap_fixed.docx"
    fix_docx(str(paper), str(out), tgt)
    _, findings = run_check(str(out), tgt)
    bad = [f for f in findings
           if f.dimension == "caption_space_before" and f.severity in ("warn", "fail")]
    assert not bad, "修正后 caption_space_before 仍报: %s" % \
        [(f.severity, f.actual) for f in bad]
