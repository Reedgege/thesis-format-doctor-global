# -*- coding: utf-8 -*-
"""报告必须说清「每个维度是谁说了算」：模板批注 / 模板实测 / 通用兜底 / 内置规范。

为什么单独钉
------------
海外学校模板模式现在有四个来源同时存在：``school``（实测）、``school-comment``
（批注）、``school-default``（模板沉默时的通用兜底）、``spec``（内置规范补位）。
如果报告把它们混成一句 "School Template"，客户就分不清「这是学校要求的」还是
「我们自己按通用规则加的」——尤其兜底项，不标清楚就是替客户拍板。
"""

from __future__ import annotations

import os

import pytest
from docx import Document
from docx.shared import Inches, Pt

from src.engine.checker import run_check
from src.engine.docx_reader import read_docx
from src.engine.questionnaire import target_from_template_docx
from src.engine.report import build_report

from _docx_fixtures import grid_table, inject_comments, mark_comment

COMMENTS = [
    "Body text: Times New Roman 12 pt, double-spaced, justified.",
    "Paper size: A4",
    "Layout: two columns",
    "Tables: use three-line tables with no vertical lines.",
    # 这一条让「内置规范补位」也出现在来源表里（APA 的悬挂缩进由 spec 给出）
    "References: APA 7th edition.",
]


def _build_template(path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.top_margin = sec.bottom_margin = Inches(1)
    sec.left_margin = sec.right_margin = Inches(1)
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].paragraph_format.line_spacing = 1.0
    body = doc.add_paragraph("Template body text.", style="Normal")
    grid_table(doc)
    for cid in range(1, len(COMMENTS) + 1):
        mark_comment(body, cid)
    tmp = str(path) + ".tmp.docx"
    doc.save(tmp)
    inject_comments(tmp, str(path), list(enumerate(COMMENTS, start=1)))
    os.remove(tmp)
    return path


def _build_paper(path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_paragraph("Body text before the table.", style="Normal")
    grid_table(doc)
    doc.add_paragraph("Body text after the table.", style="Normal")
    doc.save(str(path))
    return path


@pytest.fixture()
def report(tmp_path):
    tpl = _build_template(tmp_path / "tpl.docx")
    target = target_from_template_docx(read_docx(str(tpl)))
    paper = _build_paper(tmp_path / "paper.docx")
    prof, findings = run_check(str(paper), target)
    return build_report(str(paper), target, prof, findings)


def test_meta_source_distinguishes_comment_from_measured(report):
    # 批注派生的维度必须以 "School Template (comments)" 出现，而不是笼统的 "School Template"
    assert "School Template (comments)" in report.meta["source"]


def test_resolution_table_labels_each_origin(report):
    labels = {origin for _dim, origin in report.resolution}
    assert "School Template (comments)" in labels, labels
    assert "General default (template silent)" in labels, \
        "表后间隙兜底必须标成「通用兜底」，不能冒充模板要求；现有：%s" % labels
    assert "General Spec" in labels, "悬挂缩进等由内置规范补位的也要标 spec"


def test_notes_explain_comments_and_fallback(report):
    joined = " | ".join(report.notes)
    assert "Word comments" in joined, "应说明有几条规则来自模板批注"
    assert "Table spacing" in joined and "general default" in joined.lower()


def test_markdown_renders_notes_section(report):
    md = report.markdown
    assert "Dimension Resolution" in md
    assert "School Template (comments)" in md
    assert "Notes" in md and "Word comments" in md
