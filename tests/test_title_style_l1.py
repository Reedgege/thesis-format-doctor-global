# -*- coding: utf-8 -*-
"""回归：Word 内置「Title」样式不得占掉 L1 标题字号档。

背景（2026-10-29 实测真 bug）
--------------------------------
``_heading_level_from_name`` 故意把 Title / 题目 等「整段标题样式」判为 **1 级**
（保护主标题不被压成正文，见 test_review_fixes.py::test_title_style_is_heading）。
但主采样循环 ``heading_formats`` 早期没有排除前置块，于是：

    论文大标题（Title 样式 24pt）+ 真正的 Heading 1（18pt）
    → heading_formats = {1: 24pt}   ← 真 H1 被 seen_levels 直接跳过
    → target_from_template_docx 的 L1 = 24pt
    → 修正把 18pt 章节标题放大成 24pt（比漏检更伤）

修法：``heading_formats`` 主采样排除 ``skip`` / ``_is_front_style`` /
``doc_title_index``；``detect_title_page`` 与 ``doc_title_index`` 不再对
「Title 样式首段」早退。
"""

from __future__ import annotations

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from src.engine.docx_reader import doc_title_index, read_docx


def _doc_with_title_style(path, *, centered=False):
    doc = Document()
    if "Title" not in [s.name for s in doc.styles]:
        pytest.skip("当前 python-docx 默认模板无 Title 样式")

    st = doc.styles["Title"]
    st.font.size = Pt(24)
    st.font.bold = True
    doc.styles["Heading 1"].font.size = Pt(18)

    title = doc.add_paragraph("Paper Title Of The Study", style="Title")
    if centered:
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Jane Doe")
    doc.add_paragraph("Abstract", style="Heading 1")
    doc.add_paragraph("Abstract text.")
    doc.add_paragraph("1  Introduction", style="Heading 1")
    doc.add_paragraph("Body text.")
    doc.save(str(path))
    return path


def test_title_style_does_not_hijack_l1(tmp_path):
    """论文大标题（Title 样式）不得顶替真正的 Heading 1 样本。"""
    p = _doc_with_title_style(tmp_path / "title_style.docx")
    prof = read_docx(str(p))
    assert 1 in prof.heading_formats, "必须采到一级标题"
    assert prof.heading_formats[1].size_pt == pytest.approx(18.0), (
        "一级标题样本被论文大标题（24pt）顶掉了：%r"
        % (prof.heading_formats[1].size_pt,)
    )


def test_doc_title_detected_with_title_style(tmp_path):
    """Title 样式的首段仍应能被认成「论文大标题」（居中 + 加粗 + 全篇最大）。"""
    p = _doc_with_title_style(tmp_path / "title_style_centered.docx", centered=True)
    doc = Document(str(p))
    assert doc_title_index(doc) == 0, "Title 样式的论文大标题未被识别"
