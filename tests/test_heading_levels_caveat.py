# -*- coding: utf-8 -*-
"""标题层级结论只提示、不绑定修正（v2.3.95）。

老板 2026-10-23 拿 v2.3.93 的四份产物核对：APA 学生版 / 专业版**修完复检**
仍然报 `No Heading styles; N levels recognised from manual bold headings
(levels are inferred from font size)`，而报告同时给这一条绑了 `heading` 修正
→ 客户点「修正」后同一个 warn 还在，只会觉得工具坏了。

病根：这一条讲的是**标题层级结构**（有没有套 Word 标题样式），而修正器
**造不出标题、也合并不了层级** —— 标题是客户自己写的内容。工具能做的只有
「按规范的每级格式去排版已有的手排标题」，那是另一条 finding（精确到级）。

所以本文件钉住：
1. 手排标题（无样式）→ `Heading levels` 是 **info**，且**不带 FixAction**；
2. 修完复检不许再出现 warn/fail（能收敛）；
3. 每级**格式**不一致仍然要 warn + 可修（别把可修的活也一起说成不能修）。
"""
from __future__ import annotations

import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.engine.checker import run_check
from src.engine.fixer import fix_docx
from src.engine.questionnaire import target_from_spec

BODY = ("Digital transformation reshapes how manufacturing firms coordinate with "
        "their suppliers, and the evidence base keeps growing across markets.")


def _make(path, *, heading_size=14.0, heading_align=WD_ALIGN_PARAGRAPH.CENTER):
    """手排标题（整段加粗、短、无句读）+ 正文，**不用任何 Heading 样式**。"""
    doc = Document()
    doc.add_paragraph("A Study of Digital Transformation and Resilience")
    doc.add_paragraph(BODY)
    for i, (title, size) in enumerate((("1. Introduction", heading_size),
                                       ("2. Literature Review", heading_size),
                                       ("2.1 Prior Work", 12.0)), start=1):
        h = doc.add_paragraph()
        r = h.add_run(title)
        r.bold = True
        r.font.size = Pt(size)
        h.alignment = heading_align
        doc.add_paragraph(BODY)
    doc.add_paragraph("References")
    doc.add_paragraph("Ivanov, D. (2020). Predicting the impacts.")
    doc.save(str(path))
    return str(path)


def _levels(findings):
    return [f for f in findings
            if f.dimension == "heading" and f.message == "Heading levels"]


def test_manual_bold_headings_are_a_prompt_not_a_violation(tmp_path):
    p = _make(tmp_path / "a.docx")
    prof, findings = run_check(p, target_from_spec("APA-Student"))
    got = _levels(findings)
    assert got, "手排标题必须被认出来并告知（不能装作没看见）"
    assert got[0].severity == "info", got[0].severity
    assert "manual bold" in got[0].actual
    # 🔴 关键：不许给它绑修正 —— 工具造不出标题样式
    assert getattr(got[0], "fix", None) is None, got[0].fix


def test_manual_bold_headings_converge_after_fix(tmp_path):
    """改前 info、改后也 info → 复检不会冒出新 warn/fail（这是老板看到的那条）。"""
    p = _make(tmp_path / "b.docx")
    t = target_from_spec("APA-Student")
    out = str(tmp_path / "b_fixed.docx")
    fix_docx(p, out, t)
    _, f2 = run_check(out, t)
    assert not [f for f in f2 if f.dimension == "heading"
                and f.severity in ("warn", "fail")]
    got = _levels(f2)
    assert got and got[0].severity == "info"


def test_heading_format_mismatch_is_still_warn_and_fixable(tmp_path):
    """每级格式真的不符（这里 L1 故意左对齐，APA 要居中）→ 仍要 warn + 可修 + 收敛。"""
    p = _make(tmp_path / "c.docx", heading_align=WD_ALIGN_PARAGRAPH.LEFT)
    t = target_from_spec("APA-Student")
    _, findings = run_check(p, t)
    fmt = [f for f in findings
           if f.dimension == "heading" and f.message.startswith("Level 1")]
    assert fmt, [f.message for f in findings if f.dimension == "heading"]
    assert fmt[0].severity in ("warn", "fail")
    assert getattr(fmt[0], "fix", None) is not None

    out = str(tmp_path / "c_fixed.docx")
    fix_docx(p, out, t)
    _, f2 = run_check(out, t)
    assert not [f for f in f2 if f.dimension == "heading"
                and f.severity in ("warn", "fail")]
