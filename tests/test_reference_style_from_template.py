# -*- coding: utf-8 -*-
"""参考文献流派（下拉 > 批注 > 不查）在海外学校模板模式下的回归。

守护点
------
1. 模板批注写 "References: APA 7th edition." → ``target.reference_style == "APA"``，
   悬挂缩进由 APA 规范补 0.5″；
2. ``resolve_spec()`` 能按 reference_style 取到规范，但 ``spec_is_baseline()`` 必须是
   **False** —— 否则 APA 的页面/标题格式会外溢到学校模板通道（实测事故：论文大标题被
   改成 ``I. THE IMPACT OF…``）；
3. 无悬挂缩进的论文 → checker 报 ``reference_hanging_indent`` 并**绑 FixAction**；
4. ``fix_docx`` 后复检通过，产物悬挂缩进确实为 0.5″；
5. 显式下拉压过批注；两者都没有时**不查**参考文献（reference_style=None）；
6. **IEEE 是编号制**（下拉三项 APA / MLA / IEEE 里的第三项）：要求 numbered=True、
   悬挂 0.25″，未编号条目要报出来并补 [n] 前缀 —— 与 APA/MLA 的查改路径不同，单独守护。
"""

from __future__ import annotations

import os

import pytest
from docx import Document
from docx.shared import Inches, Pt

from src.engine.checker import run_check
from src.engine.docx_reader import read_docx
from src.engine.fixer import fix_docx
from src.engine.questionnaire import target_from_template_docx
from src.engine.specs import resolve_spec, spec_is_baseline

from _docx_fixtures import inject_comments, mark_comment

APA_COMMENT = "References: APA 7th edition."


def _build_template(path, with_comment=True):
    doc = Document()
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(12)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    body = doc.add_paragraph("Template body text.", style="Normal")
    if with_comment:
        mark_comment(body, 1)
    tmp = str(path) + ".tmp.docx"
    doc.save(tmp)
    if with_comment:
        inject_comments(tmp, str(path), [(1, APA_COMMENT)])
        os.remove(tmp)
    else:
        os.replace(tmp, str(path))
    return path


def _build_paper(path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(12)
    doc.add_paragraph("References", style="Heading 1")
    for line in (
        "Example, A. (2024). A study of things. Journal of Things, 12(3), 45-52.",
        "Sample, B. (2023). Another study of things. Example Press.",
    ):
        p = doc.add_paragraph(line, style="Normal")
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.first_line_indent = Inches(0)
    doc.save(str(path))
    return path


@pytest.fixture()
def apa_target(tmp_path):
    p = _build_template(tmp_path / "tpl.docx")
    return read_docx(str(p)), target_from_template_docx(read_docx(str(p)))


def test_comment_sets_style_and_spec(apa_target):
    _, tgt = apa_target
    assert tgt.reference_style == "APA"
    assert tgt.resolution["reference_style"] == "school-comment"
    assert abs(float(tgt.reference_hanging_indent_in) - 0.5) < 1e-6

    assert resolve_spec(tgt).key.startswith("APA")
    # 学校模板通道绝不能被当成「内置规范基线」，否则页面/标题会被 APA 覆盖
    assert spec_is_baseline(tgt) is False


def test_hanging_indent_checked_and_fix_bound(apa_target, tmp_path):
    _, tgt = apa_target
    paper = _build_paper(tmp_path / "paper.docx")
    prof = read_docx(str(paper))
    assert prof.reference_entries, "前置：参考文献条目必须被识别"

    _, findings = run_check(str(paper), tgt)
    hang = [f for f in findings if f.dimension == "reference_hanging_indent"]
    assert hang, "APA 要求悬挂缩进 0.5″，必须报出来"
    assert hang[0].severity in ("warn", "fail")
    assert hang[0].fix is not None, "没绑 FixAction → 报告说可修却修不动"


def test_fix_applies_hanging_indent(apa_target, tmp_path):
    _, tgt = apa_target
    paper = _build_paper(tmp_path / "paper2.docx")
    out = tmp_path / "paper2_fixed.docx"
    fix_docx(str(paper), str(out), tgt)

    _, findings = run_check(str(out), tgt)
    bad = [f for f in findings
           if f.dimension == "reference_hanging_indent" and f.severity in ("warn", "fail")]
    assert not bad, "修完复检仍报: %s" % [(f.dimension, f.actual) for f in bad]
    assert abs(float(read_docx(str(out)).reference_hanging_indent_in or 0) - 0.5) < 0.02


def test_dropdown_overrides_comment(apa_target):
    prof, _ = apa_target
    tgt = target_from_template_docx(prof, reference_style="MLA")
    assert tgt.reference_style == "MLA"
    assert resolve_spec(tgt).key.startswith("MLA")


def test_no_style_means_no_reference_check(tmp_path):
    p = _build_template(tmp_path / "tpl_plain.docx", with_comment=False)
    tgt = target_from_template_docx(read_docx(str(p)))
    assert tgt.reference_style is None
    assert tgt.reference_hanging_indent_in is None

    paper = _build_paper(tmp_path / "paper3.docx")
    _, findings = run_check(str(paper), tgt)
    assert not [f for f in findings if f.dimension == "reference_hanging_indent"], \
        "没有流派时不该判悬挂缩进"


# ---------------------------------------------------------------------------
# IEEE（下拉三项里的第三项）：编号制，查改路径与 APA/MLA 不同
# ---------------------------------------------------------------------------
IEEE_COMMENT = "References: IEEE style, numbered [1]."


def _build_ieee_template(path):
    doc = Document()
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(10)
    body = doc.add_paragraph("Template body text.", style="Normal")
    mark_comment(body, 1)
    tmp = str(path) + ".tmp.docx"
    doc.save(tmp)
    inject_comments(tmp, str(path), [(1, IEEE_COMMENT)])
    os.remove(tmp)
    return path


def _build_ieee_paper(path):
    """IEEE 风格的条目内容，但**不带 [n] 编号**（要被查出来并补上）。"""
    doc = Document()
    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(10)
    doc.add_paragraph("References", style="Heading 1")
    for line in (
        'J. A. Smith, "Title one," J. Name, vol. 1, no. 2, pp. 3-4, 2020.',
        'B. Sample, "Title two," J. Name, vol. 2, no. 3, pp. 5-6, 2021.',
    ):
        p = doc.add_paragraph(line, style="Normal")
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.first_line_indent = Inches(0)
    doc.save(str(path))
    return path


def test_ieee_comment_sets_numbered_and_spec(tmp_path):
    p = _build_ieee_template(tmp_path / "ieee.docx")
    tgt = target_from_template_docx(read_docx(str(p)))

    assert tgt.reference_style == "IEEE"
    assert tgt.reference_numbered is True, "IEEE 是顺序编码制，numbered 必须为 True"
    assert resolve_spec(tgt).key.startswith("IEEE")
    assert spec_is_baseline(tgt) is False


def test_ieee_numbering_checked_and_fixed(tmp_path):
    tpl = _build_ieee_template(tmp_path / "ieee2.docx")
    tgt = target_from_template_docx(read_docx(str(tpl)))
    paper = _build_ieee_paper(tmp_path / "ieee_paper.docx")

    _, findings = run_check(str(paper), tgt)
    # checker 把编号结论挂在 "reference" 维度上，修复动作的 dimension 才是 reference_numbered
    num = [f for f in findings
           if f.dimension == "reference" and f.severity in ("warn", "fail")]
    assert num, "IEEE 编号制：未编号条目必须报出来"
    assert any(getattr(getattr(f, "fix", None), "dimension", None) == "reference_numbered"
               for f in num), "没绑 reference_numbered 修复动作 → 修不动"

    out = tmp_path / "ieee_paper_fixed.docx"
    fix_docx(str(paper), str(out), tgt)
    prof = read_docx(str(out))
    assert prof.reference_entries, "修完应仍能读到条目"
    assert all(str(e).strip().startswith("[") for e in prof.reference_entries), \
        "IEEE 修正后条目应带 [n] 前缀"

    _, findings2 = run_check(str(out), tgt)
    assert not [f for f in findings2
                if f.dimension == "reference" and f.severity in ("warn", "fail")]
