# -*- coding: utf-8 -*-
"""标题结构识别 + 摘要标题 + Keywords 行（v2.3.97）。

老板 2026-10-23 两问：
  「apa 的摘要部分标题是否需要居中？二级标题是否加粗？」→ 都是产品真缺失，修。
  「各层级标题修改是修正格式重要的一部分，要控制好」→ 下面几组反向用例就是
  「控制好」的护栏：宁可少认一条，也绝不把正文段落当标题刷格式。

四条铁律级约束：
  ① 编号结构行必须过四道闸门（像标题 + 序列成立 + 层级自顶向下 + 层数上限）；
  ② 前置/后置材料标题（Abstract / References …）不进层级仲裁，否则会把真章节挤深一级；
  ③ 摘要标题与 Keywords 行的格式各有唯一主人（abstract_heading / keywords 维度）；
  ④ MLA / IEEE 不规定这两项 → 一个字都不说不改。
"""

from __future__ import annotations

import os
import sys

import pytest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.checker import run_check
from src.engine.docx_reader import (
    split_body_soft_breaks, structural_heading_map, find_abstract_section,
)
from src.engine.fixer import fix_docx, _assign_heading_levels
from src.engine.questionnaire import target_from_dict, target_from_spec


# ---------------------------------------------------------------------------
# 夹具
# ---------------------------------------------------------------------------

def _bold_para(doc, text, size=12, align=None, bold=True):
    p = doc.add_paragraph(text)
    if align is not None:
        p.alignment = align
    for r in p.runs:
        r.bold = bold
        r.font.size = Pt(size)
    return p


def _numbered_doc(path, chapters=("1. Introduction", "2. Method", "3. Results")):
    doc = Document()
    doc.add_paragraph("A Study of Things")
    doc.add_paragraph("Body opening paragraph with a normal sentence in it.")
    for c in chapters:
        doc.add_paragraph(c)          # 不加粗的编号小节行（旧实现两头落空）
        doc.add_paragraph("Paragraph under %s that runs long enough to be body text."
                          % c)
    doc.save(str(path))
    return str(path)


def _apa_doc(path, ah_align=WD_ALIGN_PARAGRAPH.LEFT, ah_bold=True,
             kw_indent=0.0, kw_italic=None):
    """摘要标题 + 摘要正文 + Keywords 行（Keywords 故意挤在一个 run 里）。"""
    doc = Document()
    doc.add_paragraph("Sleep Quality and Cognitive Performance")
    ah = doc.add_paragraph("Abstract")
    ah.alignment = ah_align
    for r in ah.runs:
        r.bold = ah_bold
        r.font.size = Pt(12)
    doc.add_paragraph("Sleep quality affects cognitive performance in young adults.")
    kw = doc.add_paragraph("Keywords: sleep quality, cognitive performance, young adults")
    kw.paragraph_format.first_line_indent = Inches(kw_indent)
    for r in kw.runs:
        r.italic = kw_italic
        r.font.size = Pt(12)
    doc.add_paragraph("Introduction")
    doc.save(str(path))
    return str(path)


def _find(findings, dim):
    return [f for f in findings if f.dimension == dim]


def _struct(path):
    doc = Document(path)
    split_body_soft_breaks(doc)
    paras = list(doc.paragraphs)
    return paras, structural_heading_map(paras)


# ---------------------------------------------------------------------------
# 一、结构标题识别：四道闸门
# ---------------------------------------------------------------------------

def test_numbered_sections_are_recognised(tmp_path):
    """「1. Introduction」这种不加粗的编号小节行必须认出来（旧实现两头落空）。"""
    src = _numbered_doc(tmp_path / "n.docx")
    paras, lvl = _struct(src)
    by_text = {(p.text or "").strip(): lvl.get(p._p) for p in paras}
    assert by_text["1. Introduction"] == 1
    assert by_text["2. Method"] == 1
    assert by_text["3. Results"] == 1


def test_sentence_starting_with_a_number_is_not_a_heading(tmp_path):
    """闸门①：以编号开头、但以句号结尾的长句是正文，不是标题。"""
    doc = Document()
    doc.add_paragraph("1. Introduction")
    doc.add_paragraph("2. Method")
    doc.add_paragraph("3. Results show that the model fits the data well, and the "
                      "coefficients stay stable across every specification we tried.")
    doc.save(str(tmp_path / "s.docx"))
    paras, lvl = _struct(tmp_path / "s.docx")
    long_p = next(p for p in paras if (p.text or "").startswith("3. Results show"))
    assert lvl.get(long_p._p) is None


def test_numbering_must_start_at_one(tmp_path):
    """闸门②：编号序列不从 1 开始（3./4./5.）就不是章节编号 → 一概不认。"""
    doc = Document()
    doc.add_paragraph("3. Something")
    doc.add_paragraph("4. Something else")
    doc.add_paragraph("5. A third one")
    doc.save(str(tmp_path / "n3.docx"))
    _paras, lvl = _struct(tmp_path / "n3.docx")
    assert lvl == {}, "不从 1 开始的编号序列不得被当成章节层级"


def test_single_numbered_line_is_not_a_heading(tmp_path):
    """闸门③：孤零零一条编号不算层级（至少要两项成序列）。"""
    doc = Document()
    doc.add_paragraph("1. Introduction")
    doc.save(str(tmp_path / "one.docx"))
    _paras, lvl = _struct(tmp_path / "one.docx")
    assert lvl == {}


def test_absolute_document_title_is_not_a_section(tmp_path):
    """论文大标题（居中加粗、全篇最大字号）不是章节标题。"""
    doc = Document()
    t = doc.add_paragraph("Digital Transformation and Supply Chain Resilience")
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in t.runs:
        r.bold = True
        r.font.size = Pt(18)
    _bold_para(doc, "Introduction", size=14)
    _bold_para(doc, "Method", size=14)
    doc.save(str(tmp_path / "t.docx"))
    d = Document(str(tmp_path / "t.docx"))
    split_body_soft_breaks(d)
    lvl = _assign_heading_levels(d, target_from_spec("APA-Student"))
    by_text = {(p.text or "").strip(): lvl.get(p._p) for p in d.paragraphs}
    assert by_text["Digital Transformation and Supply Chain Resilience"] is None
    assert by_text["Introduction"][0] == 1
    assert by_text["Method"][0] == 1


def test_matter_titles_do_not_steal_a_heading_level(tmp_path):
    """② Abstract 加粗居中也不许抢层级 —— 否则 1. Introduction 被挤成 L3（套上加粗斜体）。"""
    doc = Document()
    doc.add_paragraph("A Study of Things")
    _bold_para(doc, "Abstract", size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph("Abstract body text that summarizes the whole study briefly.")
    _bold_para(doc, "1. Introduction", size=12, bold=False)
    doc.add_paragraph("Body under introduction, long enough to look like body text.")
    _bold_para(doc, "2. Method", size=12, bold=False)
    doc.add_paragraph("Body under method, long enough to look like body text.")
    doc.save(str(tmp_path / "m.docx"))
    paras, lvl = _struct(tmp_path / "m.docx")
    by_text = {(p.text or "").strip(): lvl.get(p._p) for p in paras}
    assert by_text["Abstract"] is None, "Abstract 是前置材料标题，不属于章节层级"
    assert by_text["1. Introduction"] == 1
    assert by_text["2. Method"] == 1


def test_references_still_gets_top_level_formatting(tmp_path, legacy_text_transform):
    """④ References 不进层级仲裁，但**仍要套 L1 格式**（IEEE 要求 REFERENCES 全大写）。

    编号/全大写属「改文字」，产品默认锁死（v2.3.84）；这里验的是底层能力仍在。
    """
    doc = Document()
    for text in ("Introduction", "Method", "References"):
        _bold_para(doc, text, size=12)
    doc.add_paragraph("1. Anderson, K. (2019). Sleep duration and memory. JSR, 12(3), 45-67.")
    doc.save(str(tmp_path / "r.docx"))
    out = tmp_path / "r_out.docx"
    fix_docx(str(tmp_path / "r.docx"), str(out), target_from_spec("IEEE"))
    texts = [p.text for p in Document(str(out)).paragraphs]
    # v2.3.100（裁定 Q-1 / P-2）：IEEE 标题的**字符类**（全大写）只提示不执行
    # → 参考文献标题保持作者原样 "References"，不再被改成 "REFERENCES"。
    assert "References" in texts
    assert not any(t.startswith("III. REFERENCES") for t in texts)


# ---------------------------------------------------------------------------
# 二、APA 摘要标题：居中 + 加粗
# ---------------------------------------------------------------------------

def test_apa_flags_and_fixes_abstract_heading(tmp_path):
    src = _apa_doc(tmp_path / "a.docx", ah_align=WD_ALIGN_PARAGRAPH.LEFT, ah_bold=False)
    tgt = target_from_spec("APA-Student")
    _prof, findings = run_check(src, tgt)
    got = _find(findings, "abstract_heading")
    assert got and got[0].severity == "warn"
    assert "centered" in got[0].expected and "bold" in got[0].expected
    assert getattr(got[0], "fix", None) is not None, "摘要标题格式必须可修"

    out = tmp_path / "a_out.docx"
    fix_docx(src, str(out), tgt)
    ah = next(p for p in Document(str(out)).paragraphs
              if (p.text or "").strip() == "Abstract")
    assert ah.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert all(r.bold for r in ah.runs if (r.text or "").strip())
    _p2, again = run_check(str(out), tgt)
    assert _find(again, "abstract_heading")[0].severity == "pass"


def test_apa_flags_level_two_headings_not_bold(tmp_path):
    """二级标题（左对齐、不加粗）也要被认出来并加粗。"""
    doc = Document()
    doc.add_paragraph("A Study of Things")
    for text in ("1. Introduction", "1.1 Background", "1.2 Motivation", "2. Method"):
        _bold_para(doc, text, size=12, bold=False)
        doc.add_paragraph("Body under %s, long enough to look like body text." % text)
    doc.save(str(tmp_path / "h2.docx"))
    out = tmp_path / "h2_out.docx"
    fix_docx(str(tmp_path / "h2.docx"), str(out), target_from_spec("APA-Student"))
    l211 = next(p for p in Document(str(out)).paragraphs
                if (p.text or "").strip() == "1.1 Background")
    assert all(r.bold for r in l211.runs if (r.text or "").strip()), \
        "APA Level 2 要求加粗"


# ---------------------------------------------------------------------------
# 三、Keywords 行：标签斜体 + 缩进 0.5″
# ---------------------------------------------------------------------------

def test_apa_flags_and_fixes_keywords(tmp_path):
    src = _apa_doc(tmp_path / "k.docx", kw_indent=0.0, kw_italic=None)
    tgt = target_from_spec("APA-Student")
    _prof, findings = run_check(src, tgt)
    got = _find(findings, "keywords")
    assert got and got[0].severity == "warn"
    assert getattr(got[0], "fix", None) is not None

    out = tmp_path / "k_out.docx"
    fix_docx(src, str(out), tgt)
    kw = next(p for p in Document(str(out)).paragraphs
              if (p.text or "").lower().startswith("keywords"))
    assert abs(kw.paragraph_format.first_line_indent.inches - 0.5) < 0.01
    # 文字一字不改，只把标签与词条切开：标签斜体、词条不斜体
    assert kw.text == ("Keywords: sleep quality, cognitive performance, young adults")
    runs = [r for r in kw.runs if r.text]
    assert runs[0].text == "Keywords:" and runs[0].font.italic is True
    assert any(not r.font.italic for r in runs[1:]), "词条不该被一起斜体"

    _p2, again = run_check(str(out), tgt)
    assert _find(again, "keywords")[0].severity == "pass"


@pytest.mark.parametrize("spec_key", ["MLA", "IEEE"])
def test_other_specs_say_nothing_about_abstract_heading_or_keywords(tmp_path, spec_key):
    """MLA/IEEE 不规定这两项 → 一个结论都不给、一个字节都不改。"""
    src = _apa_doc(tmp_path / "o.docx", kw_indent=0.0, kw_italic=None)
    tgt = target_from_spec(spec_key)
    _prof, findings = run_check(src, tgt)
    assert not _find(findings, "abstract_heading")
    assert not _find(findings, "keywords")

    out = tmp_path / ("o_%s.docx" % spec_key)
    res = fix_docx(src, str(out), tgt)
    applied = res.get("applied") or {}
    assert "abstract_heading" not in applied
    assert "keywords" not in applied


def test_channel_silence_means_no_finding(tmp_path):
    """三通道都没提 → 闭嘴（不是给个 info 敷衍）。"""
    src = _apa_doc(tmp_path / "q.docx")
    tgt = target_from_dict({"reference_style": "MLA"}, "user")
    _prof, findings = run_check(src, tgt)
    assert not _find(findings, "abstract_heading")
    assert not _find(findings, "keywords")


def test_user_channel_can_state_keywords(tmp_path):
    """手填通道能表达 Keywords 两项，且照判照修。"""
    src = _apa_doc(tmp_path / "u.docx", kw_indent=0.0, kw_italic=None)
    tgt = target_from_dict({"keywords_label_italic": True, "keywords_indent_in": 0.5},
                           "user")
    _prof, findings = run_check(src, tgt)
    got = _find(findings, "keywords")
    assert got and got[0].severity == "warn" and got[0].source == "user"
    out = tmp_path / "u_out.docx"
    fix_docx(src, str(out), tgt)
    _p2, again = run_check(str(out), tgt)
    assert _find(again, "keywords")[0].severity == "pass"


def test_no_keywords_line_is_info_not_warning(tmp_path):
    """没有 Keywords 行 → 只提示，绝不绑一个修不动的 FixAction。"""
    doc = Document()
    doc.add_paragraph("Abstract")
    doc.add_paragraph("Abstract body without any keywords line at all.")
    doc.add_paragraph("Introduction")
    doc.save(str(tmp_path / "nk.docx"))
    tgt = target_from_spec("APA-Student")
    _prof, findings = run_check(str(tmp_path / "nk.docx"), tgt)
    got = _find(findings, "keywords")
    assert got and got[0].severity == "info"
    assert getattr(got[0], "fix", None) is None
