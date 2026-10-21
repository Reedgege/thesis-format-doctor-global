"""块引用误判（撇号）与「整篇行距」两处修正的回归（v2.3.86）。

背景（老板 2026-10-21 追问「这几条都完工了吗」时逐条复量抓出来的两条漏项）：

① 清单第①「APA/MLA 强制 0.5″ 首行缩进」在真样本上只做到 54/61 —— 每章第一段丢缩进。
   真因：`has_quote = any(q in text for q in ('"', "“", "”", "'"))` **把英文撇号当引号**，
   任何 50 词以上、带一个 `firm's` / `company's` 的正文段都被判成「块引用」：
     - 报告里给客户报假问题「1 处长引文未按块引用缩进」；
     - 修正器把这些段从正文缩进里豁免掉 → 首行缩进丢失。
② 清单第①「2.0 行距」在 MLA 上只做到 60/61 —— 论文大标题那一行留在 1.0 倍。
   真因：行距只喂 `no_heading`，MLA 没有标题层级（heading_levels=0）只能靠 is_heading
   兜底，大标题被判成伪标题 → 被标题保护挡在整篇双倍行距之外。
   （APA 因为 heading_levels=5 走 level_map 过滤，恰好侥幸没被挡。）
"""

from __future__ import annotations

import os
import sys

from docx import Document

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.docx_reader import _is_quoted_paragraph, read_docx
from src.engine.questionnaire import target_from_spec, TargetProfile
from src.engine.checker import run_check
from src.engine.fixer import fix_docx


# ------------------------------------------------------------- 引号判定口径
def test_apostrophe_is_not_a_quote():
    """英文撇号（所有格/缩写）绝不能被当成引号。"""
    for txt in (
        "In recent years, more researchers have turned their attention to a firm's ability "
        "to integrate, build and reconfigure internal and external competencies.",
        "The company's supply chain resilience depends on its partners' flexibility.",
        "Don't treat an apostrophe as an opening quotation mark.",
    ):
        assert _is_quoted_paragraph(txt) is False, txt


def test_paragraph_wrapped_in_quotes_is_a_quote():
    for txt in (
        '"Supply chain resilience is the ability of a firm to react to disruptions."',
        "“Supply chain resilience is the ability of a firm to react to disruptions.”",
        "'Supply chain resilience is the ability to react.'",
    ):
        assert _is_quoted_paragraph(txt) is True, txt


def test_partial_quote_inside_sentence_is_not_a_block_quote():
    """句中夹一段引文 ≠ 整段被引用，不该当块引用。"""
    txt = ('The author argues that "resilience is dynamic" and then explains why '
           "the firm's partners matter over time.")
    assert _is_quoted_paragraph(txt) is False


def test_empty_and_short_text_safe():
    assert _is_quoted_paragraph("") is False
    assert _is_quoted_paragraph('"') is False
    assert _is_quoted_paragraph("   ") is False


def test_cjk_corner_quotes_recognised():
    assert _is_quoted_paragraph("「供应链韧性是企业应对中断的能力。」") is True


# ------------------------------------------------------------- 修：首行缩进
def _make_apostrophe_doc(path: str) -> str:
    """造一篇「带撇号的长正文段」的论文（就是真样本里被误判的形态）。"""
    doc = Document()
    doc.add_paragraph("A Study of Nothing In Particular")
    for _ in range(2):
        doc.add_paragraph("This is a plain body paragraph without any apostrophe at all, "
                          "long enough to be sampled as body text by the profiler.")
    doc.add_paragraph(
        "In recent years, more researchers have turned their attention to a firm's ability "
        "to integrate, build and reconfigure internal and external competencies to address "
        "rapidly changing environments, which is exactly the kind of long paragraph that "
        "used to be mistaken for a block quotation because of that single apostrophe.")
    doc.add_paragraph("Another plain body paragraph without apostrophes, again long enough "
                      "to be counted as body text.")
    doc.save(path)
    return path


def test_apostrophe_paragraph_still_gets_first_line_indent(tmp_path):
    """🔴 这条就是老板那份论文「每章第一段没缩进」的根因，钉死它。"""
    src = _make_apostrophe_doc(str(tmp_path / "ap.docx"))
    out = str(tmp_path / "ap_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    doc = Document(out)
    target_para = [p for p in doc.paragraphs if p.text.startswith("In recent years")][0]
    fli = target_para.paragraph_format.first_line_indent
    assert fli is not None and abs(fli.inches - 0.5) < 0.01


def test_profiler_does_not_flag_apostrophe_paragraph_as_quote(tmp_path):
    src = _make_apostrophe_doc(str(tmp_path / "ap2.docx"))
    prof = read_docx(src)
    flagged = [c for c in (prof.block_quote_candidates or []) if c.get("has_quote")]
    assert flagged == []


def test_checker_reports_no_block_quote_problem(tmp_path):
    src = _make_apostrophe_doc(str(tmp_path / "ap3.docx"))
    _, findings = run_check(src, target_from_spec("APA-Student"))
    assert not [f for f in findings if f.dimension == "block_quote"]


# ------------------------------------------------------------- 修：整篇行距
def _spacing_of(para):
    try:
        return para.paragraph_format.line_spacing
    except Exception:
        return None


def test_mla_title_line_is_double_spaced(tmp_path):
    """MLA 没有标题层级：大标题那行也必须吃到整篇 2.0 倍行距。"""
    src = _make_apostrophe_doc(str(tmp_path / "mla.docx"))
    out = str(tmp_path / "mla_fixed.docx")
    fix_docx(src, out, target_from_spec("MLA"))
    doc = Document(out)
    assert abs(float(_spacing_of(doc.paragraphs[0])) - 2.0) < 0.02


def test_apa_title_line_is_double_spaced(tmp_path):
    src = _make_apostrophe_doc(str(tmp_path / "apa.docx"))
    out = str(tmp_path / "apa_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    doc = Document(out)
    assert abs(float(_spacing_of(doc.paragraphs[0])) - 2.0) < 0.02


def test_line_spacing_does_not_touch_alignment_owner(tmp_path):
    """行距补了标题页段落，但对齐仍由 title_page 独占（一个属性只能一个 owner）。"""
    src = _make_apostrophe_doc(str(tmp_path / "own.docx"))
    out = str(tmp_path / "own_fixed.docx")
    fix_docx(src, out, target_from_spec("MLA"))
    doc = Document(out)
    # MLA 标题居中（title_page 管辖），正文左对齐（body_alignment 管辖）。
    # 这份合成稿的标题没被识别成标题页，所以这里只钉「没被正文对齐污染」。"""
    assert str(doc.paragraphs[0].alignment) not in ("LEFT (0)", "0", "JUSTIFY (3)", "3")
