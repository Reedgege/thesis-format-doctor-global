"""摘要定位的**严格限定文字回退**（v2.3.96）。

老板 2026-10-23 批准：「给摘要定位加一条严格限定的文字回退（只认前 15 段内、
整行匹配 Abstract、长度很短）」。

背景（真实论文 student_paper_unformatted_20pages.docx 实测）：标题页与 "Abstract"
全是 Normal 手排、没套任何标题样式 → 只认样式级标题的定位一律失败 →
「摘要另起一页 / 摘要顶格 / 正文另起一页」三件事永远只能提示、客户点了修正也不动。

本文件钉死五条契约：
  ① 有 Keywords 行做右边界时，没套样式的 "Abstract" 也能定位，且 body **含** Keywords 行
     （分页锚点取 max(body)+1，必须落在它之后）；
  ② 找不到 Keywords 行 → 一律不定位（右边界不可知 = 不把分页符插进客户正文中间）；
  ③ 只认整行、短行、前 15 段以内（长句 / 正文深处 / 子串都不算）；
  ④ 摘要**正文**顶格、Keywords 行保持 0.5″ —— 两者的样式口径相反，绝不能被一起刷；
  ⑤ 主体流程：APA 判 warn → 点了能修 → 复检 pass，且文字一字不动。

另有一条反向契约：摘要能定位之后，学校/手填/AI 通道**没提**「正文另起一页」时
依然一句都不许说（v2.3.93 的「谁说话就听谁的」不许被这版破坏）。
"""

from __future__ import annotations

import os
import sys

import pytest

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine import checker
from src.engine.docx_reader import (
    read_docx, find_abstract_section, is_keywords_line,
    abstract_body_style_indices,
)
from src.engine.fixer import fix_docx
from src.engine.questionnaire import target_from_spec, target_from_dict


# ---------------------------------------------------------------------------
# 夹具：真实学生论文的首页结构（全 Normal 手排，没有任何标题样式）
# ---------------------------------------------------------------------------

BODY = "This paper studies how digital transformation affects resilience."


def _make_paper(path, *, keywords="Keywords: digital transformation, supply chain",
                abstract_text=BODY, indent=None, pad_before=2,
                heading="Abstract"):
    """构造一份「手排摘要」的论文：题名页 → Abstract → 摘要正文 → Keywords → 正文。

    indent=None 表示不写首行缩进（Word 默认 0）；pad_before 控制标题页占几段。
    """
    doc = Document()
    doc.add_paragraph("Digital Transformation and Supply Chain Resilience")
    for i in range(max(0, pad_before - 1)):
        doc.add_paragraph("Title page line %d" % (i + 1))
    doc.add_paragraph(heading)
    a = doc.add_paragraph(abstract_text)
    if indent is not None:
        a.paragraph_format.first_line_indent = Inches(indent)
    if keywords is not None:
        k = doc.add_paragraph(keywords)
        k.paragraph_format.first_line_indent = Inches(0.5)
    doc.add_paragraph("1. Introduction")
    for i in range(3):
        p = doc.add_paragraph("Body paragraph %d about the methods." % (i + 1))
        p.paragraph_format.first_line_indent = Inches(0.5)
    doc.save(str(path))
    return str(path)


def _paras(path):
    return list(Document(path).paragraphs)


def _sev(findings, dim):
    got = [f for f in findings if f.dimension == dim]
    return got[0].severity if got else None


def _finding(findings, dim):
    got = [f for f in findings if f.dimension == dim]
    return got[0] if got else None


def _all_text(path):
    return "\n".join((p.text or "") for p in Document(path).paragraphs)


# ---------------------------------------------------------------------------
# ① 能定位：没套样式的 "Abstract" + Keywords 行
# ---------------------------------------------------------------------------

def test_unstyled_abstract_is_located_when_keywords_line_exists(tmp_path):
    p = _make_paper(tmp_path / "a.docx", pad_before=2)
    head, body = find_abstract_section(_paras(p))
    assert head == 2, "第 2 段是手排的 Abstract，应被文字回退认出来"
    assert body == [3, 4], "摘要正文 + Keywords 行都要算进摘要块"
    assert is_keywords_line(_paras(p)[4])


def test_located_abstract_makes_has_abstract_and_metrics_readable(tmp_path):
    """从「读不到」变成「读得出来」：摘要块一带三（缩进 / 另起一页 / 正文另起一页）。"""
    p = _make_paper(tmp_path / "a.docx", indent=0.5)
    prof = read_docx(p)
    assert prof.has_abstract is True
    assert prof.abstract_first_line_indent_in is not None
    assert prof.text_break_before is False, "正文紧跟摘要、没有分页符"


def test_style_indices_drop_the_keywords_line(tmp_path):
    p = _make_paper(tmp_path / "a.docx", indent=0.5)
    paras = _paras(p)
    head, body = find_abstract_section(paras)
    assert abstract_body_style_indices(paras, body) == [3], "样式口径只留摘要正文"
    assert abstract_body_style_indices(paras, [4]) == [4], "全被剔光时退回原列表，不许丢段"


# ---------------------------------------------------------------------------
# ② 右边界不可知 → 一律不定位
# ---------------------------------------------------------------------------

def test_no_keywords_line_keeps_the_abstract_unlocated(tmp_path):
    p = _make_paper(tmp_path / "a.docx", keywords=None)
    assert find_abstract_section(_paras(p)) == (None, [])


def test_keywords_line_without_abstract_text_is_rejected(tmp_path):
    """标题后直接跟 Keywords、中间一句摘要正文都没有 → 不像摘要段，放弃。"""
    p = _make_paper(tmp_path / "a.docx", abstract_text="   ")
    assert find_abstract_section(_paras(p)) == (None, [])


def test_no_abstract_at_all_stays_unlocated(tmp_path):
    doc = Document()
    doc.add_paragraph("A paper without any abstract")
    doc.add_paragraph("1. Introduction")
    for i in range(3):
        doc.add_paragraph("Body %d." % i)
    p = str(tmp_path / "n.docx")
    doc.save(p)
    assert find_abstract_section(_paras(p)) == (None, [])


# ---------------------------------------------------------------------------
# ③ 严格边界：整行 / 短行 / 前 15 段
# ---------------------------------------------------------------------------

def test_long_sentence_starting_with_abstract_is_not_a_heading(tmp_path):
    """正文里一句 "Abstract writing is ..." 绝不能被当成摘要标题。"""
    body = "Abstract writing is a skill that many students struggle with a lot."
    p = _make_paper(tmp_path / "a.docx", abstract_text=body, heading=body)
    assert len(body) > 40
    assert find_abstract_section(_paras(p)) == (None, [])


def test_abstract_beyond_the_first_fifteen_paragraphs_is_ignored(tmp_path):
    p = _make_paper(tmp_path / "a.docx", pad_before=20)
    assert find_abstract_section(_paras(p)) == (None, [])


def test_keywords_variants_are_recognised():
    """关键词行只认行首标签 + 冒号：中英文、全角半角、Key words 都算。"""
    from docx import Document as _D
    doc = _D()
    for text, want in (
        ("Keywords: alpha, beta", True),
        ("Key words: alpha", True),
        ("key%swords:  alpha" % " ", True),
        ("关键词：数字转型", True),
        ("KEYWORDS: X", True),
        ("The keywords: this sentence merely mentions keywords", False),
        ("Keywords without a colon", False),
        ("", False),
    ):
        para = doc.add_paragraph(text)
        assert is_keywords_line(para) is want, text


# ---------------------------------------------------------------------------
# ④ 摘要正文顶格，Keywords 行不许被刷平
# ---------------------------------------------------------------------------

def _school_target(**kw):
    """学校/手填通道：显式要求「摘要顶格」（spec_key=None = 非规范基线）。"""
    data = {"abstract_first_line_indent_in": 0.0}
    data.update(kw)
    t = target_from_dict(data, "school")
    assert t.abstract_first_line_indent_in == 0.0
    return t


def test_fixer_flattens_abstract_body_but_keeps_keywords_indent(tmp_path):
    p = _make_paper(tmp_path / "a.docx", indent=0.5)
    t = _school_target()
    _prof, findings = checker.run_check(p, t)
    assert _sev(findings, "abstract_indent") == "warn"

    out = str(tmp_path / "out.docx")
    fix_docx(p, out, t)
    paras = _paras(out)
    assert paras[3].paragraph_format.first_line_indent in (None,) or \
        abs(paras[3].paragraph_format.first_line_indent.inches) < 1e-3, "摘要正文应顶格"
    kw = paras[4].paragraph_format.first_line_indent
    assert kw is not None and abs(kw.inches - 0.5) < 1e-3, \
        "Keywords 行必须保持 0.5″（APA：摘要正文顶格、Keywords 行缩进），不许被一起刷平"


def test_apa_baseline_does_not_flatten_the_abstract_behind_the_users_back(tmp_path):
    """APA 基线只**报**摘要缩进、不擅自改（摘要页是客户的版式选择）。

    真实论文的摘要本来就是顶格 → 报 pass；关键是正文首行缩进流程
    不许把它刷成 0.5″（v2.3.96 之前正是这样：豁免集合走的是另一份只认样式标题的实现）。
    """
    p = _make_paper(tmp_path / "a.docx", indent=None)
    t = target_from_spec("APA-Student")
    _prof, findings = checker.run_check(p, t)
    assert _sev(findings, "abstract_indent") == "pass"

    out = str(tmp_path / "out.docx")
    fix_docx(p, out, t)
    _prof2, f2 = checker.run_check(out, t)
    assert _sev(f2, "abstract_indent") == "pass", "修完不许把摘要刷成缩进"


# ---------------------------------------------------------------------------
# ⑤ 主体流程：APA 判 warn → 能修 → 复检 pass → 文字不动
# ---------------------------------------------------------------------------

def test_apa_reports_and_fixes_body_new_page_on_unstyled_abstract(tmp_path):
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Student")

    _prof, findings = checker.run_check(p, t)
    f = _finding(findings, "text_new_page")
    assert f is not None and f.severity == "warn", "摘要定位得到 → 必须判得出来"
    assert f.fix is not None, "报了就必须能修（不许留报了却修不了的缺口）"

    before = _all_text(p)
    out = str(tmp_path / "out.docx")
    fix_docx(p, out, t)
    assert _all_text(out) == before, "纯版式改动，文字一个字都不许动"

    _prof2, f2 = checker.run_check(out, t)
    assert _sev(f2, "text_new_page") == "pass", "修完复检必须收敛"

    paras = _paras(out)
    idx = [i for i, x in enumerate(paras) if (x.text or "").strip().startswith("1. Introduction")]
    assert idx, "正文首段还在"
    assert bool(paras[idx[0]].paragraph_format.page_break_before) is True


def test_apa_reports_abstract_new_page_on_unstyled_abstract(tmp_path):
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_spec("APA-Student")
    _prof, findings = checker.run_check(p, t)
    assert _sev(findings, "abstract_new_page") == "warn"

    out = str(tmp_path / "out.docx")
    fix_docx(p, out, t)
    _prof2, f2 = checker.run_check(out, t)
    assert _sev(f2, "abstract_new_page") == "pass"


# ---------------------------------------------------------------------------
# 反向契约：通道没说话就闭嘴（v2.3.93）
# ---------------------------------------------------------------------------

def test_school_channel_stays_silent_about_text_new_page(tmp_path):
    """摘要这次能定位了，但学校通道没提「正文另起一页」→ 一句都不许说。"""
    p = _make_paper(tmp_path / "a.docx")
    t = target_from_dict({"first_line_indent_in": 0.5}, "school")
    _prof, findings = checker.run_check(p, t)
    assert _sev(findings, "text_new_page") is None
    assert _sev(findings, "abstract_new_page") is None


def test_no_abstract_means_no_text_new_page_finding(tmp_path):
    """文档里压根没有摘要 → 不许判「正文没另起一页」（没有可对齐的锚点）。"""
    doc = Document()
    doc.add_paragraph("Title of the paper")
    doc.add_paragraph("1. Introduction")
    for i in range(3):
        doc.add_paragraph("Body %d." % i)
    p = str(tmp_path / "n.docx")
    doc.save(p)
    t = target_from_spec("APA-Student")
    _prof, findings = checker.run_check(p, t)
    assert _sev(findings, "text_new_page") is None
