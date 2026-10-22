"""正文粘连段拆分（v2.3.87）全链路测试：拆 → 读 → 查 → 改 → 预览。

背景（老板 2026-10-22 拍板「先拆粘连段」）：
v2.3.84 只给**参考文献区**做了软换行拆分，正文没做。真实学生论文
`student_paper_unformatted_20pages.docx` 实测：87 个段落里 **30 个含软换行**
（共 40 处），其中 30 个的第一行就是章节标题。后果两层：

  ① 粘块整体被当成一个正文段落 → 标题行跟着吃到 0.5″ 首行缩进
     （客户一眼看出「章节标题缩进了、正文反倒没缩进」）；
  ② 检查器压根看不见这些标题（它们不是独立段落），标题维度整块漏检。

本文件覆盖：
- 软换行拆成真实段落，**文字一行都不许变**；
- 🔴 `<w:br w:type="page|column">` 是分页/分栏符，绝不可当段边界（v2.3.84 踩过的坑）；
- 幂等（跑第二遍返回 0，文档不再变）；
- 拆开后「编号开头的短标题行」不吃正文首行缩进，正文段落照常 0.5″；
- 标题行仍然吃**行距**（APA §2.21 整篇双倍行距包含标题，不能顺带豁免）；
- 参考文献区的粘连段仍然照旧拆（v2.3.84 的行为不能被这次改动回退）；
- 预览与批注如实告诉客户「拆了几处、文字一个字没动」。
"""

from __future__ import annotations

import os
import sys

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.checker import run_check
from src.engine.docx_reader import (
    is_section_heading_line,
    read_docx,
    split_body_soft_breaks,
    split_paragraph_on_soft_breaks,
)
from src.engine.fixer import compute_changes, fix_docx
from src.engine.questionnaire import target_from_spec

BODY_1 = (
    "Digital technology adoption has reshaped how manufacturing firms coordinate "
    "with their upstream and downstream partners across regions."
)
BODY_2 = (
    "Existing literature on supply chain risk management can be broadly divided "
    "into two streams, neither of which fully explains the mechanism at work here."
)


# --------------------------------------------------------------------- helpers
def _lines(path: str) -> list:
    """把每个段落按软换行拆成「行」，用来做**字符级**比对（段数会变，行数不该变）。"""
    doc = Document(path)
    out: list = []
    for p in doc.paragraphs:
        buf: list = []
        for child in p._p:
            if child.tag == qn("w:pPr"):
                continue
            if child.tag == qn("w:r"):
                for el in child:
                    if el.tag == qn("w:rPr"):
                        continue
                    if el.tag in (qn("w:br"), qn("w:cr")):
                        out.append("".join(buf))
                        buf = []
                    elif el.tag == qn("w:t"):
                        buf.append(el.text or "")
        out.append("".join(buf))
    return out


def _make_glued_paper(path: str, *, with_refs: bool = True) -> str:
    """造一篇「章节标题 + 正文被 Shift+Enter 粘成一段」的短文（真实样本的缩小版）。"""
    doc = Document()
    doc.add_paragraph("Digital Transformation and Supply Chain Resilience")

    p = doc.add_paragraph()
    r = p.add_run("1. Introduction")
    r.add_break()                      # <w:br/> —— 学生最爱用的 Shift+Enter
    p.add_run(BODY_1)

    p = doc.add_paragraph()
    p.add_run(BODY_2)

    if with_refs:
        doc.add_paragraph("References")
        prefs = doc.add_paragraph()
        prefs.add_run("Ivanov, D. (2020). Predicting the impacts of epidemic outbreaks.")
        prefs.add_run().add_break()
        prefs.add_run("Sheffi, Y., & Rice, J. B. (2005). A supply chain view.")
    doc.save(path)
    return path


def _make_break_paper(path: str, kind: str) -> str:
    """造一篇段落内部**含分页符 / 分栏符**的文档（这两类绝不能被拆）。"""
    doc = Document()
    doc.add_paragraph("Digital Transformation and Supply Chain Resilience")
    doc.add_paragraph(BODY_1)
    p = doc.add_paragraph()
    p.add_run("A line before the layout break.")
    p.add_run().add_break(WD_BREAK.PAGE if kind == "page" else WD_BREAK.COLUMN)
    p.add_run("A line after the layout break.")
    doc.save(path)
    return path


def _find(paras, prefix: str):
    for p in paras:
        if (p.text or "").strip().startswith(prefix):
            return p
    return None


# ------------------------------------------------------------------ 拆分函数
def test_split_body_soft_breaks_splits_glued_paragraph(tmp_path):
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    doc = Document(src)
    before = len(doc.paragraphs)
    n = split_body_soft_breaks(doc)
    assert n == 2                                   # 正文 1 处 + 参考文献 1 处
    assert len(doc.paragraphs) == before + 2


def test_text_is_character_identical_after_split(tmp_path):
    """铁律：拆分只增加段落标记，字符一个都不许变。"""
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    out = tmp_path / "glued_split.docx"
    doc = Document(src)
    split_body_soft_breaks(doc)
    doc.save(str(out))
    assert _lines(src) == _lines(str(out))


def test_split_is_idempotent(tmp_path):
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    doc = Document(src)
    assert split_body_soft_breaks(doc) == 2
    assert split_body_soft_breaks(doc) == 0          # 第二遍无事可做
    assert split_body_soft_breaks(doc) == 0


def test_page_break_is_never_a_paragraph_boundary(tmp_path):
    """🔴 分页符是排版语义，拆掉就等于吃掉客户的分页。"""
    src = _make_break_paper(str(tmp_path / "page.docx"), "page")
    doc = Document(src)
    before = len(doc.paragraphs)
    assert split_body_soft_breaks(doc) == 0
    assert len(doc.paragraphs) == before
    # 分页符原样还在
    p = _find(doc.paragraphs, "A line before")
    types = [b.get(qn("w:type")) for b in p._p.iter(qn("w:br"))]
    assert types == ["page"]


def test_column_break_is_never_a_paragraph_boundary(tmp_path):
    src = _make_break_paper(str(tmp_path / "col.docx"), "column")
    doc = Document(src)
    before = len(doc.paragraphs)
    assert split_body_soft_breaks(doc) == 0
    assert len(doc.paragraphs) == before
    p = _find(doc.paragraphs, "A line before")
    types = [b.get(qn("w:type")) for b in p._p.iter(qn("w:br"))]
    assert types == ["column"]


def test_cr_is_split_but_page_break_is_not(tmp_path):
    """同一段里既有 <w:cr> 又有分页符：只拆 <w:cr>，分页符留在原地。"""
    doc = Document()
    doc.add_paragraph("Title")
    p = doc.add_paragraph()
    p.add_run("1. Introduction")
    br = p.add_run()
    br._r.append(br._r.makeelement(qn("w:cr"), {}))
    p.add_run(BODY_1)
    doc.save(str(tmp_path / "cr.docx"))

    doc2 = Document(str(tmp_path / "cr.docx"))
    assert split_body_soft_breaks(doc2) == 1
    assert _find(doc2.paragraphs, "1. Introduction").text.strip() == "1. Introduction"


def test_read_docx_reports_split_count(tmp_path):
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    prof = read_docx(src)
    assert prof.soft_break_splits == 2


def test_read_docx_sees_heading_hidden_in_glued_paragraph(tmp_path):
    """拆开之后检查器才看得见这个标题（拆之前它是正文的第一行，不是独立段落）。"""
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    doc = Document(src)
    rec = split_paragraph_on_soft_breaks(doc.paragraphs[1])
    assert rec == 1
    assert doc.paragraphs[1].text.strip() == "1. Introduction"


# --------------------------------------------------------- 首行缩进 / 行距归属
def test_heading_line_gets_no_body_first_line_indent(tmp_path):
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    out = tmp_path / "fixed.docx"
    fix_docx(src, str(out), target_from_spec("APA-Student"))

    doc = Document(str(out))
    head = _find(doc.paragraphs, "1. Introduction")
    assert head is not None
    fi = head.paragraph_format.first_line_indent
    assert fi is None or abs(fi.inches) < 0.01


def test_body_paragraph_still_gets_first_line_indent(tmp_path):
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    out = tmp_path / "fixed.docx"
    fix_docx(src, str(out), target_from_spec("APA-Student"))

    doc = Document(str(out))
    body = _find(doc.paragraphs, "Digital technology adoption")
    assert body is not None
    fi = body.paragraph_format.first_line_indent
    assert fi is not None and abs(fi.inches - 0.5) < 0.01


def test_heading_line_still_gets_line_spacing(tmp_path):
    """行距是整篇属性（APA §2.21），标题也在内 —— 豁免缩进不能顺带豁免行距。"""
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    out = tmp_path / "fixed.docx"
    fix_docx(src, str(out), target_from_spec("APA-Student"))

    doc = Document(str(out))
    head = _find(doc.paragraphs, "1. Introduction")
    ls = head.paragraph_format.line_spacing
    assert isinstance(ls, float) and abs(ls - 2.0) < 0.02


def test_text_unchanged_through_full_fix(tmp_path):
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    out = tmp_path / "fixed.docx"
    fix_docx(src, str(out), target_from_spec("APA-Student"))
    assert _lines(src) == _lines(str(out))


def test_reference_glued_paragraph_still_split(tmp_path):
    """v2.3.84 的参考文献拆分不能被这次改动回退。"""
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    out = tmp_path / "fixed.docx"
    fix_docx(src, str(out), target_from_spec("APA-Student"))

    doc = Document(str(out))
    hits = [p for p in doc.paragraphs if (p.text or "").strip().startswith("Ivanov")]
    assert len(hits) == 1
    assert "Sheffi" not in hits[0].text       # Sheffi 已是独立段落
    assert _find(doc.paragraphs, "Sheffi") is not None


# ------------------------------------------------------ is_section_heading_line
def test_section_heading_line_positives(tmp_path):
    doc = Document()
    for text in ("1. Introduction", "2.1 Sample Design", "3.2.1 Detail",
                 "I. INTRODUCTION", "A. Method", "一、引言", "Chapter 3"):
        doc.add_paragraph(text)
    doc.save(str(tmp_path / "h.docx"))
    doc2 = Document(str(tmp_path / "h.docx"))
    for p in doc2.paragraphs:
        assert is_section_heading_line(p) is True, p.text


def test_section_heading_line_negatives(tmp_path):
    doc = Document()
    for text in (
        "1. Introduction.",                      # 以句号结尾 → 是句子不是标题
        BODY_1,                                  # 正常正文
        "1. " + BODY_1,                          # 编号开头的正文段落（长 + 句号）
        "2. This section describes the model, and it ends with a comma,",  # 有句读
        "References",
        "",
    ):
        doc.add_paragraph(text)
    doc.save(str(tmp_path / "n.docx"))
    doc2 = Document(str(tmp_path / "n.docx"))
    for p in doc2.paragraphs:
        assert is_section_heading_line(p) is False, p.text


# ---------------------------------------------------------------------- 预览
def test_preview_mentions_split_in_zh(tmp_path):
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    lines = compute_changes(src, target_from_spec("APA-Student"), lang="zh")
    assert any("拆分粘连段落" in ln for ln in lines), lines


def test_preview_mentions_split_in_en(tmp_path):
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    lines = compute_changes(src, target_from_spec("APA-Student"), lang="en")
    assert any("Glued paragraphs split" in ln for ln in lines), lines


def test_preview_is_english_only(tmp_path):
    """英文界面绝不泄漏中文。"""
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    lines = compute_changes(src, target_from_spec("APA-Student"), lang="en")
    for ln in lines:
        assert not any("\u4e00" <= ch <= "\u9fff" for ch in ln), ln


def test_checker_does_not_regress_after_split(tmp_path):
    """拆完再查一遍：首行缩进维度必须是 pass（不能出现「修好了却恒报错」）。"""
    src = _make_glued_paper(str(tmp_path / "glued.docx"))
    out = tmp_path / "fixed.docx"
    tgt = target_from_spec("APA-Student")
    fix_docx(src, str(out), tgt)
    _, findings = run_check(str(out), tgt)
    fli = [f for f in findings if f.dimension == "first_line_indent"]
    assert fli and fli[0].severity == "pass", fli[0].actual if fli else "missing"
