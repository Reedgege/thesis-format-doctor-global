# -*- coding: utf-8 -*-
"""表格与正文之间的空隙（v2.3.98）。

老板 2026-10-23 截图问：「规范是怎么写的，看着这个一级标题跟表格好近呀」。
取证结果：原文档那张表下方的 "Chapter 23" 段前 14pt，被我们「正文段前段后=0」
一起清成了 0 —— 因为此前**没有任何维度**管「表格与文字之间的空隙」。

APA 7（Publication Manual 7th, Tables and Figures）原文：
    "If embedded, place the table or figure at either the top or the bottom of the
     page with an extra double-spaced line between the table or the figure and any text."
12pt 正文的一条双倍行距 = 24pt。

五条护栏：
  ① 判得出、修得动、修完复检收敛（不留红杠）；
  ② **不插空段落**、不碰一个字（只写紧邻段落的 w:spacing）；
  ③ 表格紧邻段落不被「正文段前段后=0」清掉（两个维度不互清，防 v2.3.90 的自伤）；
  ④ 表格旁边已经是空段落 / 旁边是另一张表 → 不判不改；
  ⑤ MLA / IEEE 不规定这项 → 一个字都不说不改。
"""

from __future__ import annotations

import os
import sys

import pytest

from docx import Document
from docx.shared import Pt

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.checker import run_check
from src.engine.docx_reader import is_caption_paragraph, read_docx
from src.engine.fixer import fix_docx
from src.engine.questionnaire import target_from_dict, target_from_spec


# ---------------------------------------------------------------------------
# 夹具
# ---------------------------------------------------------------------------

_BODY = ("The sample covers 1,204 firms across four provinces, and every model is "
         "estimated with firm-level clustered standard errors.")


def _paper(path, before_body_pt=0.0, after_body_pt=0.0, before_table_pt=0.0,
           after_table_pt=0.0, blank_before=False, blank_after=False,
           two_tables=False):
    """正文段 → （可选空段）→ 表格 → （可选空段）→ 正文段。

    before_table_pt / after_table_pt 是**表格紧邻那两段**的段后 / 段前间距，
    也就是我们要管的那个空隙。
    """
    doc = Document()
    doc.add_paragraph("Digital Transformation and Firm Performance")
    lead = doc.add_paragraph(_BODY)
    lead.paragraph_format.space_before = Pt(before_body_pt)
    lead.paragraph_format.space_after = Pt(before_table_pt)

    if blank_before:
        doc.add_paragraph("")

    doc.add_table(rows=2, cols=2)
    if two_tables:
        doc.add_table(rows=2, cols=2)

    if blank_after:
        doc.add_paragraph("")

    tail = doc.add_paragraph(_BODY)
    tail.paragraph_format.space_before = Pt(after_table_pt)
    tail.paragraph_format.space_after = Pt(after_body_pt)
    doc.save(str(path))
    return str(path)


def _find(findings, dim):
    return [f for f in findings if f.dimension == dim]


def _problems(findings, dim):
    """只取要改的（warn / fail）；pass 与 info 是「已合规」与「无从判定」，不算问题。"""
    return [f for f in _find(findings, dim) if f.severity in ("warn", "fail")]


def _gap_side(path, which):
    """读回第一张表某一侧紧邻段落的 w:spacing 值（空段返回 None）。"""
    prof = read_docx(path)
    samples = prof.table_gap_samples or []
    assert samples, "文档里有表格，采样不该为空"
    key = "before_pt" if which == "before" else "after_pt"
    blank = "before_blank" if which == "before" else "after_blank"
    if samples[0].get(blank):
        return None
    return samples[0].get(key)


# ---------------------------------------------------------------------------
# 一、判得出 / 修得动 / 收敛
# ---------------------------------------------------------------------------

def test_apa_flags_tight_table_gap(tmp_path):
    """APA 下表格紧贴文字（0pt）→ 必须报出来，且挂得上自动修正。"""
    src = _paper(tmp_path / "tight.docx")
    _prof, findings = run_check(src, target_from_spec("APA-Student"))
    gaps = _find(findings, "table_gap")
    assert gaps, "APA 要求表格与文字间空一条双倍行距，0pt 必须报"
    assert any(f.fix is not None for f in gaps), "这一维度必须修得动（不是只提示）"


def test_apa_fix_opens_the_gap_and_recheck_is_clean(tmp_path):
    """修完 24pt，复检必须收敛（否则就是「修完又报红」的自伤）。"""
    src = _paper(tmp_path / "fix.docx")
    tgt = target_from_spec("APA-Student")
    out = str(tmp_path / "fixed.docx")
    res = fix_docx(src, out, tgt)
    assert "table_gap_pt" in res.get("applied", {}), res.get("applied")

    assert _gap_side(out, "before") == pytest.approx(24.0, abs=0.5)
    assert _gap_side(out, "after") == pytest.approx(24.0, abs=0.5)

    _prof2, findings2 = run_check(out, tgt)
    assert _problems(findings2, "table_gap") == [], \
        [(f.severity, f.actual) for f in _find(findings2, "table_gap")]


def test_fix_does_not_insert_blank_paragraphs(tmp_path):
    """🔴 只写 w:spacing，绝不插空段落（插段会把段落顺序/下标全体错位）。"""
    src = _paper(tmp_path / "noblank.docx")
    before = len(Document(src).paragraphs)
    out = str(tmp_path / "noblank_out.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    after = len(Document(out).paragraphs)
    assert after == before, "修正器插了空段落：%d → %d" % (before, after)


def test_fix_does_not_touch_a_single_character(tmp_path):
    """空格子只动间距，正文文字逐字不变。"""
    src = _paper(tmp_path / "chars.docx")
    out = str(tmp_path / "chars_out.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    src_text = [p.text for p in Document(src).paragraphs]
    out_text = [p.text for p in Document(out).paragraphs]
    assert src_text == out_text


def test_gap_survives_a_heading_right_below_the_table(tmp_path):
    """🔴 表格下方紧邻的是标题时，24pt 不能被「标题不加额外段前」清掉。

    真实论文（student_paper_unformatted_20pages.docx）就是这样：每张表下面紧跟
    「Chapter 23: Discussion and Strategic Insights」。`_format_heading_one` 按
    APA「标题不加额外段前」写 before=0，v2.3.98 第一版就是把 24pt 当场清成 0，
    复检照样报红、第二遍再修还是 0 —— 客户每点一次「修正」原地打转。
    """
    doc = Document()
    doc.add_paragraph("Digital Transformation and Firm Performance")
    lead = doc.add_paragraph(_BODY)
    doc.add_table(rows=2, cols=2)
    h = doc.add_paragraph("2. Discussion and Strategic Insights")
    for r in h.runs:
        r.bold = True
        r.font.size = Pt(12)
    doc.add_paragraph(_BODY)
    src = str(tmp_path / "hdg.docx")
    doc.save(src)

    tgt = target_from_spec("APA-Student")
    out = str(tmp_path / "hdg_out.docx")
    res = fix_docx(src, out, tgt)
    assert "table_gap_pt" in res.get("applied", {}), res.get("applied")

    assert _gap_side(out, "after") == pytest.approx(24.0, abs=0.5), \
        "表格下方的标题段段前被清成 0 —— table_gap 必须排在标题格式化之后"

    # 且第二遍必须零动作（否则就是「修不完」的死循环）
    out2 = str(tmp_path / "hdg_out2.docx")
    res2 = fix_docx(out, out2, tgt)
    assert "table_gap_pt" not in res2.get("applied", {}), res2.get("applied")


# ---------------------------------------------------------------------------
# 二、两维度不互清（v2.3.90 踩过的「一段两个主人」）
# ---------------------------------------------------------------------------

def test_body_spacing_rule_does_not_wipe_the_table_gap(tmp_path):
    """「正文段前段后=0」不得把刚给表格留的 24pt 清掉。"""
    src = _paper(tmp_path / "both.docx")
    tgt = target_from_spec("APA-Student")
    out = str(tmp_path / "both_out.docx")
    fix_docx(src, out, tgt)
    assert _gap_side(out, "before") == pytest.approx(24.0, abs=0.5)
    assert _gap_side(out, "after") == pytest.approx(24.0, abs=0.5)

    _p, findings = run_check(out, tgt)
    assert _problems(findings, "space_before") == []
    assert _problems(findings, "space_after") == []


def test_table_adjacent_paragraph_does_not_raise_the_body_mode(tmp_path):
    """表格旁边留了 24pt，不得让「正文段距众数」被抬成 24。"""
    doc = Document()
    for _ in range(6):
        doc.add_paragraph(_BODY)                       # 6 段正常正文（段距 0）
    lead = doc.add_paragraph(_BODY)
    lead.paragraph_format.space_after = Pt(24)         # 表格上方那段
    doc.add_table(rows=2, cols=2)
    tail = doc.add_paragraph(_BODY)
    tail.paragraph_format.space_before = Pt(24)        # 表格下方那段
    doc.save(str(tmp_path / "mode.docx"))

    prof = read_docx(tmp_path / "mode.docx")
    assert prof.space_before_pt == pytest.approx(0.0, abs=0.01)
    assert prof.space_after_pt == pytest.approx(0.0, abs=0.01)


# ---------------------------------------------------------------------------
# 三、不该开口时闭嘴
# ---------------------------------------------------------------------------

def test_existing_blank_line_counts_as_a_gap(tmp_path):
    """表格旁边已经是空段落（视觉上已经空了一行）→ 不判不改。"""
    src = _paper(tmp_path / "blank.docx", blank_before=True, blank_after=True)
    tgt = target_from_spec("APA-Student")
    _p, findings = run_check(src, tgt)
    assert _problems(findings, "table_gap") == [], \
        [(f.severity, f.actual) for f in _find(findings, "table_gap")]
    # 真空行已存在 → 只给 info，不能说「间距够」（那是假绿）
    assert all(f.severity == "info" for f in _find(findings, "table_gap"))
    out = str(tmp_path / "blank_out.docx")
    res = fix_docx(src, out, tgt)
    assert "table_gap_pt" not in res.get("applied", {})


def test_table_next_to_another_table_is_skipped(tmp_path):
    """相邻两表之间没有文字 → 不去给表格自己编间距。"""
    src = _paper(tmp_path / "two.docx", blank_before=True, two_tables=True)
    prof = read_docx(src)
    samples = prof.table_gap_samples or []
    assert len(samples) == 2
    inner = [s for s in samples if s.get("before_pt") is not None
             or s.get("after_pt") is not None]
    # 内测那张表（第 2 张）的 before 邻居是另一张表 → 必须是 None（不判）
    assert samples[1].get("before_pt") is None


@pytest.mark.parametrize("spec_key", ["MLA", "IEEE"])
def test_style_silence_means_no_finding(tmp_path, spec_key):
    """MLA / IEEE 不规定表格与文字的间距 → 一个结论都不给、一个字节不改。"""
    src = _paper(tmp_path / ("sil_%s.docx" % spec_key))
    tgt = target_from_spec(spec_key)
    _p, findings = run_check(src, tgt)
    assert _find(findings, "table_gap") == []

    out = str(tmp_path / ("sil_out_%s.docx" % spec_key))
    res = fix_docx(src, out, tgt)
    assert "table_gap_pt" not in res.get("applied", {})
    assert _gap_side(out, "before") == pytest.approx(0.0, abs=0.01)


def test_user_channel_can_state_the_gap(tmp_path):
    """手填/AI 通道明说「表格与正文间距 30 磅」→ 按客户说的判和修。"""
    src = _paper(tmp_path / "user.docx")
    tgt = target_from_dict({"table_gap_pt": 30.0}, source="questionnaire")
    _p, findings = run_check(src, tgt)
    gaps = _find(findings, "table_gap")
    assert gaps and any(f.fix is not None for f in gaps)

    out = str(tmp_path / "user_out.docx")
    fix_docx(src, out, tgt)
    assert _gap_side(out, "before") == pytest.approx(30.0, abs=0.5)
    assert _gap_side(out, "after") == pytest.approx(30.0, abs=0.5)


def test_channel_silence_means_no_finding(tmp_path):
    """三通道都没提、也没选内置规范 → 零结论（不是给个 info 凑数）。"""
    src = _paper(tmp_path / "none.docx")
    tgt = target_from_dict({"font_size_pt": 12.0}, source="manual")
    assert getattr(tgt, "table_gap_pt", None) is None
    _p, findings = run_check(src, tgt)
    assert _find(findings, "table_gap") == []


def test_no_tables_is_info_only(tmp_path):
    """文档里根本没有表格 → 只给 info，绝不挂修正动作。"""
    doc = Document()
    doc.add_paragraph("Digital Transformation and Firm Performance")
    doc.add_paragraph(_BODY)
    doc.save(str(tmp_path / "notbl.docx"))
    _p, findings = run_check(str(tmp_path / "notbl.docx"),
                             target_from_spec("APA-Student"))
    gaps = _find(findings, "table_gap")
    assert all(f.fix is None for f in gaps)


# ---------------------------------------------------------------------------
# 四、顺带修的假绿：把「提到表格的正文句」当表题
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("text, expected", [
    ("Table 1 below shows the detailed parameter estimates for every model.",
     False),                      # 正文句里顺口提一句 → 不是表题
    ("Table 2 above reports the robustness checks.", False),
    ("Table 1", True),            # 真正的表号行
    ("Table 1", True),
    ("Table 3.  Descriptive statistics by province", True),
    ("Figure 2", True),
    ("表 1  各变量描述性统计", True),
])
def test_caption_detection_does_not_swallow_sentences(text, expected):
    """「Table 1 below shows…」这种正文句曾被当成表题，导致 caption_position_table
    在**根本没有表题**的文档上判 pass（假绿）。"""
    doc = Document()
    p = doc.add_paragraph(text)
    assert is_caption_paragraph(p) is expected
