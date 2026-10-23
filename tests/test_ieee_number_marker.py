"""IEEE 参考文献「1.」→「[1]」列表符号规范化（v2.3.86）。

老板 2026-10-21 拍板：把 `1. / 1) / (1)` 这类数字列表符号换成 IEEE 的 `[n]`，
**只换符号，条目正文一个字不动**（与「文字 100% 锁死」的冲突按此裁定）。

为什么必须做：真样本条目是 `1. Ivanov, D. (2020)…`，
检查器如实报 `No numbered [n] form detected`（IEEE 要 [1]），而旧修正器把 `1.`
也算「已有编号」→ 一个字没改 → 「报了却修不了」的经典脱节。
"""

from __future__ import annotations

import os
import sys

from docx import Document

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.checker import run_check
from src.engine.docx_reader import reference_paragraph_indices
from src.engine.fixer import fix_docx, compute_changes
from src.engine.questionnaire import target_from_spec, TargetProfile

ENTRIES = [
    "1. Ivanov, D. (2020). Predicting the impacts of epidemic outbreaks on global supply chains.",
    "2. Sheffi, Y., & Rice, J. B. (2005). A supply chain view of the resilient enterprise.",
    "3. Teece, D. J., Pisano, G., & Shuen, A. (1997). Dynamic capabilities and strategic management.",
]


def _make(path, entries=ENTRIES):
    doc = Document()
    doc.add_paragraph("A Study of IEEE References")
    doc.add_paragraph("Body paragraph long enough to be sampled as body text by the profiler.")
    doc.add_paragraph("References")
    for e in entries:
        doc.add_paragraph(e)
    doc.save(path)
    return path


def _ref_texts(path):
    doc = Document(path)
    return [doc.paragraphs[i].text for i in reference_paragraph_indices(doc)]


def _strip_marker(text: str) -> str:
    """去掉开头的列表符号（[n] / n. / n) / (n)），用于比对「正文是否一字未动」。"""
    import re
    return re.sub(r"^\s*(?:\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\s*[.)])\s*", "", text or "")


def test_plain_numbers_become_brackets(tmp_path):
    src = _make(str(tmp_path / "a.docx"))
    out = str(tmp_path / "a_fixed.docx")
    fix_docx(src, out, target_from_spec("IEEE"))
    texts = _ref_texts(out)
    assert texts[0].startswith("[1] ")
    assert texts[1].startswith("[2] ")
    assert texts[2].startswith("[3] ")


def test_entry_body_text_untouched(tmp_path):
    src = _make(str(tmp_path / "b.docx"))
    out = str(tmp_path / "b_fixed.docx")
    fix_docx(src, out, target_from_spec("IEEE"))
    before = [_strip_marker(t) for t in _ref_texts(src)]
    after = [_strip_marker(t) for t in _ref_texts(out)]
    assert before == after


def test_other_marker_styles_normalised(tmp_path):
    entries = [
        "1) Smith, J. A. (2019). First entry.",
        "(2) Lee, K. (2018). Second entry.",
        "3. Chen, L. (2017). Third entry.",
    ]
    src = _make(str(tmp_path / "c.docx"), entries=entries)
    out = str(tmp_path / "c_fixed.docx")
    fix_docx(src, out, target_from_spec("IEEE"))
    texts = _ref_texts(out)
    assert texts[0].startswith("[1] Smith")
    assert texts[1].startswith("[2] Lee")
    assert texts[2].startswith("[3] Chen")


def test_already_bracketed_not_double_prefixed(tmp_path):
    entries = [
        "[1] Ivanov, D. (2020). First entry.",
        "[2] Sheffi, Y. (2005). Second entry.",
    ]
    src = _make(str(tmp_path / "d.docx"), entries=entries)
    out = str(tmp_path / "d_fixed.docx")
    fix_docx(src, out, target_from_spec("IEEE"))
    texts = _ref_texts(out)
    assert texts[0].startswith("[1] Ivanov")
    assert "[1] [1]" not in texts[0]


def test_apa_does_not_normalise_markers(tmp_path):
    """APA 是著者-年制（不编号）：绝不该动任何前缀。"""
    entries = ["Ivanov, D. (2020). First entry.", "Sheffi, Y. (2005). Second entry."]
    src = _make(str(tmp_path / "e.docx"), entries=entries)
    out = str(tmp_path / "e_fixed.docx")
    fix_docx(src, out, target_from_spec("APA-Student"))
    assert _ref_texts(out) == entries


def test_numbering_forbidden_strips_plain_marker(tmp_path):
    """🔴 v2.3.95 契约升级：目标**明确不编号**时，`1.` 必须被去掉（旧版原样保留）。

    旧行为（v2.3.86 ~ v2.3.94）：只有「该补编号 → 补 [n]」一条路，反方向一个字节都不动。
    结果老板 2026-10-23 拿 v2.3.93 的四份产物核对时，APA 学生版 / APA 专业版 / MLA
    三份的条目仍是 `1. Ivanov, D. (2020)…` —— APA 7 §9.44 与 MLA 的 Works Cited
    都不许给条目加序号，客户就这么交上去了，而报告全绿。
    性质与 `1.` → `[1]` 完全同类：只动段首那个列表记号，条目正文一个字不碰。
    """
    src = _make(str(tmp_path / "f.docx"))
    out = str(tmp_path / "f_fixed.docx")
    t = TargetProfile(reference_style="APA", reference_numbered=False)
    fix_docx(src, out, t)
    texts = _ref_texts(out)
    assert texts[0].startswith("Ivanov, D. (2020)."), texts
    assert texts[1].startswith("Sheffi, Y."), texts
    assert texts[2].startswith("Teece, D. J."), texts


def test_numbering_forbidden_also_strips_brackets(tmp_path):
    """已经是 `[1]` 的条目也要去掉 —— 「不编号」的规范连方括号编号也不许留。"""
    entries = ["[1] Ivanov, D. (2020). First entry.", "[2] Sheffi, Y. (2005). Second entry."]
    src = _make(str(tmp_path / "f2.docx"), entries=entries)
    out = str(tmp_path / "f2_fixed.docx")
    t = TargetProfile(reference_style="APA", reference_numbered=False)
    fix_docx(src, out, t)
    assert _ref_texts(out) == ["Ivanov, D. (2020). First entry.",
                               "Sheffi, Y. (2005). Second entry."]


def test_numbering_forbidden_leaves_entry_body_untouched(tmp_path):
    """去编号只删段首记号；条目正文（含标点）必须逐字不变。"""
    src = _make(str(tmp_path / "f3.docx"))
    out = str(tmp_path / "f3_fixed.docx")
    fix_docx(src, out, TargetProfile(reference_style="MLA", reference_numbered=False))
    before = [_strip_marker(t) for t in _ref_texts(src)]
    assert _ref_texts(out) == before


def test_numbering_unspecified_leaves_entries_alone(tmp_path):
    """reference_numbered 未声明（None）= 不规定 → 一个字节都不许动。"""
    src = _make(str(tmp_path / "f4.docx"))
    out = str(tmp_path / "f4_fixed.docx")
    fix_docx(src, out, TargetProfile(reference_style=None, reference_numbered=None))
    assert _ref_texts(out) == _ref_texts(src)


def test_checker_reports_numbered_entries_when_style_forbids(tmp_path):
    """条目不编号的规范（APA）遇到 `1.` 条目 → 报告要说话，且可修、修完收敛。"""
    src = _make(str(tmp_path / "g2.docx"))
    t = target_from_spec("APA-Student")
    _, f1 = run_check(src, t)
    ref = [f for f in f1 if f.dimension == "reference" and f.severity in ("warn", "fail")]
    assert ref, "APA 下条目带 1./2./3. 必须报出来"
    assert any(getattr(f, "fix", None) and f.fix.dimension == "reference_numbered"
               for f in ref)
    out = str(tmp_path / "g2_fixed.docx")
    fix_docx(src, out, t)
    _, f2 = run_check(out, t)
    assert not [f for f in f2 if f.dimension == "reference"
                and f.severity in ("warn", "fail")]


def test_preview_mentions_marker_removal(tmp_path):
    """预览必须先说「序号会被去掉」，客户才知道段首的 `1.` 为什么不见了。"""
    src = _make(str(tmp_path / "h2.docx"))
    lines = compute_changes(src, target_from_spec("APA-Student"), lang="en")
    assert any("remove the leading list number" in ln for ln in lines), lines
    assert not any("add [n] sequential code" in ln for ln in lines), lines



def test_checker_warn_then_pass_after_fix(tmp_path):
    src = _make(str(tmp_path / "g.docx"))
    out = str(tmp_path / "g_fixed.docx")
    t = target_from_spec("IEEE")
    _, f1 = run_check(src, t)
    ref1 = [f for f in f1 if f.dimension == "reference"]
    assert ref1 and ref1[0].severity == "warn"
    fix_docx(src, out, t)
    _, f2 = run_check(out, t)
    ref2 = [f for f in f2 if f.dimension == "reference"]
    assert not ref2 or all(f.severity != "warn" for f in ref2), ref2


def test_preview_mentions_marker_change(tmp_path):
    src = _make(str(tmp_path / "h.docx"))
    lines = compute_changes(src, target_from_spec("IEEE"), lang="en")
    assert any("list marker" in ln for ln in lines), lines


def test_preview_line_is_english(tmp_path):
    src = _make(str(tmp_path / "i.docx"))
    for ln in compute_changes(src, target_from_spec("IEEE"), lang="en"):
        assert all(ord(ch) < 0x2E80 for ch in ln), ln


def test_fix_is_idempotent(tmp_path):
    src = _make(str(tmp_path / "j.docx"))
    once = str(tmp_path / "j1.docx")
    twice = str(tmp_path / "j2.docx")
    t = target_from_spec("IEEE")
    fix_docx(src, once, t)
    fix_docx(once, twice, t)
    assert _ref_texts(once) == _ref_texts(twice)
