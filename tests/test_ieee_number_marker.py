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


def test_numbering_flag_off_keeps_plain_marker(tmp_path):
    """目标没要求编号（reference_numbered=False）时，1. 原样保留。"""
    src = _make(str(tmp_path / "f.docx"))
    out = str(tmp_path / "f_fixed.docx")
    t = TargetProfile(reference_style="IEEE", reference_numbered=False)
    fix_docx(src, out, t)
    assert _ref_texts(out)[0].startswith("1. ")


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
