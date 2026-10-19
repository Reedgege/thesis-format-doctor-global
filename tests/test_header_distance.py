"""页眉距页顶（header_distance_in）全链路测试：问卷/AI 能填 → 检查能判 → 修正能改。

此前该维度只能由学校模板读出、或在复核弹窗手改；问卷与 AI 模板都提供不了。
本次补齐后它必须是真的可用项（不是「填了只存不用」的死字段）。
"""

from __future__ import annotations

import os
import sys

from docx import Document
from docx.shared import Inches

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.engine.docx_reader import read_docx
from src.engine.questionnaire import (
    build_target, target_from_dict, target_from_ai, TargetProfile,
)
from src.engine.checker import run_check
from src.engine.fixer import fix_docx, compute_changes


def _make_paper(path: str, header_distance_in: float = None) -> str:
    doc = Document()
    doc.add_paragraph("A Study of Header Distance")
    for _ in range(3):
        doc.add_paragraph("Body paragraph with enough words to look like a real paper, "
                          "repeated a few times for a plausible body.")
    doc.add_paragraph("References")
    doc.add_paragraph("[1] J. Smith, \"A title,\" J. Name, vol. 1, 2020.")
    if header_distance_in is not None:
        for sec in doc.sections:
            sec.header_distance = Inches(header_distance_in)
    doc.save(path)
    return path


# ------------------------------------------------------------------ 输入入口
def test_target_from_dict_reads_header_distance():
    t = target_from_dict({"header_distance_in": 0.5})
    assert abs(float(t.page["header_distance_in"]) - 0.5) < 1e-9
    assert t.resolution.get("header_distance_in") == "questionnaire"


def test_target_from_ai_reads_header_distance(tmp_path):
    p = str(tmp_path / "ai_hd.yaml")
    with open(p, "w", encoding="utf-8") as f:
        f.write("filled:\n  header_distance_in: 0.5\n  reference_style: MLA\n")
    t = target_from_ai(p)
    assert abs(float(t.page["header_distance_in"]) - 0.5) < 1e-9
    assert t.resolution.get("header_distance_in") == "ai_import"


def test_questionnaire_schema_has_header_distance():
    from src.engine.questionnaire import QUESTIONNAIRE_SCHEMA
    keys = {f["key"] for f in QUESTIONNAIRE_SCHEMA}
    assert "header_distance_in" in keys
    # 默认必须留空：多数格式不规定，不能强塞一个值
    field = next(f for f in QUESTIONNAIRE_SCHEMA if f["key"] == "header_distance_in")
    assert field.get("default") in ("", None)


# ------------------------------------------------------------------ 检查层
def test_reader_reads_header_distance(tmp_path):
    p = _make_paper(str(tmp_path / "hd.docx"), header_distance_in=0.2)
    prof = read_docx(p)
    assert prof.header_distance_in is not None
    assert abs(prof.header_distance_in - 0.2) < 0.02


def test_checker_flags_header_distance_mismatch(tmp_path):
    p = _make_paper(str(tmp_path / "hd2.docx"), header_distance_in=0.2)
    target = target_from_dict({"header_distance_in": 0.5})   # 非 spec 来源 → 硬判
    _, findings = run_check(p, target)
    hd = [f for f in findings if f.dimension == "header_distance"]
    assert len(hd) == 1
    assert hd[0].severity in ("warn", "fail")
    # 必须绑定修正动作，否则就是「只报不改」
    assert hd[0].fix is not None and hd[0].fix.dimension == "header_distance"


def test_checker_skips_header_distance_when_target_has_none(tmp_path):
    p = _make_paper(str(tmp_path / "hd3.docx"), header_distance_in=0.2)
    target = build_target("APA-Student")   # APA 不规定页眉距页顶 → None
    assert target.page.get("header_distance_in") is None
    _, findings = run_check(p, target)
    assert not [f for f in findings if f.dimension == "header_distance"]


# ------------------------------------------------------------------ 修正层
def test_fixer_applies_header_distance(tmp_path):
    src = _make_paper(str(tmp_path / "in.docx"), header_distance_in=0.2)
    out = str(tmp_path / "out.docx")
    target = target_from_dict({"header_distance_in": 0.5})
    _, findings = run_check(src, target)
    res = fix_docx(src, out, target, findings=findings)
    assert abs(float(res["applied"].get("header_distance_in")) - 0.5) < 1e-9
    got = Document(out).sections[0].header_distance
    assert abs(got.inches - 0.5) < 0.02


def test_fixer_header_distance_idempotent(tmp_path):
    src = _make_paper(str(tmp_path / "in2.docx"), header_distance_in=0.5)
    out = str(tmp_path / "out2.docx")
    target = target_from_dict({"header_distance_in": 0.5})
    _, findings = run_check(src, target)
    res = fix_docx(src, out, target, findings=findings)
    assert "header_distance_in" not in res["applied"]


# ------------------------------------------------------------------ 预览层
def test_compute_changes_lists_header_distance(tmp_path):
    p = _make_paper(str(tmp_path / "c.docx"), header_distance_in=0.2)
    lines = compute_changes(p, target_from_dict({"header_distance_in": 0.5}))
    assert any(("Header distance" in ln or "页眉距页顶" in ln) for ln in lines)
