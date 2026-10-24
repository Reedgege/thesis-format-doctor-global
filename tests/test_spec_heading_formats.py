"""内置规范的「每级标题格式」必须真的进入目标画像、参与检查与修正（v2.3.69）。

老板 2026-10-20：「ieee apa mla这些都是有具体要求的」。

此前 ``target_from_spec`` 只带了 ``heading_levels``（层级数）、漏传
``heading_formats``（每级具体格式），造成两处脱节：
  ① 手填问卷选 APA/IEEE 后「每级标题明细」表全空，客户以为规范没要求；
  ② 检查器的每级标题比对整段被跳过（只在 heading_formats 非空时才比），
     而修正器自己另读 spec.heading_formats 照改 —— 正是代码注释警告过的
     「检查说没事、修正却改一批」。

MLA / Other 官方没有多级标题体系（heading_levels=0），必须保持为空、不得臆造。
"""
from __future__ import annotations

import pytest
from docx import Document

from src.engine.checker import run_check
from src.engine.fixer import fix_docx
from src.engine.questionnaire import target_from_spec


# ---------------------------------------------------------------- 画像层
def test_target_from_spec_carries_per_level_heading_formats():
    """APA / IEEE 的五级格式必须随目标画像一起出来。"""
    for key in ("APA", "APA-Student", "APA-Professional", "IEEE"):
        hf = target_from_spec(key).heading_formats
        # v2.3.100（C 批 D-11 / 裁定 Q-1）：IEEE 官方 ESM「There are four levels」，
        # 故 IEEE 只定义 1-4 级；其余规范仍是 5 级。
        _want = {1, 2, 3, 4} if key == "IEEE" else {1, 2, 3, 4, 5}
        assert set(hf) == _want, f"{key} 每级标题格式缺失：{sorted(hf)}"
        assert all(isinstance(v, dict) for v in hf.values())


def test_mla_and_other_have_no_heading_levels():
    """MLA / Other 官方无多级标题体系 —— 必须留空，不许臆造。"""
    for key in ("MLA", "Other"):
        t = target_from_spec(key)
        assert not t.heading_formats, f"{key} 不该有每级标题格式"
        assert not t.heading_levels, f"{key} 不该强制标题层级"


def test_ieee_heading_values_match_spec_table():
    """IEEE：按官方 ESM + 会议模板 Heading1~4 样式（v2.3.100 重写，裁定 Q-1）。

    ESM 原文：一级「Roman numerals, centered, small caps（无 bold）」；
    二级「capital letters + periods, flush left, italic, upper and lowercase」；
    三级「indented one em, run into the text, italic, followed by a colon」；
    四级「indented two ems, lowercase letters, first letter capitalized」。
    只有**四级**（"There are four levels of section headings"）。
    """
    hf = target_from_spec("IEEE").heading_formats
    assert set(hf) == {1, 2, 3, 4}
    # L1：居中 + small caps（**不加粗**）+ 罗马编号
    assert hf[1]["alignment"] == "center"
    assert hf[1]["bold"] is False
    assert hf[1]["small_caps"] is True
    assert hf[1]["number_prefix"] == "roman"
    assert hf[1]["uppercase"] is False
    # L2：左对齐 + 斜体 + **不加粗** + 大写字母
    assert hf[2]["alignment"] == "left"
    assert hf[2]["bold"] is False
    assert hf[2]["italic"] is True
    assert hf[2]["number_prefix"] == "upper_letter"
    # L3/L4：缩进 + 冒号接排（run-in）
    for lvl, prefix, ind in ((3, "number_paren", 0.2), (4, "lower_letter", 0.35)):
        assert hf[lvl]["number_prefix"] == prefix
        assert hf[lvl]["italic"] is True
        assert hf[lvl]["alignment"] == "left"
        assert hf[lvl]["run_in"] is True
        assert hf[lvl]["indent_in"] == pytest.approx(ind)


def test_apa_heading_values_match_spec_table():
    """APA：L1 居中加粗；L2 左对齐加粗；L3 加粗斜体；L4/L5 缩进接排。"""
    hf = target_from_spec("APA-Student").heading_formats
    assert hf[1]["alignment"] == "center" and hf[1]["bold"] is True
    assert hf[2]["alignment"] == "left" and hf[2]["bold"] is True
    assert hf[3]["bold"] is True and hf[3]["italic"] is True
    assert hf[4]["run_in"] is True and float(hf[4]["indent_in"]) == 0.5
    assert hf[5]["run_in"] is True and hf[5]["italic"] is True


def test_resolution_marks_heading_formats_as_spec():
    """来源标注必须是 spec —— 检查器据此把发现降级，不硬判客户失败。"""
    assert target_from_spec("IEEE").resolution.get("heading_formats") == "spec"
    assert target_from_spec("APA-Student").resolution.get("heading_formats") == "spec"


# ---------------------------------------------------------------- 端到端
def _doc_with_left_aligned_l1(tmp_path):
    """一份 L1 标题左对齐、无编号、非全大写的文档（与 APA/IEEE 要求都不符）。"""
    doc = Document()
    h = doc.add_heading("Introduction", level=1)
    h.alignment = None          # 交还给样式默认（样板里 Heading 1 是左对齐）
    doc.add_paragraph("Body text for reading.")
    doc.add_paragraph("References")
    doc.add_paragraph("Smith, J. (2020). Title of the work. Publisher.")
    out = tmp_path / "heading_bad.docx"
    doc.save(str(out))
    return str(out)


def test_checker_now_compares_spec_heading_formats(tmp_path):
    """spec 来源也要比对每级标题格式（此前整段跳过）。"""
    path = _doc_with_left_aligned_l1(tmp_path)
    _prof, findings = run_check(path, target_from_spec("IEEE"))
    per_level = [f for f in findings
                 if f.dimension == "heading" and f.message.startswith("Level 1 ")]
    assert per_level, "spec 来源的每级标题格式没有被检查（旧缺口回归）"


def test_spec_heading_findings_never_fail_the_paper(tmp_path):
    """规范来源不硬判失败：IEEE 只能 info，APA 最高 warn，绝不出 fail。"""
    path = _doc_with_left_aligned_l1(tmp_path)
    for key, allowed in (("IEEE", {"info", "pass"}), ("APA-Student", {"warn", "info", "pass"})):
        _prof, findings = run_check(path, target_from_spec(key))
        per_level = [f for f in findings
                     if f.dimension == "heading" and f.message.startswith("Level 1 ")]
        assert per_level, f"{key} 未产生每级标题发现"
        assert not any(f.severity == "fail" for f in per_level), \
            f"{key} 的每级标题发现不该是 fail：{[f.severity for f in per_level]}"
        assert all(f.severity in allowed for f in per_level), \
            f"{key} 发现严重度越界：{[f.severity for f in per_level]}"


def test_fix_then_recheck_converges_for_spec(tmp_path):
    """修完再查必须收敛 —— check→fix 单一真相源（旧版检查跳查时这条形同虚设）。"""
    src = _doc_with_left_aligned_l1(tmp_path)
    out = str(tmp_path / "fixed.docx")
    fix_docx(src, out, target_from_spec("IEEE"))
    _prof, findings = run_check(out, target_from_spec("IEEE"))
    bad = [f for f in findings
           if f.dimension == "heading" and f.message.startswith("Level 1 ")
           and f.severity == "fail"]
    assert not bad, "修完再查仍是 fail：检查与修正不一致"
