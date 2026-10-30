# -*- coding: utf-8 -*-
"""规范 → 画像 的参考文献版面接线回归（v2.3.116）。

老板 2026-10-30：「以规范为准」——**规范显式声明了什么就判什么，不许推断**。

背景（实测缺口）：StyleSpec 里 `reference_new_page` / `reference_heading_bold` 早就
有值，但 `target_from_spec` 从没搬进 TargetProfile → 纯规范路径下检查器读到 None，
一律降级成 info「只提示」，APA 明文要求的「另起一页 / 标题居中加粗 / 条目双倍行距」
形同虚设。

⚠️ 反向红线：**dataclass 默认值不等于规范声明**。`reference_heading_bold` 原默认 True，
IEEE / Other 并没声明过它，一旦把默认值也接进画像，就等于给它们凭空加约束
（实测把 `test_fix_ieee_realistic_bold_normal_headings_numbered` 打挂）。
本文件第 2 条用例专门守住这一点。
"""
from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from src.engine import checker
from src.engine.questionnaire import build_target, target_from_spec

APA_KEYS = ("APA-Student", "APA-Professional")
UNTENDED_KEYS = ("IEEE-Conference", "IEEE-Journal", "Other")


def test_apa_spec_wires_reference_layout():
    """APA 显式声明的四项必须进画像（以前全丢）。"""
    for k in APA_KEYS:
        t = target_from_spec(k)
        assert t.reference_new_page is True, "%s: APA 要求参考文献另起一页" % k
        assert t.reference_heading_bold is True, "%s: APA 要求标题加粗" % k
        assert t.reference_heading_centered is True, "%s: APA 要求标题居中" % k
        assert t.reference_line_spacing == 2.0, "%s: APA 要求条目双倍行距" % k
        # 来源必须标 spec，报告里才不会写成「学校模板要求」
        for f in ("reference_new_page", "reference_heading_bold",
                  "reference_heading_centered", "reference_line_spacing"):
            assert t.resolution.get(f) == "spec", "%s: %s 来源应为 spec" % (k, f)


def test_unstated_specs_do_not_inherit_dataclass_defaults():
    """IEEE / Other **没有声明**这四项 → 必须是 None（不判不改）。

    这是「不推断」的红线：dataclass 默认值（如 reference_heading_bold=True）不是规范声明。
    """
    for k in UNTENDED_KEYS:
        t = target_from_spec(k)
        assert t.reference_new_page is None, "%s 未声明另起一页 → 应 None" % k
        assert t.reference_heading_bold is None, "%s 未声明标题加粗 → 应 None" % k
        assert t.reference_heading_centered is None, "%s 未声明标题居中 → 应 None" % k
        assert t.reference_line_spacing is None, "%s 未声明条目行距 → 应 None" % k


def test_mla_declared_values_wired():
    """MLA 显式声明：另起一页 + 标题居中不加粗；没声明条目行距。"""
    t = target_from_spec("MLA")
    assert t.reference_new_page is True
    assert t.reference_heading_bold is False        # forbid
    assert t.reference_heading_centered is True
    assert t.reference_line_spacing is None


def test_heading_style_is_trusted_not_flagged(tmp_path):
    """官方样例用 Heading 1 承载 'References'（粗体/居中来自样式）→ 必须判 pass。

    接线后若丢掉「信任样式」的保护，会把金标准论文误判红（负作用）。
    """
    doc = Document()
    doc.add_paragraph("Body text.")
    p = doc.add_paragraph("References")
    p.style = doc.styles["Heading 1"]
    doc.add_paragraph("Smith, J. (2020). A book. Publisher.")
    src = str(tmp_path / "ref_style.docx")
    doc.save(src)

    _, findings = checker.run_check(src, target_from_spec("APA-Student"))
    rh = [f for f in findings if f.dimension == "reference_heading"]
    assert rh and rh[0].severity == "pass", \
        "Heading 样式承载的标题应判 pass（信任样式），实际 %s" % rh


def test_manual_heading_missing_format_fails(tmp_path):
    """手动段落且未加粗未居中 → fail（APA 明文要求，可自动修正）。"""
    doc = Document()
    doc.add_paragraph("Body text.")
    doc.add_paragraph("References")
    doc.add_paragraph("Smith, J. (2020). A book. Publisher.")
    src = str(tmp_path / "ref_plain.docx")
    doc.save(src)

    _, findings = checker.run_check(src, target_from_spec("APA-Student"))
    rh = [f for f in findings if f.dimension == "reference_heading"]
    assert rh and rh[0].severity == "fail", "手动未格式化应 fail，实际 %s" % rh


def test_template_spacing_conflict_is_noted_but_template_wins(tmp_path):
    """模板条目行距与所选规范冲突时：**只提示、不改值**（模板优先铁律）。"""
    doc = Document()
    doc.add_paragraph("Body text.")
    p = doc.add_paragraph()
    r = p.add_run("Body sample.")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    doc.add_paragraph("References")
    e = doc.add_paragraph("Smith, J. (2020). A book. Publisher.")
    e.paragraph_format.line_spacing = 1.0          # 模板=单倍
    e.paragraph_format.left_indent = Pt(36)
    e.paragraph_format.first_line_indent = Pt(-36)
    tpl = str(tmp_path / "tpl.docx")
    doc.save(tpl)

    t = build_target(template=tpl, reference_style="APA")
    # 模板优先：值保持 1.0，不被改成 APA 的 2.0
    assert t.reference_line_spacing == 1.0, "模板值应优先，实测 %r" % (t.reference_line_spacing,)
    notes = " ".join(t.notes or [])
    assert "line spacing" in notes.lower() and "APA" in notes, \
        "应提示模板行距与 APA 不一致，实际 notes=%r" % (t.notes,)
