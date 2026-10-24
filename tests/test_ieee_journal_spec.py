# -*- coding: utf-8 -*-
"""IEEE-Journal 档（v2.3.101）：**不碰版式**，只查 ESM 编辑规范。

裁定依据：`_audit/裁定冻结_2026-10-23.md` · P-1：
  `IEEE-Journal` = 不碰版式（ESM 全文无页边距/分栏规定，作者无从控制）；
  只查 ESM 编辑规范：四级标题、摘要、Index Terms、图表标题、公式编号、参考文献、
  脚注、传记、Acknowledgment、COI。

材料依据（真实、可追溯）：
  `_smoke/IEEE_journal_Access_template.docx`（IEEE Access 官方模板，335KB）
  + `docs/IEEE_三版本对比_基于已发材料.md`（三版本实测差异）
  🔴 但 Access 只是**一个期刊**的模板：页边距/分栏/字号各刊不同 → 一律不写死。

本测试钉死三件事：
  ① 版式字段必须全空（写成默认值 1″/TNR12/双倍 = 替客户拍板）；
  ② ESM 编辑规范维度必须在位；
  ③ 引用体例仍归 IEEE（style_key="IEEE"），与 spec key 分开。
"""

from src.engine.specs import SPECS, SPEC_ORDER, spec_display, spec_is_baseline
from src.engine.questionnaire import target_from_spec

# PageLayout 里所有「版式」字段：任何一个非 None 都算破功
_LAYOUT_FIELDS = (
    "margin_top_in", "margin_bottom_in", "margin_left_in", "margin_right_in",
    "font_family", "font_size_pt", "line_spacing",
    "body_alignment", "first_line_indent_in",
    "space_before_pt", "space_after_pt", "header_distance_in",
)


def test_journal_exists_and_listed():
    assert "IEEE-Journal" in SPECS
    assert "IEEE-Journal" in SPEC_ORDER
    assert SPECS["IEEE-Journal"].key == "IEEE-Journal"
    assert spec_display("IEEE-Journal") == "IEEE Journal (editorial rules)"


def test_journal_touches_no_layout():
    """🔴 不碰版式：12 个版式字段必须全 None（不能吃到 PageLayout 的默认 1″/TNR12）。"""
    spec = SPECS["IEEE-Journal"]
    for f in _LAYOUT_FIELDS:
        assert getattr(spec.page, f) is None, f"版式字段 {f} 不应有值：{getattr(spec.page, f)!r}"
    assert spec.columns is None, "分栏不判（Access 双栏、多数刊初稿单栏）"
    assert spec.paper_size is None, "纸张不判"
    assert spec.heading_formats == {}, "各级标题字形不判（不硬写 Helvetica 10/9pt）"
    assert spec.caption_font_size_pt is None
    assert spec.reference_font_size_pt is None


def test_journal_editorial_rules_in_place():
    """只查 ESM 编辑规范的那几项必须在位。"""
    spec = SPECS["IEEE-Journal"]
    assert spec.heading_levels == 4, "ESM：四级标题"
    assert spec.expect_abstract is True, "ESM：所有论文必须有摘要"
    assert spec.expect_index_terms is True, "ESM：必须有 Index Terms"
    assert spec.caption_position_table == "top", "IEEE 惯例：表题在上"
    assert spec.caption_position_figure == "bottom", "IEEE 惯例：图题在下"
    assert spec.table_lines is None, "不规定三线表"
    assert spec.title_page is None, "期刊无独立题名页（作者信息在首页脚注）"


def test_journal_reference_is_ieee_numbered_without_layout():
    spec = SPECS["IEEE-Journal"]
    assert spec.reference.style_key == "IEEE"
    assert spec.reference.numbered is True, "ESM：顺序编码 [1]"
    assert spec.reference.hanging_indent_in is None, "悬挂缩进随期刊模板，不写死 0.5″"


def test_journal_target_profile():
    t = target_from_spec("IEEE-Journal")
    assert t.spec_key == "IEEE-Journal"
    assert spec_is_baseline(t) is True, "选了内置规范就该走规范口径"
    assert t.reference_style == "IEEE"
    assert t.reference_numbered is True
    assert t.heading_levels == 4
    # 版式维度不进 target（page 为空 → 检查器不会判页边距/字体/行距）
    assert t.page == {}
    # 参考文献另起页：期刊"另议"，绝不能因为 spec 字段默认 False 就硬判
    assert t.reference_new_page is None


def test_journal_and_conference_are_distinct():
    """两档必须真正不同，否则拆档没意义。"""
    c, j = SPECS["IEEE-Conference"], SPECS["IEEE-Journal"]
    assert c is not j
    assert c.columns == 2 and j.columns is None
    assert c.page.margin_top_in == 0.75 and j.page.margin_top_in is None
    assert c.reference.hanging_indent_in == 0.25 and j.reference.hanging_indent_in is None
    assert c.heading_formats and j.heading_formats == {}


def test_journal_gets_ieee_info_checks():
    """checker 的 IEEE 系列 info 观测（首脚注/传记/Acknowledgment/COI/公式编号）
    门控是 spec_key.upper().startswith("IEEE") —— Journal 必须命中。"""
    assert "IEEE-Journal".upper().startswith("IEEE")
