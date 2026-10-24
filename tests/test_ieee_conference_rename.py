# -*- coding: utf-8 -*-
"""IEEE 拆档：Conference 改名/迁移 + Journal [待建] 不得可选。

裁定依据（冻结，不许自行改口）：
  `_audit/裁定冻结_2026-10-23.md` · P-1 · IEEE 拆两档：
    - `IEEE-Conference`：本轮唯一实现目标（现有单档 IEEE 改名/迁移而来）。
    - `IEEE-Journal`：**[待建]**，建成前不得被任何输入通道选中；不做双档并存。

关键边界：**spec key** 与 **reference_style（引用体例）key** 是两件事。
  本测试同时钉死：引用体例 key 必须仍是 "IEEE"，否则参考文献核对的入口会被一起改坏。
"""

from src.engine.specs import (
    SPECS, SPEC_ORDER, REFERENCE_STYLE_KEYS, get_spec, list_specs,
    spec_display, spec_is_baseline,
)
from src.engine.questionnaire import target_from_spec


def test_spec_order_uses_ieee_conference():
    assert "IEEE-Conference" in SPEC_ORDER
    assert "IEEE" not in SPEC_ORDER, "旧 key 不得进下拉白名单（避免出现两个 IEEE 条目）"
    # 下拉实际渲染的就是 SPEC_ORDER，list_specs 必须同步
    keys = [s.key for s in list_specs()]
    assert "IEEE-Conference" in keys and "IEEE" not in keys


def test_canonical_spec_key_and_name():
    spec = SPECS["IEEE-Conference"]
    assert spec.key == "IEEE-Conference"
    assert spec.name == "IEEE Conference (two-column)"
    assert spec_display("IEEE-Conference") == "IEEE Conference (two-column)"


def test_old_ieee_key_is_compat_alias_to_conference():
    """旧配置 / 旧调用仍能解析到会议档（与 APA → APA-Student 同一套路）。"""
    assert "IEEE" in SPECS
    assert SPECS["IEEE"] is SPECS["IEEE-Conference"]
    assert get_spec("IEEE") is get_spec("IEEE-Conference")


def test_target_from_spec_both_keys_resolve_to_conference():
    t_new = target_from_spec("IEEE-Conference")
    t_old = target_from_spec("IEEE")
    assert t_new.spec_key == "IEEE-Conference"
    # 两条路都必须判为「真的选了内置规范」，否则标题页/页面/标题层级会被降级
    assert spec_is_baseline(t_new) is True
    assert spec_is_baseline(t_old) is True


def test_reference_style_key_stays_ieee():
    """🔴 硬边界：引用体例 key 不能跟着 spec key 一起改名。"""
    assert "IEEE" in REFERENCE_STYLE_KEYS
    assert "IEEE-Conference" not in REFERENCE_STYLE_KEYS
    assert SPECS["IEEE-Conference"].reference.style_key == "IEEE"


def test_ieee_journal_is_not_selectable_yet():
    """[待建] 档位在建成前不得出现在任何可选入口。"""
    assert "IEEE-Journal" not in SPECS
    assert "IEEE-Journal" not in SPEC_ORDER
    assert all(s.key != "IEEE-Journal" for s in list_specs())


def test_conference_params_survive_rename():
    """改名不能丢参数：会议档冻结值仍须在位。"""
    spec = SPECS["IEEE-Conference"]
    assert spec.columns == 2
    assert spec.title_page.title_size_pt == 24.0
    assert spec.title_page.title_bold is False
    assert spec.byline_size_pt == 11.0
    assert spec.abstract_font_size_pt == 9.0
    assert spec.reference_font_size_pt == 8.0
    assert spec.page.body_alignment == "justify"
    assert abs(spec.page.first_line_indent_in - 0.2) < 1e-9
