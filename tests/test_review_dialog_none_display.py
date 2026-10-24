# -*- coding: utf-8 -*-
"""复核弹窗「不规定（None）≠ 具体值」回归（v2.3.102，老板 2026-10-24 截图）。

抓到两个 bug（同一类：把 None 当成一个具体值在展示）：
  ① **Headings**：IEEE-Journal 明明 `heading_levels=4`（只是按裁定不规定各级字形），
     弹窗却说 "this source sets no heading levels" —— 与事实相反，客户会以为要自己补。
  ② **Columns**：`columns=None`（不规定栏数）被显示成 "Single column"，且确认时回写 "1"。
     `apply_page_overrides` 只把**空串**当"未改动" → None 被静默升级成「强制单栏」，
     双栏期刊论文一检查就报"应为单栏"。
"""

from src.engine.questionnaire import apply_page_overrides, target_from_spec
from src.engine.specs import get_spec, spec_is_baseline
from src.ui import i18n
from src.ui.app import _columns_display_value, _columns_edit_value

NONE_TXT, ONE_TXT, TWO_TXT = "—", "Single column", "Two columns"


# —— ② Columns ——

def test_columns_none_displays_dash_not_single():
    assert _columns_display_value(None, NONE_TXT, ONE_TXT, TWO_TXT) == NONE_TXT
    assert _columns_display_value(1, NONE_TXT, ONE_TXT, TWO_TXT) == ONE_TXT
    assert _columns_display_value(2, NONE_TXT, ONE_TXT, TWO_TXT) == TWO_TXT


def test_columns_dash_edits_to_empty_string():
    assert _columns_edit_value(NONE_TXT, NONE_TXT, TWO_TXT) == ""
    assert _columns_edit_value(ONE_TXT, NONE_TXT, TWO_TXT) == "1"
    assert _columns_edit_value(TWO_TXT, NONE_TXT, TWO_TXT) == "2"


def test_journal_columns_none_survives_confirm():
    """「—」→ 空串 → apply_page_overrides 视为未改动 → 仍是 None（不升级成单栏）。"""
    t = target_from_spec("IEEE-Journal")
    assert t.columns is None
    apply_page_overrides(t, {"columns": _columns_edit_value(NONE_TXT, NONE_TXT, TWO_TXT)})
    assert t.columns is None, "「不规定栏数」被静默升级成了单栏"
    assert "columns" not in t.resolution


def test_explicit_column_change_still_applies():
    """反向守卫：别把修复做成"永不写回"——客户真的改成双栏时必须生效。"""
    t = target_from_spec("APA-Student")
    apply_page_overrides(t, {"columns": _columns_edit_value(TWO_TXT, NONE_TXT, TWO_TXT)})
    assert t.columns == 2
    assert t.resolution.get("columns") == "user"


# —— ① Headings ——

def test_headings_levels_only_message_exists_in_both_langs():
    for lang in ("en", "zh"):
        d = i18n.STRINGS[lang]
        assert "review_headings_levels_only" in d, "缺 %s 文案" % lang
        txt = d["review_headings_levels_only"]
        assert "{name}" in txt and "{n}" in txt
        out = txt.format(name="IEEE Journal (editorial rules)", n=4)
        # 不能再退回到「不设标题层级」的说法
        assert "sets no heading levels" not in out


def test_journal_hits_the_levels_only_branch():
    """弹窗分支条件：spec_is_baseline=True 且 heading_levels>0 → levels_only，不是 missing。"""
    t = target_from_spec("IEEE-Journal")
    assert spec_is_baseline(t) is True
    assert int(getattr(get_spec(t.spec_key), "heading_levels", 0) or 0) == 4


def test_mla_still_hits_not_required_branch():
    """MLA 仍是「规范本身不要求分级标题」——别被这次改动牵连。"""
    t = target_from_spec("MLA")
    assert spec_is_baseline(t) is True
    assert int(getattr(get_spec(t.spec_key), "heading_levels", 0) or 0) == 0
