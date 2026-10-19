"""导出的 AI 提示词 / 模板，必须和界面真实措辞、问卷真实字段对得上。

为什么单独立这个文件：
模板是**客户拿去问 AI、再拿回来导入**的东西。它一旦和界面漂移，客户就会卡在
根本不存在的东西上。实测踩过两次：

1. 加了 `columns` / `header_distance_in` 两个维度，模板没跟上 → AI 填了也传不进来
   （`test_ai_template_covers_all_schema_fields` 就是那次留下的护栏，在 test_columns.py）；
2. 模板里写着 `load it with "Import AI form"` —— 界面上**没有**这个按钮：真实按钮叫
   `Import`，来源项叫 `Import configuration (AI)`。客户照着找找不到。

本文件守第 2 类：文案一致性。
"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

TEMPLATE = os.path.join(ROOT, "src", "data", "ai_questionnaire_template.yaml")


def _text() -> str:
    with open(TEMPLATE, "r", encoding="utf-8") as f:
        return f.read()


def test_prompt_points_to_real_buttons():
    """提示词里指路的按钮名 / 来源名，必须是界面上真实存在的那几个词。

    - 来源下拉那一项：`Import configuration (AI)`（i18n `src_ai`）
    - 按钮：`Import`（i18n `btn_ai_import`）
    """
    from src.ui import i18n

    src_ai = i18n.STRINGS["en"]["src_ai"]
    btn = i18n.STRINGS["en"]["btn_ai_import"]
    text = _text()

    assert "Import AI form" not in text, "提示词里还留着界面上不存在的按钮名"
    assert src_ai in text, "提示词没告诉客户该选哪个来源（%r）" % src_ai
    assert '"%s"' % btn in text, "提示词没告诉客户点哪个按钮（%r）" % btn


def test_prompt_is_english_only():
    """海外版铁律：面向客户的模板全英文，不得出现中文。"""
    text = _text()
    bad = [ch for ch in text if "\u4e00" <= ch <= "\u9fff"]
    assert not bad, "AI 模板里混入中文：%r" % "".join(bad[:20])


def test_template_states_columns_and_header_distance():
    """两个易漏维度必须在模板里写清含义（AI 才不会瞎猜）。"""
    low = _text().lower()
    assert "single column" in low and "two-column" in low, "栏数没讲清单栏/双栏含义"
    assert "header_distance_in" in low, "页眉距页顶没进模板"
    assert "omit" in low, "页眉距没有说明『不规定就别填』（否则 AI 会编一个值）"
