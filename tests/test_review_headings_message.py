# -*- coding: utf-8 -*-
"""复核弹窗「Headings」文案回归（v2.3.94，老板 2026-10-23 截图圈出）。

老板截图：格式来源选「MLA 9th — Standard (Word)」，Headings 一栏写着
**「Not detected — confirm or fill in」**（红圈标出）。

这是错的，而且是两处错：

1. **MLA 9th 自己就不定义分级编号标题**（``specs.py`` 里 ``heading_levels=0``，
   notes 原文 "Numbered headings are not required (use a consistent format of your
   own)"）。所以「没检测到」是假话 —— 工具什么都没漏，是这份规范不要求。
   「规范不要求」与「没检测到」的客户动作**完全相反**：前者什么都不用做，
   后者要客户去补。
2. 就算真缺，这个弹窗里标题表是**只读**的 —— 让客户 "fill in" 是指点一个
   它自己做不到的动作。必须说清去哪儿填（Advanced options 里的手动问卷）。

本文件钉住三档分支：
  * 内置规范且定义了标题（APA / IEEE）→ 照常出每级表格，一条提示都不许有；
  * 内置规范但不定义标题（MLA）→ 「不要求，无需填写」，且**全弹窗不许出现**
    「Not detected」；
  * 其他来源（学校模板 / 手填 / AI / Other）读不到标题 → 提示「去哪儿填」。
"""
from __future__ import annotations

import importlib.util
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

NOT_REQUIRED = "does not require numbered heading levels"
GO_FILL = "fill them in under Advanced options"
OLD_WRONG = "Not detected — confirm or fill in"


@pytest.fixture(scope="session")
def _tk_session():
    if importlib.util.find_spec("tkinter") is None:
        pytest.skip("本机 Python 无 tkinter（GUI 回归需带 tkinter 的解释器）")
    import tkinter as tk

    try:
        root = tk.Tk()
    except Exception as e:                     # noqa: BLE001
        pytest.skip("Tk 不可用：%s" % e)
    yield root
    try:
        root.destroy()
    except Exception:
        pass


@pytest.fixture
def ui(tmp_path, monkeypatch, _tk_session):
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))
    for w in list(_tk_session.winfo_children()):
        try:
            w.destroy()
        except Exception:
            pass
    from src.ui import app as ui_app
    # 掐掉启动授权核验（同 test_dialog_sticky_actions：不真的起线程/轮询）
    monkeypatch.setattr(ui_app.App, "_refresh_license_at_launch",
                        lambda self: None, raising=False)
    monkeypatch.setattr(ui_app.App, "_start_launch_check",
                        lambda self: None, raising=False)
    monkeypatch.setattr(ui_app.App, "_poll_launch_check",
                        lambda self: None, raising=False)
    a = ui_app.App(_tk_session)
    _tk_session.update_idletasks()
    yield a
    try:
        a._close_popups()
    except Exception:
        pass


def _label_texts(win):
    """收集弹窗里所有 Label 的文字（含嵌套 Frame 里的）。"""
    out = []

    def walk(w):
        for ch in w.winfo_children():
            try:
                if ch.winfo_class() == "Label":
                    out.append(str(ch.cget("text")))
            except Exception:
                pass
            walk(ch)

    walk(win)
    return out


def _open(ui, target):
    ui._confirm_dialog(target)
    wins = [w for w in ui.root.winfo_children() if w.winfo_class() == "Toplevel"]
    assert wins, "没有找到确认弹窗"
    win = wins[-1]
    win.update_idletasks()
    return win


def test_mla_says_not_required_instead_of_not_detected(ui):
    """MLA：说「不要求」，绝不说「没检测到」（老板截图的那一条）。"""
    from src.engine.questionnaire import target_from_spec
    texts = _label_texts(_open(ui, target_from_spec("MLA")))
    joined = "\n".join(texts)
    assert NOT_REQUIRED in joined, "MLA 应说明该规范不要求分级标题，实际标签：%s" % texts
    assert "MLA 9th" in joined
    # 🔴 关键：整张弹窗都不许出现「没检测到」
    assert OLD_WRONG not in joined, "MLA 下不许出现「Not detected」：%s" % texts


def test_apa_and_ieee_still_show_the_heading_table(ui):
    """APA / IEEE 定义了 5 级标题 → 照常出表格，不许出现任何「缺失」提示。"""
    from src.engine.questionnaire import target_from_spec
    for key in ("APA-Student", "IEEE"):
        texts = _label_texts(_open(ui, target_from_spec(key)))
        joined = "\n".join(texts)
        assert NOT_REQUIRED not in joined
        assert GO_FILL not in joined
        assert OLD_WRONG not in joined
        assert "Level" in texts, "%s 应出标题表（含 Level 表头）" % key


def test_missing_headings_points_to_where_you_can_fill_them(ui):
    """学校模板 / 手填 / AI / Other 读不到标题 → 说清去哪儿填，不说「confirm or fill in」。"""
    from src.engine.questionnaire import target_from_dict, target_from_spec
    cases = {
        "Other": target_from_spec("Other"),
        "school": target_from_dict({"margin_top_in": 1.0}, "school"),
        "user": target_from_dict({"margin_top_in": 1.0}, "user"),
        "ai_import": target_from_dict({"margin_top_in": 1.0}, "ai_import"),
    }
    for name, target in cases.items():
        texts = _label_texts(_open(ui, target))
        joined = "\n".join(texts)
        assert GO_FILL in joined, "%s 应提示去哪儿填标题：%s" % (name, texts)
        assert NOT_REQUIRED not in joined, "%s 不是内置规范，不许说「不要求」" % name
