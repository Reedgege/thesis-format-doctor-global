"""海外版 v2.3.19「× 清除」按钮回归测试（tkinter，真起窗口）。

钉住论文 / 学校模板 / LaTeX / AI 配置 四个导入行的小「✕」清除行为：
· 选中后对应 ✕ 出现；
· 点 ✕ 清空该选择（StringVar 置空）；
· 已做的检查/修正结论随之作废、回到初始态；
· 只清单项时保留其它输入；
· ✕ 在清空后自动隐藏。

无 tkinter / 无法建窗的环境自动 skip（见 ``_tk_session``）。
"""
from __future__ import annotations

import importlib.util
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


@pytest.fixture(scope="session")
def _tk_session():
    """整个会话共用一个 Tk root（与 test_ui_redesign 同款）。无法建窗则 skip。"""
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
    """建一个干净的 App（每个用例重建界面，避免状态串味）。"""
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))
    for w in list(_tk_session.winfo_children()):
        try:
            w.destroy()
        except Exception:
            pass
    from src.ui import app as ui_app
    a = ui_app.App(_tk_session)
    _tk_session.update_idletasks()
    yield a
    try:
        a._close_popups()
    except Exception:
        pass


def _visible(w):
    try:
        return w.winfo_manager() != ""
    except Exception:
        return False


def test_paper_clear_shows_after_select_and_hides_after_clear(ui):
    """选论文后投放区 ✕ 出现；点 ✕ 清空论文且 ✕ 消失。"""
    ui.docx_path.set("C:/tmp/thesis.docx")
    ui._refresh_rows()
    assert _visible(ui._paper_clear), "选中论文后 ✕ 应显示"

    ui._clear_input("paper")
    assert ui.docx_path.get() == "", "点 ✕ 后论文路径应清空"
    assert not _visible(ui._paper_clear), "清空后 ✕ 应隐藏"


def test_template_clear_preserves_other_inputs(ui):
    """清模板时论文（若存在）应保留；模板 ✕ 正确显隐。"""
    ui.docx_path.set("C:/tmp/thesis.docx")
    ui.school_template.set("C:/tmp/tpl.docx")
    ui._refresh_rows()
    assert _visible(ui._tpl_clear), "选中模板后 ✕ 应显示"

    ui._clear_input("template")
    assert ui.school_template.get() == "", "模板应被清空"
    assert ui.docx_path.get() != "", "清模板不应影响已选论文"
    assert not _visible(ui._tpl_clear), "清空后模板 ✕ 应隐藏"


def test_latex_and_ai_clear(ui):
    ui.latex_template.set("C:/tmp/main.tex")
    ui.ai_json.set("C:/tmp/ai.json")
    ui._refresh_rows()
    assert _visible(ui._latex_clear), "选中 LaTeX 后 ✕ 应显示"
    assert _visible(ui._ai_clear), "选中 AI 配置后 ✕ 应显示"

    ui._clear_input("latex")
    assert ui.latex_template.get() == "", "LaTeX 应被清空"
    assert not _visible(ui._latex_clear), "清空后 LaTeX ✕ 应隐藏"

    ui._clear_input("ai")
    assert ui.ai_json.get() == "", "AI 配置应被清空"
    assert not _visible(ui._ai_clear), "清空后 AI ✕ 应隐藏"


def test_clear_invalidates_results(ui):
    """已有体检结论时清输入，结论应作废、phase 回到 ready。"""
    ui.docx_path.set("C:/tmp/thesis.docx")
    ui._phase = "results"
    ui._last_issues = 3
    ui._refresh_rows()

    ui._clear_input("paper")
    assert ui._phase in ("ready", "idle"), "清论文后结论应作废、回到初始"
    assert ui._last_issues is None, "旧结论计数应被清空"
