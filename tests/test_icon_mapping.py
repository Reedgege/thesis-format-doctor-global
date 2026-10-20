"""图标映射回归（tkinter 真窗）。

锁住 2026-10-20 老板发现的两个图标 bug：

1. 学校模板面板标题的 IconBadge 必须解析到「圆顶大楼」university_template.png
   —— v2.3.62 重做面板时误用 "doc"（→ your_documents 文档图标，与左卡大标题撞图标）。
2. 「格式来源」行的 IconBadge 必须解析到「引用格式」citation_style.png
   —— v2.3.62 把行键 row_spec_title 改名 source_title，但 _FIELD_ICONS 没跟上，
      导致键缺失、图标静默丢失。

不止查映射表，还要查**界面里真正构造出来的 IconBadge 控件**，
确保以后有人再改这两处不会静默退化。

无 tkinter / 无法建窗的环境自动 skip。
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
    monkeypatch.setenv("TFD_RECENTS_FILE", str(tmp_path / "recents.json"))
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


def _walk(widget):
    out = [widget]
    for c in widget.winfo_children():
        out.extend(_walk(c))
    return out


def _icon_badges(root):
    from src.ui.widgets import IconBadge
    return [w for w in _walk(root) if isinstance(w, IconBadge)]


def test_school_template_header_uses_university_icon(ui):
    from src.ui import icon_loader
    # 学校面板标题 IconBadge（第一个 badge）必须传 "bank"（→ university_template.png）
    badges = _icon_badges(ui._school_frame)
    assert badges, "学校面板没有任何 IconBadge"
    header = badges[0]
    assert header._icon == "bank", "学校面板标题图标应为 bank，实为 %r" % header._icon
    # 解析到的 PNG 必须真实存在且是圆顶大楼图标
    png = icon_loader.ICON_NAME_MAP["bank"]
    assert png == "university_template", png
    path = icon_loader.find_icon(png)
    assert path and path.endswith("university_template.png"), path


def test_format_source_row_uses_citation_style_icon(ui):
    from src.ui import app as ui_app, icon_loader
    # bug#2 精确锁：_FIELD_ICONS 的行键必须跟上行键改名（source_title → quote）
    assert ui_app._FIELD_ICONS.get("source_title") == "quote"
    # 「格式来源」行在**主来源选择器**（_build_source_selector），不在高级区；
    # 整棵界面里该行的 citation_style 图标必须存在。
    badges = [b for b in _icon_badges(ui.root) if b._icon == "quote"]
    assert badges, "界面里找不到格式来源行（quote）图标"
    png = icon_loader.ICON_NAME_MAP["quote"]
    assert png == "citation_style", png
    path = icon_loader.find_icon(png)
    assert path and path.endswith("citation_style.png"), path
