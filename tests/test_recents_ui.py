"""v2.3.78「最近使用」GUI 集成回归（tkinter，真起窗口）。

钉住：
- 手填问卷保存后写入 recents.json（内容 + last_source），下拉出现可点模板；
- AI 配置导入后写入 recents.json，下拉出现可点文件；
- 启动自动恢复：新建 App 后，上次的手填/AI 数据被读回、来源切回对应项；
- 点下拉即回填（manual 写回 questionnaire_data 并切到手填来源；ai 重新导入）；
- AI 文件已消失 → 点一下从历史移除；
- 切换语言重建界面后，下拉仍在且显示对应语言摘要（不漏中文到英文界面）。

无 tkinter / 无法建窗的环境自动 skip。
"""
from __future__ import annotations

import importlib.util
import json
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


def _read_recents(tmp_path):
    p = tmp_path / "recents.json"
    if not p.exists():
        return {"manual": [], "ai": [], "last_source": ""}
    return json.loads(p.read_text(encoding="utf-8"))


def test_manual_remember_writes_recents(ui, tmp_path):
    ui.questionnaire_data = {"font_family": "Arial", "font_size_pt": 11}
    ui._remember_manual()
    rec = _read_recents(tmp_path)
    assert rec["last_source"] == "src_manual"
    assert len(rec["manual"]) == 1
    assert rec["manual"][0]["data"] == {"font_family": "Arial", "font_size_pt": 11}
    # 下拉出现了可点模板（非占位「无记录」）
    vals = list(ui._manual_recent.cget("values"))
    assert vals and vals[0] != ui.tr("recents_none")


def test_ai_remember_writes_recents(ui, tmp_path):
    import pathlib
    cfg = tmp_path / "my_config.yaml"
    cfg.write_text("font_family: Arial\n", encoding="utf-8")
    ui._remember_ai(str(cfg))
    rec = _read_recents(tmp_path)
    assert rec["last_source"] == "src_ai"
    assert rec["ai"][0]["path"] == str(cfg)
    vals = list(ui._ai_recent.cget("values"))
    assert vals and vals[0] == "my_config.yaml"


def test_autorestore_manual_on_new_app(tmp_path, monkeypatch, _tk_session):
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))
    monkeypatch.setenv("TFD_RECENTS_FILE", str(tmp_path / "recents.json"))

    # 第一个 App：保存手填问卷
    from src.ui import app as ui_app
    a1 = ui_app.App(_tk_session)
    _tk_session.update_idletasks()
    a1.questionnaire_data = {"font_family": "Georgia", "font_size_pt": 12}
    a1._remember_manual()
    a1._close_popups()

    # 第二个 App：应当自动恢复
    a2 = ui_app.App(_tk_session)
    _tk_session.update_idletasks()
    assert a2.source_mode.get() == "src_manual"
    assert a2.questionnaire_data == {"font_family": "Georgia", "font_size_pt": 12}
    a2._close_popups()


def test_autorestore_ai_skips_missing_file(tmp_path, monkeypatch, _tk_session):
    monkeypatch.setenv("TFD_SETTINGS_FILE", str(tmp_path / "settings.json"))
    monkeypatch.setenv("TFD_RECENTS_FILE", str(tmp_path / "recents.json"))
    from src.ui import app as ui_app, recents

    # 写入一个指向不存在文件的 ai 记录 + last_source=src_ai
    rec = recents.load_recents()
    rec = recents.add_ai(rec, str(tmp_path / "ghost.yaml"))
    rec = recents.set_last_source(rec, "src_ai")
    recents.save_recents(rec)

    a = ui_app.App(_tk_session)
    _tk_session.update_idletasks()
    # 文件已消失 → 不应恢复（避免打不开的链接）
    assert a.source_mode.get() != "src_ai"
    a._close_popups()


def test_manual_pick_fills_and_switches(ui):
    data = {"font_family": "Calibri", "font_size_pt": 10}
    ui._manual_recent_items = [data]
    ui._manual_recent.config(values=["Calibri 10pt"])
    ui._manual_recent.current(0)
    ui._on_manual_recent_pick()
    assert ui.questionnaire_data == data
    assert ui.source_mode.get() == "src_manual"


def test_ai_pick_missing_removes_from_history(ui, tmp_path):
    ui._ai_recent_items = [str(tmp_path / "gone.yaml")]
    ui._ai_recent.config(values=["gone.yaml"])
    ui._ai_recent.current(0)
    ui._on_ai_recent_pick()
    # 历史里该路径应被移除（ai 列表清空）
    assert ui._ai_recent_items == []


def test_language_switch_keeps_combos_in_right_lang(ui, tmp_path):
    # 用会触发本地化词的数据（双栏），否则 en/zh 摘要会完全相同（字体名与 pt 是中性词）
    ui.questionnaire_data = {"font_family": "Arial", "font_size_pt": 11, "columns": "2"}
    ui._remember_manual()
    ui._set_lang("zh")
    ui.root.update_idletasks()
    vals = list(ui._manual_recent.cget("values"))
    assert vals and vals[0] != ui.tr("recents_none")
    # 中文界面摘要应含中文（双栏）
    assert any(ord(c) > 0x2E7F for c in vals[0])
    # 切回英文，不应漏中文
    ui._set_lang("en")
    ui.root.update_idletasks()
    vals_en = list(ui._manual_recent.cget("values"))
    assert vals_en and not any(ord(c) > 0x2E7F for c in vals_en[0])
