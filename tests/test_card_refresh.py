"""圆角卡片「随内容变高」回归测试（需要可显示的 Tk 环境；无显示时自动 skip）。

背景（老板 2026-10-04 报障）：v2.3.50 删掉 RoundCard.refresh 里的
``inner.update_idletasks()`` 后，运行期「非换行类」内容变化（高级选项展开新增整行、
导入论文后进度条/状态行出现）不会触发重新测量 → 卡片量到旧高度 → 新内容被圆角矩形
裁掉、滚不到。本测试把这条路径钉死：往卡片里追加内容并 refresh，实测高度必须增大。

运行：在有桌面的机器上用带 tkinter 的 Python 跑（CI 无显示会 skip）。
"""
from __future__ import annotations

import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

try:
    import tkinter as tk
except Exception:  # pragma: no cover
    pytest.skip("本机 Python 无 tkinter", allow_module_level=True)


@pytest.fixture(scope="module")
def viewable_root():
    """建一个真正可见的 Tk 根窗口；不可见则整模块 skip。"""
    # v2.3.53：建窗口必须容错。CI（无桌面会话 / 已有其它 Tk 根）下 ``tk.Tk()``
    # 可能直接抛 TclError —— 不做处理就是一条**硬 error**（把整个 CI 变红、卡住出包），
    # 而这里要表达的语义是"环境不具备就跳过"，与 test_ui_redesign 的
    # ``_tk_session`` 一致：取不到 Tk 就 skip，不是失败。
    try:
        root = tk.Tk()
    except Exception as e:  # pragma: no cover - 仅在异常环境下命中
        pytest.skip("Tk 不可用：%s" % e)
    try:
        root.geometry("820x640")
        root.deiconify()
        root.update_idletasks()
    except Exception as e:  # pragma: no cover
        try:
            root.destroy()
        except Exception:
            pass
        pytest.skip("窗口初始化失败：%s" % e)
    if not root.winfo_viewable():
        root.destroy()
        pytest.skip("无可见显示环境（headless），跳过卡片几何测试")
    yield root
    try:
        root.destroy()
    except Exception:
        pass


def test_card_grows_when_content_added(viewable_root):
    from src.ui import widgets
    from src.ui import theme

    host = tk.Frame(viewable_root, bg=theme.SURFACE)
    host.pack(fill="both", expand=True, padx=20, pady=20)
    card = widgets.RoundCard(host, padx=theme.CARD_PAD_X, pady=theme.CARD_PAD_Y)
    card.pack(fill="x")
    viewable_root.update_idletasks()

    # 初始内容：一行
    tk.Label(card.inner, text="Row 1").pack(anchor="w")
    card.refresh()
    viewable_root.update_idletasks()
    h_before = card._req[1]

    # 模拟「高级选项展开 / 导入论文」追加整行内容
    tk.Label(card.inner, text="Row 2 (added by advanced expand / import)").pack(anchor="w")
    card.refresh()
    viewable_root.update_idletasks()
    h_after = card._req[1]

    assert h_after > h_before, (
        f"卡片没随内容变高：{h_before}px -> {h_after}px（v2.3.50 回归未修）"
    )


def test_refresh_remeasures_when_viewable(viewable_root):
    """直接钉住根因：运行期窗口可见时，refresh 必须调 inner.update_idletasks 重新测量。"""
    from src.ui import widgets
    from src.ui import theme

    host = tk.Frame(viewable_root, bg=theme.SURFACE)
    host.pack(fill="both", expand=True)
    card = widgets.RoundCard(host, padx=theme.CARD_PAD_X, pady=theme.CARD_PAD_Y)
    card.pack(fill="x")
    viewable_root.update_idletasks()

    called = []
    orig = card.inner.update_idletasks

    def _track():
        called.append(1)
        return orig()

    card.inner.update_idletasks = _track
    try:
        assert card.winfo_viewable(), "测试前提：卡片所在窗口应可见"
        card.refresh()
    finally:
        card.inner.update_idletasks = orig

    assert called, "可见窗口下 refresh 未调用 update_idletasks —— 卡片将量到旧高度（回归）"
