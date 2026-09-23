"""冒烟测试：验证 UI 组件是否按新的 Component Kit 规范实现。

跑法：需要**带 tkinter** 的 Python。托管 venv 没有 tkinter 时整个文件自动 skip；
本机用系统 Python + venv 的 site-packages 可全量跑：

    PYTHONPATH=<venv>/Lib/site-packages python -m pytest tests/test_component_kit.py

注意两条仓库约定（CI 上曾经因此变红，别再改回去）：

1. **不许写死本机绝对路径**。CI runner 上没有 ``D:/AgentSpace/...``，
   写死会让 ``from ui import ...`` 直接 ImportError。一律用 ``__file__`` 求根。
2. **同进程只建一个 ``tk.Tk()``**。第二次 ``tk.Tk()`` 在本机会因 Tcl 库加载失败
   抛 TclError，整套跑时随机有 1 个用例失败。复用 session 级 root（见
   ``test_ui_redesign.py`` 的同名夹具约定）。
"""

from __future__ import annotations

import importlib.util
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(ROOT, "src")
for _p in (_SRC, ROOT):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from ui import theme, widgets, i18n  # noqa: E402


@pytest.fixture(scope="session")
def _tk_session():
    """整个会话共用一个 Tk root（见模块 docstring 约定 2）。"""
    if importlib.util.find_spec("tkinter") is None:
        pytest.skip("本机 Python 无 tkinter（GUI 冒烟需带 tkinter 的解释器）")
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
def host(_tk_session):
    """在每个用例里挂一个独立容器，用完销毁，避免会话 root 上组件越堆越多。"""
    import tkinter as tk

    frame = tk.Frame(_tk_session, bg=theme.BG)
    frame.pack(fill="both", expand=True)
    yield frame
    try:
        frame.destroy()
    except Exception:
        pass


def test_button_primary(host):
    """验证主按钮样式：52px 高，9px 圆角，Primary 色填充"""
    F = i18n.resolve_font_spec("en")
    btn = widgets.RoundButton(host, text="Check formatting  ›",
                              style="primary", font=F["F_BTN"], height=52)
    btn.pack(padx=40, pady=40)

    # 验证属性
    assert btn._radius == 9, f"Expected radius 9, got {btn._radius}"
    assert btn._h == 52, f"Expected height 52, got {btn._h}"
    assert btn._style == "primary", f"Expected style primary, got {btn._style}"
    print("✓ Primary button style OK")


def test_button_secondary(host):
    """验证次按钮样式：52px 高，9px 圆角，描边"""
    F = i18n.resolve_font_spec("en")
    btn = widgets.RoundButton(host, text="Fix issues",
                              style="secondary", font=F["F_BTN_S"], height=52)
    btn.pack(padx=40, pady=40)

    # 验证属性
    assert btn._radius == 9, f"Expected radius 9, got {btn._radius}"
    assert btn._h == 52, f"Expected height 52, got {btn._h}"
    assert btn._style == "secondary", f"Expected style secondary, got {btn._style}"
    print("✓ Secondary button style OK")


def test_badge_required(host):
    """验证 Required 徽标：胶囊形，PRIMARY_SOFT 背景"""
    import tkinter as tk

    F = i18n.resolve_font_spec("en")
    row = tk.Frame(host, bg=theme.SURFACE)
    row.pack(padx=40, pady=40)

    tk.Label(row, text="Citation style", bg=theme.SURFACE, fg=theme.TEXT,
             font=F["F_LABEL"]).pack(side="left")

    badge = tk.Label(row,
                     text="  Required  ",
                     bg=theme.PRIMARY_SOFT, fg=theme.PRIMARY,
                     font=F["F_HELP"], padx=6, pady=1)
    badge.pack(side="left", padx=(7, 0))

    # 验证颜色
    assert badge.cget("bg") == theme.PRIMARY_SOFT, f"Expected bg {theme.PRIMARY_SOFT}, got {badge.cget('bg')}"
    assert badge.cget("fg") == theme.PRIMARY, f"Expected fg {theme.PRIMARY}, got {badge.cget('fg')}"
    print("✓ Required badge style OK")


def test_badge_optional(host):
    """验证 Optional 徽标：胶囊形，SURFACE_SOFT 背景"""
    import tkinter as tk

    F = i18n.resolve_font_spec("en")
    row = tk.Frame(host, bg=theme.SURFACE)
    row.pack(padx=40, pady=40)

    tk.Label(row, text="University template", bg=theme.SURFACE, fg=theme.TEXT,
             font=F["F_LABEL"]).pack(side="left")

    badge = tk.Label(row,
                     text="  Optional  ",
                     bg=theme.SURFACE_SOFT, fg=theme.TEXT_2,
                     font=F["F_HELP"], padx=6, pady=1)
    badge.pack(side="left", padx=(7, 0))

    # 验证颜色
    assert badge.cget("bg") == theme.SURFACE_SOFT, f"Expected bg {theme.SURFACE_SOFT}, got {badge.cget('bg')}"
    assert badge.cget("fg") == theme.TEXT_2, f"Expected fg {theme.TEXT_2}, got {badge.cget('fg')}"
    print("✓ Optional badge style OK")


def test_round_card(host):
    """验证圆角卡片：12px 圆角，SURFACE 填充，BORDER 描边"""
    card = widgets.RoundCard(host, padx=theme.CARD_PAD_X,
                             pady=theme.CARD_PAD_Y)
    card.pack(padx=40, pady=40, fill="both", expand=True)

    # 验证属性
    assert card._radius == 12, f"Expected radius 12, got {card._radius}"
    assert card._fill == theme.SURFACE, f"Expected fill {theme.SURFACE}, got {card._fill}"
    assert card._border == theme.BORDER, f"Expected border {theme.BORDER}, got {card._border}"
    print("✓ Round card style OK")


def test_progress(host):
    """验证极简进度条：set_fraction / set_step 驱动进度，frac 可被断言。"""
    import tkinter as tk

    bar = widgets.ProgressBar(host, steps=3, bg=theme.BG)
    bar.pack(padx=40, pady=40)

    # 直接设比例
    assert bar.set_fraction(0.0) == 0.0, f"Expected frac 0.0, got {bar.frac}"
    assert bar.set_fraction(0.5) == 0.5, f"Expected frac 0.5, got {bar.frac}"
    # 越界被夹到 [0, 1]
    assert bar.set_fraction(2.0) == 1.0, f"Expected clamp to 1.0, got {bar.frac}"
    assert bar.set_fraction(-1.0) == 0.0, f"Expected clamp to 0.0, got {bar.frac}"
    # 按步设置：3 步，index=2 → 2/3
    bar.set_step(2)
    assert abs(bar.frac - 2 / 3) < 1e-6, f"Expected frac 2/3, got {bar.frac}"
    # index=3（全完成）→ 满格
    bar.set_step(3)
    assert bar.frac == 1.0, f"Expected frac 1.0, got {bar.frac}"
    print("✓ Progress bar OK")


def test_theme_tokens():
    """验证主题令牌是否与 UI_DESIGN_SPEC 对齐（theme.py 是颜色令牌的唯一来源）。

    v2.3.0 的 UI 重做把界面从「老式 Windows 工具」改成「premium 学术生产力软件」：
    冷调近白纸底 + 青蓝主色。本测试锁定这些对外视觉令牌，防止回退到旧的
    米色 / 浅蓝配色（#F6F3EC / #1F6FAE）。
    """
    assert theme.BG == "#F7F9FB", f"Expected BG #F7F9FB (cool-white), got {theme.BG}"
    assert theme.SURFACE == "#FFFFFF", f"Expected SURFACE #FFFFFF, got {theme.SURFACE}"
    assert theme.SURFACE_SOFT == "#F4F6F8", f"Expected SURFACE_SOFT #F4F6F8 (neutral cool-gray), got {theme.SURFACE_SOFT}"
    assert theme.PRIMARY == "#0E7490", f"Expected PRIMARY #0E7490 (sole saturated teal-cyan), got {theme.PRIMARY}"
    assert theme.PRIMARY_HOVER == "#0A5C6B", f"Expected PRIMARY_HOVER #0A5C6B, got {theme.PRIMARY_HOVER}"
    assert theme.PRIMARY_SOFT == "#DCF1F5", f"Expected PRIMARY_SOFT #DCF1F5, got {theme.PRIMARY_SOFT}"
    assert theme.TEXT == "#17324D", f"Expected TEXT #17324D, got {theme.TEXT}"
    assert theme.TEXT_2 == "#64748B", f"Expected TEXT_2 #64748B (neutral slate, no blue), got {theme.TEXT_2}"
    assert theme.BORDER == "#E2E8F0", f"Expected BORDER #E2E8F0 (neutral cool-gray), got {theme.BORDER}"
    assert theme.SUCCESS == "#3E7B67", f"Expected SUCCESS #3E7B67, got {theme.SUCCESS}"
    assert theme.ERROR == "#A94A43", f"Expected ERROR #A94A43, got {theme.ERROR}"
    assert theme.CARD_RADIUS == 12, f"Expected CARD_RADIUS 12, got {theme.CARD_RADIUS}"
    assert theme.PAGE_PAD == 32, f"Expected PAGE_PAD 32, got {theme.PAGE_PAD}"
    assert theme.CARD_GAP == 20, f"Expected CARD_GAP 20, got {theme.CARD_GAP}"
    print("✓ Theme tokens OK")


if __name__ == "__main__":
    import tkinter as tk

    print("Running UI component smoke tests...")
    test_theme_tokens()
    root = tk.Tk()
    for fn in (test_button_primary, test_button_secondary, test_badge_required,
               test_badge_optional, test_round_card, test_progress):
        frame = tk.Frame(root, bg=theme.BG)
        frame.pack(fill="both", expand=True)
        fn(frame)
        frame.destroy()
    root.destroy()
    print("\nAll smoke tests passed! ✓")
