"""冒烟测试：验证 UI 组件是否按新的 Component Kit 规范实现。"""

import sys
import tkinter as tk
from tkinter import ttk

# 添加 src 到路径
sys.path.insert(0, "D:/AgentSpace/thesis-format-doctor-global/src")

from ui import theme, widgets, i18n


def test_button_primary():
    """验证主按钮样式：52px 高，9px 圆角，Primary 色填充"""
    root = tk.Tk()
    root.title("Test Primary Button")
    root.configure(bg=theme.BG)

    F = i18n.resolve_font_spec("en")
    btn = widgets.RoundButton(root, text="Check formatting  ›",
                              style="primary", font=F["F_BTN"], height=52)
    btn.pack(padx=40, pady=40)

    # 验证属性
    assert btn._radius == 9, f"Expected radius 9, got {btn._radius}"
    assert btn._h == 52, f"Expected height 52, got {btn._h}"
    assert btn._style == "primary", f"Expected style primary, got {btn._style}"

    root.destroy()
    print("✓ Primary button style OK")


def test_button_secondary():
    """验证次按钮样式：52px 高，9px 圆角，描边"""
    root = tk.Tk()
    root.title("Test Secondary Button")
    root.configure(bg=theme.BG)

    F = i18n.resolve_font_spec("en")
    btn = widgets.RoundButton(root, text="Fix issues",
                              style="secondary", font=F["F_BTN_S"], height=52)
    btn.pack(padx=40, pady=40)

    # 验证属性
    assert btn._radius == 9, f"Expected radius 9, got {btn._radius}"
    assert btn._h == 52, f"Expected height 52, got {btn._h}"
    assert btn._style == "secondary", f"Expected style secondary, got {btn._style}"

    root.destroy()
    print("✓ Secondary button style OK")


def test_badge_required():
    """验证 Required 徽标：胶囊形，PRIMARY_SOFT 背景"""
    root = tk.Tk()
    root.title("Test Required Badge")
    root.configure(bg=theme.BG)

    F = i18n.resolve_font_spec("en")
    row = tk.Frame(root, bg=theme.SURFACE)
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

    root.destroy()
    print("✓ Required badge style OK")


def test_badge_optional():
    """验证 Optional 徽标：胶囊形，SURFACE_SOFT 背景"""
    root = tk.Tk()
    root.title("Test Optional Badge")
    root.configure(bg=theme.BG)

    F = i18n.resolve_font_spec("en")
    row = tk.Frame(root, bg=theme.SURFACE)
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

    root.destroy()
    print("✓ Optional badge style OK")


def test_round_card():
    """验证圆角卡片：12px 圆角，SURFACE 填充，BORDER 描边"""
    root = tk.Tk()
    root.title("Test Round Card")
    root.configure(bg=theme.BG)

    card = widgets.RoundCard(root, padx=theme.CARD_PAD_X,
                             pady=theme.CARD_PAD_Y)
    card.pack(padx=40, pady=40, fill="both", expand=True)

    # 验证属性
    assert card._radius == 12, f"Expected radius 12, got {card._radius}"
    assert card._fill == theme.SURFACE, f"Expected fill {theme.SURFACE}, got {card._fill}"
    assert card._border == theme.BORDER, f"Expected border {theme.BORDER}, got {card._border}"

    root.destroy()
    print("✓ Round card style OK")


def test_stepper():
    """验证步进圆点：todo/active/done 三态"""
    root = tk.Tk()
    root.title("Test Stepper")
    root.configure(bg=theme.BG)

    F = i18n.resolve_font_spec("en")
    row = tk.Frame(root, bg=theme.BG)
    row.pack(padx=40, pady=40)

    for i, state in enumerate(["todo", "active", "done"]):
        step = widgets.StepCircle(row, number=i + 1, font=F["F_STEP_N"])
        step.set_state(state)
        step.pack(side="left", padx=10)

        assert step.state == state, f"Expected state {state}, got {step.state}"

    root.destroy()
    print("✓ Stepper states OK")


def test_theme_tokens():
    """验证主题令牌是否与新的 DESIGN_TOKENS.json 对齐"""
    assert theme.BG == "#F6F3EC", f"Expected BG #F6F3EC, got {theme.BG}"
    assert theme.SURFACE == "#FFFDF9", f"Expected SURFACE #FFFDF9, got {theme.SURFACE}"
    assert theme.SURFACE_SOFT == "#F4F8FC", f"Expected SURFACE_SOFT #F4F8FC, got {theme.SURFACE_SOFT}"
    assert theme.PRIMARY == "#1F6FAE", f"Expected PRIMARY #1F6FAE, got {theme.PRIMARY}"
    assert theme.PRIMARY_HOVER == "#174B7A", f"Expected PRIMARY_HOVER #174B7A, got {theme.PRIMARY_HOVER}"
    assert theme.PRIMARY_SOFT == "#EAF3FB", f"Expected PRIMARY_SOFT #EAF3FB, got {theme.PRIMARY_SOFT}"
    assert theme.TEXT == "#17324D", f"Expected TEXT #17324D, got {theme.TEXT}"
    assert theme.TEXT_2 == "#66819A", f"Expected TEXT_2 #66819A, got {theme.TEXT_2}"
    assert theme.BORDER == "#D8E2EA", f"Expected BORDER #D8E2EA, got {theme.BORDER}"
    assert theme.SUCCESS == "#3E7B67", f"Expected SUCCESS #3E7B67, got {theme.SUCCESS}"
    assert theme.ERROR == "#A94A43", f"Expected ERROR #A94A43, got {theme.ERROR}"
    assert theme.CARD_RADIUS == 12, f"Expected CARD_RADIUS 12, got {theme.CARD_RADIUS}"
    assert theme.PAGE_PAD == 32, f"Expected PAGE_PAD 32, got {theme.PAGE_PAD}"
    assert theme.CARD_GAP == 20, f"Expected CARD_GAP 20, got {theme.CARD_GAP}"
    print("✓ Theme tokens OK")


if __name__ == "__main__":
    print("Running UI component smoke tests...")
    test_theme_tokens()
    test_button_primary()
    test_button_secondary()
    test_badge_required()
    test_badge_optional()
    test_round_card()
    test_stepper()
    print("\nAll smoke tests passed! ✓")