# -*- coding: utf-8 -*-
"""三平台滚轮回归（2026-11-17 老板实测 Linux 操作界面/弹窗滚轮完全不动）。

两层验证：
1. 事件归一化：Windows delta=±120、macOS delta=±1、X11 Button-4/5 都要给出正确方向；
2. 真 Tk 事件：用 ``event_generate`` 把三种事件打到 Canvas 上，滑块位置必须真的动、
   且方向正确（不是只测一个纯函数，绑定漏了照样能发现）。

用法：python tests/test_wheel_cross_platform.py（需要带 tkinter 的系统 Python）
"""
from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import tkinter as tk                                      # noqa: E402

from src.ui.wheel import bind_wheel, wheel_pixels         # noqa: E402

FAILED = []


def ck(label, cond, extra=""):
    print(("  PASS " if cond else "  FAIL ") + label + ("  | " + extra if extra else ""))
    if not cond:
        FAILED.append(label)


class E:
    """最小事件替身。"""

    def __init__(self, delta=0, num=None):
        self.delta = delta
        self.num = num


print("=== 1. 事件归一化 ===")
ck("Windows 上滚 delta=+120 → -120px", wheel_pixels(E(delta=120)) == -120)
ck("Windows 下滚 delta=-120 → +120px", wheel_pixels(E(delta=-120)) == 120)
ck("macOS 上滚 delta=+1 → -120px（旧代码只滚 1px）",
   wheel_pixels(E(delta=1)) == -120)
ck("macOS 下滚 delta=-1 → +120px", wheel_pixels(E(delta=-1)) == 120)
ck("X11 Button-4（无 delta）→ -120px", wheel_pixels(E(num=4)) == -120)
ck("X11 Button-5（无 delta）→ +120px", wheel_pixels(E(num=5)) == 120)
ck("触控板 delta=-40 → +40px（不放大成一格）", wheel_pixels(E(delta=-40)) == 40)
ck("非滚轮事件 → 0", wheel_pixels(E()) == 0)

print("\n=== 2. 真 Tk 事件打到 Canvas（三种事件绑定是否齐全）===")
root = tk.Tk()
# 🔴 不能 withdraw()：窗口未映射时 Canvas 高度为 1，yview_scroll 算出来的 fraction
# 恒为 0，合成事件看起来「没滚」—— 那是测试夹具的问题，不是绑定问题。
root.geometry("300x220+0+0")
root.update()
try:
    cv = tk.Canvas(root, width=200, height=100, highlightthickness=0,
                   yscrollincrement=1)
    cv.pack()
    cv.configure(scrollregion=(0, 0, 200, 1000))
    cv.update_idletasks()

    def handler(e):
        px = wheel_pixels(e)
        if px:
            cv.yview_scroll(px, "units")
        return "break"

    n = bind_wheel(cv, handler)
    ck("一次绑齐 MouseWheel / Button-4 / Button-5", n == 3)
    for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        ck("%s 已绑定" % seq, bool(cv.bind(seq)))

    def fire(seq, **kw):
        root.update()
        cv.event_generate(seq, x=50, y=50, when="now", **kw)
        root.update()
        return cv.yview()[0]

    cv.yview_moveto(0.5)
    base = cv.yview()[0]
    down = fire("<Button-5>")
    ck("X11 Button-5 真的向下滚", down > base, "%.4f → %.4f" % (base, down))

    cv.yview_moveto(0.5)
    up = fire("<Button-4>")
    ck("X11 Button-4 真的向上滚", up < base, "%.4f → %.4f" % (base, up))

    cv.yview_moveto(0.5)
    wdown = fire("<MouseWheel>", delta=-120)
    ck("MouseWheel delta=-120 向下滚", wdown > base, "%.4f → %.4f" % (base, wdown))

    cv.yview_moveto(0.5)
    mup = fire("<MouseWheel>", delta=1)
    ck("macOS 风格 delta=+1 也向上滚 120px", mup < base, "%.4f → %.4f" % (base, mup))
finally:
    root.destroy()

print("\n" + "=" * 66)
if FAILED:
    print("RESULT: FAIL (%d)" % len(FAILED))
    for f in FAILED:
        print("  - " + f)
    sys.exit(1)
print("RESULT: ALL PASS")
