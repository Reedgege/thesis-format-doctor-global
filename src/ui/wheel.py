# -*- coding: utf-8 -*-
"""鼠标滚轮事件归一化（Windows / macOS / X11 Linux）。

🔴 2026-11-17 老板实测：Linux 版海外软件的操作界面和弹窗，滚轮完全不动。
根因不是「系统不兼容」，是**只绑了 ``<MouseWheel>``**：X11 下 Tk 发的是
``<Button-4>``（向上）/ ``<Button-5>``（向下），Windows / macOS 才发
``<MouseWheel>``。国内两版的操作界面/弹窗所以没全军覆没，是因为它们一直
三种事件都绑了。

macOS 还有第二种体感问题：MouseWheel 的 ``delta`` 通常是 ±1（Windows 是 ±120），
主滚动区原先把 delta 直接当像素用 → 一格只滚 1px，和「滚不动」没区别。

这里把三种事件统一成**向下为正的像素数**，一格 = 120px（与 Windows 现有手感一致）。
Canvas 的 ``yscrollincrement`` 必须显式为 1 才能让 1 unit == 1px，调用方都已设置。
"""
from __future__ import annotations

NOTCH_PX = 120


def wheel_pixels(event) -> int:
    """把滚轮事件转成向下为正的像素数；不是滚轮事件或无效值返回 0。"""
    num = getattr(event, "num", None)
    if num == 4:
        return -NOTCH_PX          # X11 向上
    if num == 5:
        return NOTCH_PX           # X11 向下

    raw = getattr(event, "delta", 0) or 0
    try:
        d = float(raw)
    except (TypeError, ValueError):
        return 0
    if not d:
        return 0
    # macOS 鼠标一格给 ±1、触控板给 ±30/±60；X11 少数环境也只给 ±1
    if abs(d) < 10:
        d = NOTCH_PX if d > 0 else -NOTCH_PX
    return -int(round(d))         # delta>0 = 向上滚 → 负像素


def bind_wheel(widget, handler) -> int:
    """三种滚轮事件一次绑齐（add="+" 与已有绑定共存）。返回绑定数。"""
    for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        widget.bind(seq, handler, add="+")
    return 3
