# -*- coding: utf-8 -*-
"""布局健全性检查：抓「控件被压扁 / 文字被截断 / 越出容器」。

为什么不用截图判断：Windows 的 DPI 缩放会让 ``winfo_rootx``（逻辑像素）与
ImageGrab（物理像素）对不上，截图容易骗人。这里改看两个**程序化**信号：

  1. **压缩**：``winfo_reqwidth() > winfo_width()`` —— 控件请求的宽度比实际分配到的
     还大，说明它被挤了，文字多半被截断（"Export AI Template" 被裁成
     "Export AI Templa" 就是这么来的）。
  2. **越界**：控件右/下边缘超出所属容器 —— 说明父容器没算对尺寸。

覆盖主窗口 + 四个弹窗，两种语言各跑一遍（英文文案普遍更长，最容易溢出）。

用法（本机）：
  set PYTHONPATH=<repo>\\.venv\\Lib\\site-packages
  <sys python> tools\\check_layout.py
退出码：0 = 通过；1 = 发现问题。
"""
from __future__ import annotations

import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

# 这些控件按设计会被拉伸/压缩（不参与"被截断"判定）
_SKIP_CLASSES = ("Text", "Canvas", "Frame", "Toplevel", "TScrollbar",
                 "Scrollbar", "Progressbar", "TCombobox")
TOL = 4          # 允许几像素的取整/字体度量误差


def _in_scrolled(widget) -> bool:
    """控件是否位于可滚动容器（Canvas）里。

    滚动区里的控件"超出窗口底边"是**正常**的（内容本来就比窗口长，靠滚动看），
    不参与越界判定，否则帮助/关于弹窗会刷一屏误报。
    """
    node = widget
    for _ in range(40):
        try:
            node = node.master
        except Exception:
            return False
        if node is None:
            return False
        try:
            if node.winfo_class() == "Canvas":
                return True
        except Exception:
            pass
    return False


def _cls(w) -> str:
    try:
        return w.winfo_class()
    except Exception:
        return "?"


def _walk(widget):
    try:
        children = widget.winfo_children()
    except Exception:
        return
    for c in children:
        yield c
        yield from _walk(c)


def check_window(win, label: str, problems: list):
    try:
        win.update_idletasks()
    except Exception:
        pass
    ww, wh = win.winfo_width(), win.winfo_height()
    if ww <= 1 or wh <= 1:
        problems.append("%s: 窗口尺寸异常 (%dx%d)" % (label, ww, wh))
        return

    for child in _walk(win):
        cls = _cls(child)
        if cls in _SKIP_CLASSES:
            continue
        try:
            x, y = child.winfo_x(), child.winfo_y()
            cw, ch = child.winfo_width(), child.winfo_height()
            rw, rh = child.winfo_reqwidth(), child.winfo_reqheight()
        except Exception:
            continue
        if cw <= 1 or ch <= 1:
            continue        # 未映射/被 pack_forget 的控件
        if rw > cw + TOL:
            problems.append("%s: %s 被压缩（请求宽 %d > 实际 %d）: %r"
                            % (label, cls, rw, cw, _text_of(child)))
        # 容器尺寸算错导致的越界（滚动区内的控件跳过：它们本来就比窗口长）
        if _in_scrolled(child):
            continue
        if x + cw > ww + TOL or y + ch > wh + TOL:
            problems.append("%s: %s 越出窗口（右/下 %d,%d > %d,%d）: %r"
                            % (label, cls, x + cw, y + ch, ww, wh, _text_of(child)))


def _text_of(w) -> str:
    try:
        t = w.cget("text")
    except Exception:
        return ""
    s = str(t)
    return s[:40]


def main() -> int:
    try:
        import tkinter as tk
    except Exception as exc:
        print("[skip] 当前 Python 无 tkinter：%s" % exc, file=sys.stderr)
        return 0

    os.environ["TFD_SETTINGS_FILE"] = os.path.join(
        tempfile.mkdtemp(prefix="tfdg-layout-"), "settings.json")

    from src.ui import app as ui_app

    problems: list = []
    root = tk.Tk()
    root.withdraw()
    try:
        ui = ui_app.App(root)
        for lang in ("en", "zh"):
            if ui.lang != lang:
                ui._set_lang(lang)
            root.update_idletasks()
            root.update()
            check_window(root, "主窗口[%s]" % lang, problems)

            for name, opener in (("关于", ui._about_dialog),
                                 ("帮助", ui._help_dialog),
                                 ("激活", ui._activate_dialog),
                                 ("问卷", ui._open_questionnaire)):
                ui._close_popups()
                root.update()
                opener()
                win = ui._popups[-1]
                for _ in range(4):
                    root.update()
                check_window(win, "%s[%s]" % (name, lang), problems)
                ui._close_popups()
                root.update()
    finally:
        try:
            root.destroy()
        except Exception:
            pass

    if problems:
        print("[fail] 发现 %d 处布局问题：" % len(problems), file=sys.stderr)
        for p in problems:
            print("   - " + p, file=sys.stderr)
        return 1
    print("[ok] 布局检查通过：主窗口 + 四个弹窗（中英各一遍）无控件被压缩或越界")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
