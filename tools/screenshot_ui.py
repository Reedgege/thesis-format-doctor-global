# -*- coding: utf-8 -*-
"""把真机界面截成 PNG（英文 / 中文各一张），用于人工检查排版效果。

用法（本机，需带 tkinter 的 Python + Pillow）：
  set PYTHONPATH=<repo>\\.venv\\Lib\\site-packages
  <sys python> tools/screenshot_ui.py [输出目录]

默认输出到 D:\\AgentSpace\\_ui_preview\\。
"""
from __future__ import annotations

import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


def main() -> int:
    # Windows 上必须先声明 DPI 感知，否则 winfo_width/rootx（逻辑像素）与
    # ImageGrab 的物理像素对不上，截出来的图会偏移 + 只抓到左上角一块。
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)      # per-monitor
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

    import tkinter as tk
    from PIL import ImageGrab

    from src.ui.app import App

    outdir = sys.argv[1] if len(sys.argv) > 1 else r"D:\AgentSpace\_ui_preview"
    os.makedirs(outdir, exist_ok=True)
    os.environ["TFD_SETTINGS_FILE"] = os.path.join(
        tempfile.mkdtemp(prefix="tfdg-shot-"), "settings.json")

    root = tk.Tk()
    ui = App(root)                       # 由 App 按屏幕算尺寸，别在这里覆盖
    root.geometry("+80+40")
    root.deiconify()
    root.lift()
    root.update()
    root.update_idletasks()

    def shot(name):
        for _ in range(6):               # 多刷几次，等布局稳定
            root.update()
        x, y = root.winfo_rootx(), root.winfo_rooty()
        w, h = root.winfo_width(), root.winfo_height()
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
        path = os.path.join(outdir, name)
        img.save(path)
        print("[ok] %s  (%dx%d)" % (path, img.width, img.height))

    shot("ui_en.png")
    ui._set_lang("zh")
    root.geometry("+80+40")
    shot("ui_zh.png")

    # 再补两张弹窗（英文界面下的"激活"和"关于"）
    ui._set_lang("en")
    root.update()
    for name, opener in (("activation", ui._activate_dialog),
                         ("about", ui._about_dialog)):
        ui._close_popups()
        root.update()
        opener()
        win = ui._popups[-1]
        # 必须等窗口真正映射出来，否则 winfo_rootx/width 读到的是旧值（截出来会偏）
        try:
            win.wait_visibility()
        except Exception:
            pass
        for _ in range(8):
            root.update()
        win.update_idletasks()
        x, y = win.winfo_rootx(), win.winfo_rooty()
        w, h = win.winfo_width(), win.winfo_height()
        path = os.path.join(outdir, "dialog_%s_en.png" % name)
        ImageGrab.grab(bbox=(x, y, x + w, y + h)).save(path)
        print("[ok] dialog_%s_en.png  %s" % (name, (w, h)))
    ui._close_popups()
    root.update()

    root.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
