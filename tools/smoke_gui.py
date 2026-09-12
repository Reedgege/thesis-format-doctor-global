# -*- coding: utf-8 -*-
"""真机 GUI 冒烟：用**带 tkinter 的 Python**把界面真跑一遍（主窗口 + 四个弹窗 + 双语）。

为什么需要它：CI 打包的是 Nuitka 编译后的 exe，本地 venv 没有 tkinter，
所以「窗口能不能真的起来、弹窗能不能真的打开」必须单独验一次。构造 ``App``
会执行完整的 ``_build()``（全部控件、变量、事件绑定），能在秒级内捕捉到
「某个控件构造就抛异常」这类启动崩溃 —— v2.0.1 的"双击打不开"就是这么暴露的。

覆盖点（每一项失败都会打印 [fail] 并返回非 0）：
  1. 主窗口能构建（控件数 > 0），窗口图标能挂上；
  2. **默认语言是英文**，且英文界面里除语言切换按钮外不出现中文；
  3. 切到中文再切回英文，界面能整体重建且不崩；
  4. 关于 / 帮助 / 激活 / 问卷 四个弹窗都能打开（内容非空）并正常关闭。

用法（本机）：
  set PYTHONPATH=<repo>\\.venv\\Lib\\site-packages
  C:\\Users\\yunwu\\AppData\\Local\\Programs\\Python\\Python313\\python.exe tools\\smoke_gui.py

退出码：0 = 通过；1 = 失败（错误打到 stderr）。
"""
from __future__ import annotations

import os
import sys
import tempfile
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

FAILURES: list = []


def _fail(msg: str):
    FAILURES.append(msg)
    print("[fail] " + msg, file=sys.stderr)


def _count_widgets(widget) -> int:
    n = 1
    try:
        for child in widget.winfo_children():
            n += _count_widgets(child)
    except Exception:
        pass
    return n


def _labels(root) -> list:
    import tkinter as tk

    out = []
    for w in _walk(root):
        if isinstance(w, tk.Label):
            try:
                out.append(str(w.cget("text")))
            except Exception:
                pass
    return out


def _walk(widget):
    try:
        children = widget.winfo_children()
    except Exception:
        return
    for c in children:
        yield c
        yield from _walk(c)


def _has_cjk(texts) -> list:
    return [t for t in texts if any("\u4e00" <= c <= "\u9fff" for c in t)]


def main() -> int:
    try:
        import tkinter as tk
    except Exception as exc:
        print("[skip] 当前 Python 无 tkinter：%s" % exc, file=sys.stderr)
        return 0

    # 隔离用户偏好，别把冒烟的语言选择写进真实设置文件
    os.environ.setdefault("TFD_SETTINGS_FILE",
                          os.path.join(tempfile.mkdtemp(prefix="tfdg-smoke-"),
                                       "settings.json"))

    try:
        from src.ui import app as ui_app
        from src.ui import iconpath, i18n
    except Exception:
        traceback.print_exc()
        return 1

    root = None
    try:
        root = tk.Tk()
        root.withdraw()                 # 冒烟不闪窗
        ui = ui_app.App(root)           # 完整构建界面
        root.update_idletasks()         # 真跑一遍几何计算/布局

        widgets = _count_widgets(root)
        print("[ok] Tk %s | 控件数=%d" % (tk.TkVersion, widgets))
        if widgets < 20:
            _fail("主界面控件数异常偏少（%d）—— 界面可能没建起来" % widgets)

        # --- 图标 ---
        found = iconpath.find_icon()
        icon = getattr(ui, "_window_icon", None)
        print("[ok] 图标路径=%s" % found)
        print("[ok] 图标已挂载=%s" % (icon is not None))
        if found and icon is None:
            _fail("找到了图标文件但没能挂到窗口上")

        # --- 默认语言 ---
        print("[ok] 默认语言=%s（应为 en）" % ui.lang)
        if ui.lang != "en":
            _fail("默认语言不是英文：%s" % ui.lang)
        texts = _labels(root)
        if "Thesis Format Doctor" not in texts:
            _fail("英文标题缺失：%s" % texts[:6])
        allowed = {ui.tr("lang_button")}
        stray = _has_cjk([t for t in texts if t not in allowed])
        if stray:
            _fail("英文界面里出现中文：%s" % stray[:6])

        # --- 语言切换 ---
        ui._set_lang("zh")
        root.update_idletasks()
        zh_texts = _labels(root)
        if "你的文档" not in zh_texts or "检查与修正" not in zh_texts:
            _fail("切到中文后界面没跟着变：%s" % zh_texts[:6])
        if i18n.load_lang() != "zh":
            _fail("语言偏好没有落盘")
        print("[ok] 切换到中文：控件数=%d" % _count_widgets(root))

        ui._set_lang("en")
        root.update_idletasks()
        if "Your Documents" not in _labels(root):
            _fail("切回英文失败")
        print("[ok] 切回英文：控件数=%d" % _count_widgets(root))

        # --- 弹窗 ---
        dialogs = [("关于", ui._about_dialog), ("帮助", ui._help_dialog),
                   ("激活", ui._activate_dialog), ("问卷", ui._open_questionnaire)]
        for name, opener in dialogs:
            before = len(ui._popups)
            try:
                opener()
                root.update_idletasks()
            except Exception as exc:                     # noqa: BLE001
                traceback.print_exc()
                _fail("弹窗「%s」打开失败：%s" % (name, exc))
                continue
            opened = len(ui._popups) > before
            win = ui._popups[-1] if ui._popups else None
            inner = _count_widgets(win) if win is not None else 0
            print("[ok] 弹窗「%s」已打开：控件数=%d" % (name, inner))
            if not opened or inner < 5:
                _fail("弹窗「%s」没有真正建起来（控件数=%d）" % (name, inner))
            ui._close_popups()
            root.update_idletasks()

        if FAILURES:
            print("\n[result] 冒烟失败 %d 项" % len(FAILURES), file=sys.stderr)
            return 1
        print("\n[result] 冒烟通过：主界面 + 双语切换 + 四个弹窗全部正常")
        return 0
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        try:
            if root is not None:
                root.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
