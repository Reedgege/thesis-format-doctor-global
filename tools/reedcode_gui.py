# -*- coding: utf-8 -*-
"""
卖家发码工具（兜底用）· GUI 版 · Thesis Format Doctor Global（海外版）
======================================================================
仅在「中台跑路 / 宕机 / 客户无法联网」时给客户重激活使用。

用法：双击 reedcode-global.exe（打包后），或 python reedcode_gui.py
1) 让客户在软件激活页点「复制机器码」，把机器码发给你
2) 粘贴到下方输入框，点「生成离线备用码」
3) 把生成的码发给客户，他在激活页选「离线激活」粘贴即可
4) 自动记入 reedcode_log.txt 台账

安全：离线码用对称签名（密钥与 src/license/license.py 共用，明文常量）。
"""
import os
import sys
import time
import hmac
import hashlib

import tkinter as tk
from tkinter import messagebox

if getattr(sys, "frozen", False):
    # Nuitka / PyInstaller 冻结：exe 所在目录即 HERE，台账随 exe 移动、持久保存，
    # 不会被临时目录清理掉。
    HERE = os.path.dirname(os.path.abspath(sys.executable))
else:
    HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# ---------------------------------------------------------------------------
# 离线备用码算法（与 src/license/license.py 共用同一密钥与签名方式，已内联，
# 打包后不依赖任何外部文件）。与国内版密钥不同，防互用。
# ---------------------------------------------------------------------------
_OFFLINE_KEY = b"tfdglobal|kami|offline|2026|sign|v1"


def _offline_sign(payload):
    return hmac.new(_OFFLINE_KEY, payload.encode("utf-8"), hashlib.sha256).hexdigest()[:24]


def generate_offline_code(machine_code):
    """machine_code + 时间戳 + 签名 → 离线码（与主程序 verify_offline_code 兼容）。"""
    ts = int(time.time())
    payload = "%s|%d" % (machine_code, ts)
    return payload + "|" + _offline_sign(payload)

# 学术文艺风配色（与海外版主程序呼应；深绿主调）
BG = "#1F4E46"        # 深绿底
INNER = "#F5F0E6"     # 宣纸内层面板
FG = "#2b2b2b"        # 墨色
MUTE = "#6b5d4f"      # 淡墨
ACCENT = "#C8A96A"    # 金（主按钮）
LINE = "#d8cdb8"      # 细线
TXT_ON_DARK = "#F5F0E6"


def build_ui(root):
    root.title("Thesis Format Doctor Global · 离线发码工具")
    root.geometry("560x470")
    root.resizable(False, False)
    root.configure(bg=BG)
    try:
        root.iconbitmap(os.path.join(HERE, "assets", "icon.ico"))
    except Exception:
        pass

    tk.Label(root, text="离 线 发 码 工 具", bg=BG, fg=ACCENT,
             font=("Times New Roman", 18, "bold")).pack(pady=(14, 0))
    tk.Label(root, text="（兜底用 · 仅当中台不可用 / 客户无法联网时给客户重激活）",
             bg=BG, fg=TXT_ON_DARK, font=("Microsoft YaHei", 9)).pack(pady=(2, 8))

    tk.Label(root,
        text="① 让客户在软件激活页点「复制机器码」，把机器码发给你\n"
             "② 粘贴到下方，点「生成离线备用码」\n"
             "③ 把生成的码发给客户，他在激活页选「离线激活」粘贴即可",
        bg=BG, fg=TXT_ON_DARK, font=("Microsoft YaHei", 10), justify="left"
    ).pack(padx=24, anchor="w")

    frm = tk.Frame(root, bg=BG)
    frm.pack(padx=24, pady=(10, 4), fill="x")
    tk.Label(frm, text="客户的机器码：", bg=BG, fg=TXT_ON_DARK,
             font=("Microsoft YaHei", 10)).pack(anchor="w")
    entry = tk.Entry(frm, font=("Consolas", 11), bd=1, relief="solid",
                     bg=INNER, fg=FG, insertbackground=ACCENT)
    entry.pack(fill="x", pady=(4, 0), ipady=4)

    btnf = tk.Frame(root, bg=BG)
    btnf.pack(padx=24, pady=(6, 4), fill="x")
    gen_btn = tk.Button(btnf, text="生成离线备用码", bg=ACCENT, fg="#1F4E46",
                        font=("Microsoft YaHei", 11, "bold"), relief="flat",
                        activebackground="#b8965a", padx=14, pady=6,
                        cursor="hand2")
    gen_btn.pack(side="left")

    tk.Label(root, text="生成的离线备用码（已自动复制，可直接发给客户）：",
             bg=BG, fg=MUTE, font=("Microsoft YaHei", 9)
             ).pack(padx=24, pady=(12, 2), anchor="w")
    res_text = tk.Text(root, height=4, font=("Consolas", 11),
                       bg=INNER, fg=FG, relief="solid", bd=1,
                       wrap="word", state="disabled")
    res_text.pack(padx=24, fill="x")

    status = tk.Label(root, text="", bg=BG, fg=TXT_ON_DARK, font=("Microsoft YaHei", 9))
    status.pack(padx=24, pady=(8, 0), anchor="w")

    def generate():
        target = entry.get().strip()
        if not target:
            messagebox.showwarning("提示", "请先粘贴客户的机器码")
            return
        try:
            code = generate_offline_code(target)
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            return
        res_text.configure(state="normal")
        res_text.delete("1.0", "end")
        res_text.insert("end", code)
        res_text.configure(state="disabled")
        try:
            root.clipboard_clear()
            root.clipboard_append(code)
        except Exception:
            pass
        log_path = os.path.join(HERE, "reedcode_log.txt")
        logged = False
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write("%s | 客户机器码 %s | 离线码 %s\n" % (
                    time.strftime("%Y-%m-%d %H:%M"), target, code))
            logged = True
        except Exception:
            pass
        if logged:
            status.configure(
                text="✓ 已生成并复制到剪贴板，已记入台账 reedcode_log.txt（与 exe 同目录）")
        else:
            status.configure(
                text="✓ 已生成并复制到剪贴板，但台账写入失败（可能磁盘只读），请手动保存上方离线码")

    gen_btn.configure(command=generate)


def main():
    root = tk.Tk()
    build_ui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
