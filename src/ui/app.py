"""Thesis Format Doctor Global —— tkinter 桌面 GUI（离线，论文不出本机）。

界面骨架沿用国内版（tfd-desktop）那套「宣纸 + 墨 + 黛蓝 + 朱砂」的学术文艺风：
顶栏细链接 ─ 居中大标题 ─ 朱砂细线 ─ 双栏卡片（左「文件选择」/ 右「处理步骤」）
─ 报告区 ─ 页脚 ─ 状态栏。

与国内版的三点不同（老板 2026-09-11 定）：
1. **默认英文**，右上角一键切中文；语言偏好落盘，下次启动记得住。
2. 英文字体走 Georgia（标题，学术衬线）+ Segoe UI（正文），不套中文字体，
   英文排版才立得住；中文界面仍用楷体 + 雅黑。
3. **品牌口径只留官网与邮箱**，不出现微信 / 公众号 / 小程序（海外用户不用）。

功能面：
  选规范 →（可选）学校模板 / LaTeX / 问卷 / AI JSON → 选论文
  → 运行体检（免费不限次）→ 一键修正（首次免费，之后需激活码）。
GUI 与 CLI 共用同一套 engine，保证逻辑唯一、离线安全。
"""

from __future__ import annotations

import os
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, scrolledtext, ttk

from ..engine import checker, report as report_mod
from ..engine.docx_reader import DocxReadError
from ..engine.fixer import compute_changes, fix_docx, same_file
from ..engine.questionnaire import (
    FIELD_RANGES, QUESTIONNAIRE_SCHEMA, TargetProfile, build_target,
)
from ..engine.specs import SPEC_ORDER
from ..license import license as lic
from .. import versioninfo
from . import i18n, iconpath
from .fonts import Fonts

# ---------------------------------------------------------------------------
# 配色：宣纸 / 墨 / 黛蓝 / 朱砂 —— 与国内版同一套语言（学术范 + 文艺感）
# ---------------------------------------------------------------------------
PAPER = "#f6f2ea"      # 宣纸底
PANEL = "#fdfbf6"      # 面板（微亮的纸）
CARD = "#ffffff"       # 卡片内白块
INK = "#222222"        # 主文字
BODY = "#444444"       # 正文灰
MUTED = "#666666"      # 次要文字
LINE = "#e0d9c8"       # 细线
ACCENT = "#46586f"     # 黛蓝（章节条 / 强调）
CINNABAR = "#9e4233"   # 朱砂（主按钮 / 印记）
CINNABAR_D = "#8a382b"
OKC = "#5f7d5c"        # 完成（墨绿）
ERRC = "#a0402f"       # 出错（朱红）
REPORT_BG = "#fffdf8"  # 报告纸

_BASE_W = 1180         # 设计基准宽度（与 fonts.BASE_WIDTH 一致）

_HERE = os.path.dirname(os.path.abspath(__file__))
_TEMPLATE_PATH = os.path.join(_HERE, "..", "data", "ai_questionnaire_template.json")

_FIELD_RANGE = FIELD_RANGES


def _app_version() -> str:
    """版本号（转发给 :mod:`src.versioninfo`，CLI 与 GUI 共用同一套查找逻辑）。"""
    return versioninfo.app_version()


def _icon_candidates() -> list:
    """窗口图标的候选路径（转发给 ``iconpath``，便于复用/测试）。"""
    return iconpath.icon_candidates()


def _btn_display_width(text: str, pad: int = 2) -> int:
    """按钮文案的「显示宽度」：全角字符按 2、半角按 1 累加，再加左右余量。

    ttk.Button 的 width 单位是「英文字符宽」，中日韩全角字符渲染宽度约为英文的 2 倍，
    直接拿 len(text) 当宽度会把含中文的文案截断（国内版 v1.3.98 踩过这个坑）。
    辅助按钮现已改为 grid 等宽布局、无需固定宽度，这里保留工具函数供后续按文案
    定宽的场景使用（例如新增竖排按钮）。
    """
    w = 0
    for ch in text:
        o = ord(ch)
        if (0x1100 <= o <= 0x115F or 0x2E80 <= o <= 0xA4CF
                or 0xAC00 <= o <= 0xD7A3 or 0xF900 <= o <= 0xFAFF
                or 0xFE30 <= o <= 0xFE6F or 0xFF00 <= o <= 0xFF60
                or 0xFFE0 <= o <= 0xFFE6 or 0x20000 <= o <= 0x3FFFD):
            w += 2
        else:
            w += 1
    return max(8, w + pad)


def _scroll_frame(parent, bg: str = PAPER):
    """可滚动容器：返回 inner frame，内容 pack/grid 进去即可，超出自动出滚动条。"""
    canvas = tk.Canvas(parent, bg=bg, highlightthickness=0)
    vsb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas, bg=bg)
    win = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_inner(_e=None):
        try:
            canvas.configure(scrollregion=canvas.bbox("all"))
        except Exception:
            pass

    def _on_canvas(e):
        try:
            canvas.itemconfig(win, width=e.width)
        except Exception:
            pass

    inner.bind("<Configure>", _on_inner)
    canvas.bind("<Configure>", _on_canvas)
    canvas.configure(yscrollcommand=vsb.set)
    canvas.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    def _wheel(e):
        try:
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        except Exception:
            pass

    canvas.bind_all("<MouseWheel>", _wheel)
    return inner


def _auto_wrap(label: tk.Label, container, padx: int):
    """让 label 的 wraplength 跟随容器宽度（弹窗拉宽时文字重排，不出现横向滚动）。"""

    def _apply(_e=None):
        try:
            w = container.winfo_width() - 2 * padx
            if w > 80:
                label.configure(wraplength=w)
        except Exception:
            pass

    container.bind("<Configure>", _apply, add="+")
    label.after(30, _apply)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.lang = i18n.load_lang()
        self.F = Fonts(root)
        self.F.rebuild(self.lang)

        self.version = _app_version()
        self._last_scale = None
        self._popups: list = []
        self._bar_state = "idle"
        self._report_text = ""
        self._report_is_placeholder = True
        self._step_index = 0
        self._busy = False

        # 输入源
        self.spec_key = tk.StringVar(value="APA")
        self.school_template = tk.StringVar(value="")
        self.latex_template = tk.StringVar(value="")
        self.ai_json = tk.StringVar(value="")
        self.docx_path = tk.StringVar(value="")
        self.questionnaire_data: dict = {}

        _sw = root.winfo_screenwidth()
        _sh = root.winfo_screenheight()
        root.title("Thesis Format Doctor Global")
        # 高度按屏幕 92% 给足：左栏六行材料 + 底部报告区要一屏放得下，否则报告区会被压扁。
        root.geometry("%dx%d" % (min(_BASE_W, int(_sw * 0.86)), min(1000, int(_sh * 0.94))))
        root.minsize(1040, 700)
        self._apply_window_icon()
        root.bind("<Configure>", self._on_resize)

        self._build()
        self.root.after(120, lambda: self._on_resize())

    # ------------------------------------------------------------ 基础工具
    def tr(self, key: str, *args):
        return i18n.t(self.lang, key, *args)

    def _sep(self) -> str:
        return "｜" if self.lang == "zh" else "|"

    def _apply_window_icon(self):
        """设置标题栏/任务栏图标。任何异常都吞掉——图标不该影响启动。"""
        path = iconpath.find_icon()
        if not path:
            return
        try:
            # PhotoImage 必须挂在实例上：只作为临时对象传入时一旦被 GC，
            # 部分 Tk 版本会把图标还原成默认（经典坑）。
            self._window_icon = tk.PhotoImage(file=path)
            self.root.iconphoto(True, self._window_icon)
            return
        except Exception:
            pass
        try:
            self.root.iconbitmap(default=path)
        except Exception:
            pass

    def _on_resize(self, _evt=None):
        try:
            w = self.root.winfo_width()
        except Exception:
            return
        factor = self.F.scale_for_width(w)
        if self._last_scale is not None and abs(factor - self._last_scale) < 0.02:
            return
        self._last_scale = factor
        self.F.apply_scale(factor)

    def _center(self, win: tk.Toplevel, w: int, h: int):
        try:
            x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - w) // 2)
            y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - h) // 4)
            win.geometry("%dx%d+%d+%d" % (w, h, x, y))
        except Exception:
            win.geometry("%dx%d" % (w, h))

    def _track(self, win: tk.Toplevel):
        self._popups.append(win)
        win.protocol("WM_DELETE_WINDOW", lambda: self._close_one(win))

    def _close_one(self, win):
        try:
            self._popups.remove(win)
        except ValueError:
            pass
        try:
            win.destroy()
        except Exception:
            pass

    def _close_popups(self):
        for w in list(self._popups):
            self._close_one(w)

    def _link(self, parent, text, cmd, fg=ACCENT):
        lbl = tk.Label(parent, text=text, bg=PAPER, fg=fg,
                       font=self.F["F_SUBTITLE"], cursor="hand2")
        lbl.bind("<Button-1>", lambda e: cmd())
        lbl.bind("<Enter>", lambda e: lbl.config(fg="#33465c"))
        lbl.bind("<Leave>", lambda e: lbl.config(fg=fg))
        return lbl

    # ------------------------------------------------------------ 样式
    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        F = self.F
        style.configure("TButton", font=F["F_BODY"], padding=(12, 6),
                        background="#eae3d3", foreground=INK)
        style.map("TButton", background=[("active", "#ded6c2"), ("disabled", "#f0ece1")])
        style.configure("Action.TButton", font=F["F_BTN"], padding=(12, 10),
                        background="#eae3d3", foreground=INK)
        style.map("Action.TButton",
                  background=[("active", "#ded6c2"), ("disabled", "#f0ece1")])
        style.configure("Primary.TButton", font=F["F_BTN"], padding=(12, 10),
                        background=CINNABAR, foreground="white")
        style.map("Primary.TButton",
                  background=[("active", CINNABAR_D), ("disabled", "#c8a79f")],
                  foreground=[("disabled", "#f5e9e6")])
        style.configure("Ghost.TButton", font=F["F_SMALL"], padding=(8, 5),
                        background=PANEL, foreground=ACCENT)
        style.map("Ghost.TButton",
                  background=[("active", "#efe9db"), ("disabled", "#f4f0e6")])
        style.configure("TCombobox", font=F["F_BODY"])

    # ------------------------------------------------------------ 整体布局
    def _build(self):
        self._build_style()
        self.root.configure(bg=PAPER)
        F = self.F
        sep = self._sep()

        # 顶栏：语言 ｜ 关于 ｜ 帮助（从右往左 pack）
        topbar = tk.Frame(self.root, bg=PAPER)
        topbar.pack(fill="x", padx=20, pady=(10, 0))
        self._link(topbar, self.tr("link_help"), self._help_dialog).pack(side="right")
        tk.Label(topbar, text=sep, bg=PAPER, fg="#c9c0ae",
                 font=F["F_SUBTITLE"]).pack(side="right", padx=(0, 8))
        self._link(topbar, self.tr("link_about"), self._about_dialog).pack(
            side="right", padx=(0, 8))
        tk.Label(topbar, text=sep, bg=PAPER, fg="#c9c0ae",
                 font=F["F_SUBTITLE"]).pack(side="right", padx=(0, 8))
        self._link(topbar, self.tr("lang_button"),
                   lambda: self._set_lang(i18n.other_lang(self.lang))).pack(
            side="right", padx=(0, 8))

        # 标题区
        header = tk.Frame(self.root, bg=PAPER)
        header.pack(fill="x", padx=20, pady=(10, 6))
        tk.Label(header, text=self.tr("app_title"), bg=PAPER, fg=INK,
                 font=F["F_TITLE"]).pack(anchor="center")
        tk.Label(header, text=self.tr("app_subtitle"), bg=PAPER, fg="#8b8378",
                 font=F["F_SUB"]).pack(anchor="center", pady=(5, 0))
        tk.Frame(self.root, bg=CINNABAR, height=2).pack(fill="x", padx=20)

        # 主体：上排两栏（文件选择 | 处理步骤），下排报告区
        main = tk.Frame(self.root, bg=PAPER)
        main.pack(fill="both", expand=True, padx=20, pady=(10, 0))
        main.columnconfigure(0, weight=1, uniform="half", minsize=470)
        main.columnconfigure(1, weight=1, uniform="half", minsize=470)
        # 报告区：给足最小高度（minsize），否则会被左栏六行卡片挤成一条缝
        main.rowconfigure(1, weight=1, minsize=150)
        left = tk.Frame(main, bg=PAPER)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        right = tk.Frame(main, bg=PAPER)
        right.grid(row=0, column=1, sticky="nsew")
        self._build_left(left)
        self._build_right(right)

        report_card = tk.Frame(main, bg=PANEL, highlightthickness=1,
                               highlightbackground=LINE)
        report_card.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(14, 0))
        r_hdr = tk.Frame(report_card, bg=PANEL)
        r_hdr.pack(fill="x", padx=14, pady=(8, 4))
        tk.Frame(r_hdr, bg=ACCENT, width=4, height=15).pack(side="left", padx=(0, 7))
        tk.Label(r_hdr, text=self.tr("report_title"), bg=PANEL, fg=INK,
                 font=F["F_CARD_HDR"]).pack(side="left")
        self.report = scrolledtext.ScrolledText(
            report_card, wrap="word", height=7, bg=REPORT_BG, fg=INK,
            relief="flat", bd=0, padx=12, pady=8, font=F["F_BODY"],
            insertbackground=ACCENT, highlightthickness=0)
        self.report.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        self.report.configure(state="disabled")

        # 页脚（只留官网 + 邮箱，不出现微信 / 公众号 / 小程序）
        footer = tk.Frame(self.root, bg=PAPER)
        footer.pack(side="bottom", fill="x", pady=(4, 10))
        tk.Label(footer, text=self.tr("footer_line1"), bg=PAPER, fg=MUTED,
                 font=F["F_FOOT"]).pack()
        f2 = tk.Frame(footer, bg=PAPER)
        f2.pack(pady=(3, 0))
        tk.Label(f2, text=self.tr("footer_line2"), bg=PAPER, fg=MUTED,
                 font=F["F_FOOT"]).pack(side="left")
        tk.Label(f2, text="  ·  ", bg=PAPER, fg="#c9c0ae",
                 font=F["F_FOOT"]).pack(side="left")
        site = tk.Label(f2, text=self.tr("footer_site"), bg=PAPER, fg=ACCENT,
                        font=F["F_FOOT"], cursor="hand2")
        site.bind("<Button-1>", lambda e: webbrowser.open(i18n.BRAND_SITE_URL))
        site.pack(side="left")

        # 状态栏
        tk.Frame(self.root, bg=LINE, height=1).pack(fill="x", side="bottom", padx=20)
        statusbar = tk.Frame(self.root, bg=PANEL)
        statusbar.pack(side="bottom", fill="x", padx=20, pady=(4, 4))
        self.bar_dot = tk.Label(statusbar, text="●", bg=PANEL, fg=OKC, font=F["F_FOOT"])
        self.bar_dot.pack(side="left", padx=(0, 6))
        tk.Label(statusbar, text=self.tr("bar_left", self.version), bg=PANEL,
                 fg=MUTED, font=F["F_FOOT"]).pack(side="left")
        self.bar_right = tk.Label(statusbar, text="", bg=PANEL, fg=MUTED,
                                  font=F["F_FOOT"])
        self.bar_right.pack(side="right")

        # 重建后恢复：报告内容、各行状态、时间线、状态栏
        if self._report_is_placeholder or not self._report_text:
            self._set_report(self.tr("report_placeholder"), placeholder=True)
        else:
            self._set_report(self._report_text)
        self._refresh_rows()
        self._set_bar(self._bar_state)
        self._set_status(self.tr("status_idle"), "idle" if self._bar_state == "idle"
                         else self._bar_state)

    # ------------------------------------------------------------ 左栏
    def _build_left(self, parent):
        F = self.F
        card = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        card.pack(fill="both", expand=True)

        hdr = tk.Frame(card, bg=PANEL)
        hdr.pack(fill="x", padx=14, pady=(10, 4))
        tk.Frame(hdr, bg=ACCENT, width=4, height=15).pack(side="left", padx=(0, 7))
        tk.Label(hdr, text=self.tr("files_title"), bg=PANEL, fg=INK,
                 font=F["F_CARD_HDR"]).pack(side="left")
        tk.Label(hdr, text=self.tr("badge_offline"), bg="#e8f0f6", fg=ACCENT,
                 font=F["F_FOOT"], padx=8, pady=2).pack(side="right")

        # ① 引用规范（下拉）
        self._spec_cb = self._spec_row(card)
        # ② 待检论文（必选）
        self._paper_box, self._paper_name = self._file_row(
            card, self.tr("ic_paper"), self.tr("row_paper_title"),
            self.tr("mark_required"), CINNABAR, self.tr("row_paper_desc"),
            [(self.tr("btn_choose"), self._pick_docx)])
        # ③ 学校模板（可选）
        self._tpl_box, self._tpl_name = self._file_row(
            card, self.tr("ic_template"), self.tr("row_template_title"),
            self.tr("mark_optional"), MUTED, self.tr("row_template_desc"),
            [(self.tr("btn_choose"), self._pick_template)])
        # ④ LaTeX 模板（可选）
        self._latex_box, self._latex_name = self._file_row(
            card, self.tr("ic_latex"), self.tr("row_latex_title"),
            self.tr("mark_optional"), MUTED, self.tr("row_latex_desc"),
            [(self.tr("btn_choose"), self._pick_latex)])
        # ⑤ 格式问卷（可选）、⑥ AI 填表 JSON（可选）——两行分开，避免两个按钮挤爆一行
        self._q_box, self._q_name = self._file_row(
            card, self.tr("ic_questionnaire"), self.tr("row_questionnaire_title"),
            self.tr("mark_optional"), MUTED, self.tr("row_questionnaire_desc"),
            [(self.tr("btn_questionnaire"), self._open_questionnaire)])
        self._ai_box, self._ai_name = self._file_row(
            card, self.tr("ic_ai"), self.tr("row_ai_title"),
            self.tr("mark_optional"), MUTED, self.tr("row_ai_desc"),
            [(self.tr("btn_ai_json"), self._pick_ai_json)])

        tk.Label(card, text=self.tr("chapter_one"), bg=PANEL, fg="#b8b0a0",
                 font=F["F_FOOT"]).pack(side="bottom", pady=(0, 6))

    def _spec_row(self, card):
        """引用规范行：行内是下拉框（选规范），下方一行小字提示 Other 的用法。"""
        F = self.F
        box = tk.Frame(card, bg=CARD, highlightthickness=1, highlightbackground="#e3dccb")
        box.pack(fill="x", padx=14, pady=4)
        row = tk.Frame(box, bg=CARD)
        row.pack(fill="x", padx=10, pady=4)

        tk.Label(row, text=self.tr("ic_spec"), bg=CARD, fg=CINNABAR, font=F["F_ICON"],
                 padx=7, pady=2, highlightthickness=1,
                 highlightbackground=CINNABAR).pack(side="left", padx=(0, 8))

        txt = tk.Frame(row, bg=CARD)
        txt.pack(side="left", fill="x", expand=True)
        tl = tk.Frame(txt, bg=CARD)
        tl.pack(fill="x")
        tk.Label(tl, text=self.tr("row_spec_title"), bg=CARD, fg=INK,
                 font=F["F_SUBTITLE"]).pack(side="left")
        tk.Label(tl, text=" " + self.tr("mark_required"), bg=CARD, fg=CINNABAR,
                 font=F["F_FOOT"]).pack(side="left")
        tk.Label(txt, text=self.tr("row_spec_desc"), bg=CARD, fg=MUTED,
                 font=F["F_FOOT"], anchor="w").pack(anchor="w")

        cb = ttk.Combobox(row, textvariable=self.spec_key, values=SPEC_ORDER,
                          state="readonly", width=14, font=F["F_BODY"])
        cb.pack(side="right", padx=(8, 0))

        tk.Label(box, text=self.tr("chip_style_hint"), bg=CARD, fg="#8a5a1a",
                 font=F["F_FOOT"], anchor="w").pack(fill="x", padx=12, pady=(0, 8))
        return cb

    def _file_row(self, parent, icon, title, mark, mark_color, desc, buttons):
        """一行「文件选择」：朱砂印记 + 标题 + 可选标记 + 状态/描述 + 右侧按钮。"""
        F = self.F
        box = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground="#e3dccb")
        box.pack(fill="x", padx=14, pady=4)
        row = tk.Frame(box, bg=CARD)
        row.pack(fill="x", padx=10, pady=4)

        tk.Label(row, text=icon, bg=CARD, fg=CINNABAR, font=F["F_ICON"],
                 padx=7, pady=2, highlightthickness=1,
                 highlightbackground=CINNABAR).pack(side="left", padx=(0, 8))

        txt = tk.Frame(row, bg=CARD)
        txt.pack(side="left", fill="x", expand=True)
        tl = tk.Frame(txt, bg=CARD)
        tl.pack(fill="x")
        tk.Label(tl, text=title, bg=CARD, fg=INK,
                 font=F["F_SUBTITLE"]).pack(side="left")
        tk.Label(tl, text=" " + mark, bg=CARD, fg=mark_color,
                 font=F["F_FOOT"]).pack(side="left")
        name_lbl = tk.Label(txt, text=desc, bg=CARD, fg=MUTED, font=F["F_FOOT"],
                            anchor="w", justify="left", wraplength=340)
        name_lbl.pack(anchor="w", fill="x")

        # 按钮从右往左排（列表中第一个按钮最靠右）
        for text, cmd in reversed(buttons):
            ttk.Button(row, text=text, style="Ghost.TButton", command=cmd).pack(
                side="right", padx=(6, 0))
        return box, name_lbl

    # ------------------------------------------------------------ 右栏
    def _build_right(self, parent):
        F = self.F
        card = tk.Frame(parent, bg=PANEL, highlightthickness=1, highlightbackground=LINE)
        card.pack(fill="both", expand=True)

        hdr = tk.Frame(card, bg=PANEL)
        hdr.pack(fill="x", padx=14, pady=(12, 4))
        tk.Frame(hdr, bg=ACCENT, width=4, height=15).pack(side="left", padx=(0, 7))
        tk.Label(hdr, text=self.tr("steps_title"), bg=PANEL, fg=INK,
                 font=F["F_CARD_HDR"]).pack(side="left")
        self._step_counter = tk.Label(hdr, text="", bg=PANEL, fg=MUTED, font=F["F_FOOT"])
        self._step_counter.pack(side="right")

        self._build_timeline(card)

        tk.Frame(card, bg=LINE, height=1).pack(fill="x", padx=14, pady=(10, 0))

        status_row = tk.Frame(card, bg=PANEL)
        status_row.pack(fill="x")
        self.status_dot = tk.Label(status_row, text="●", bg=PANEL, fg=MUTED,
                                   font=F["F_BODY"])
        self.status_dot.pack(side="left", padx=(14, 6), pady=(8, 4))
        self.status_lbl = tk.Label(status_row, text=self.tr("status_idle"), bg=PANEL,
                                   fg=INK, font=F["F_STAT"], anchor="w",
                                   justify="left", wraplength=430)
        self.status_lbl.pack(side="left", fill="x", expand=True, pady=(8, 4))

        btn_row = tk.Frame(card, bg=PANEL)
        btn_row.pack(fill="x", padx=14, pady=(6, 4))
        self._check_btn = ttk.Button(btn_row, text=self.tr("btn_check"),
                                     style="Primary.TButton", command=self._run)
        self._check_btn.pack(fill="x")

        btn_row2 = tk.Frame(card, bg=PANEL)
        btn_row2.pack(fill="x", padx=14, pady=(0, 4))
        self._fix_btn = ttk.Button(btn_row2, text=self.tr("btn_fix"),
                                   style="Action.TButton", command=self._run_fix)
        self._fix_btn.pack(fill="x")

        aux = tk.Frame(card, bg=PANEL)
        aux.pack(fill="x", padx=14, pady=(0, 12))
        aux.columnconfigure(0, weight=1, uniform="aux")
        aux.columnconfigure(1, weight=1, uniform="aux")
        # 两行放三个辅助按钮：三个挤一行时 "Export AI Template" 会被 cell 宽度截断
        pairs = [((self.tr("btn_save_report"), self._save_report), 0, 0),
                 ((self.tr("btn_export_ai"), self._export_template), 0, 1),
                 ((self.tr("btn_activate"), self._activate_dialog), 1, 0)]
        for (text, cmd), r, c in pairs:
            b = ttk.Button(aux, text=text, style="Ghost.TButton", command=cmd)
            span = 2 if (r == 1) else 1
            b.grid(row=r, column=c, columnspan=span, sticky="we",
                   padx=(0, 6) if (c == 0 and span == 1) else 0, pady=(0, 4))

        tk.Label(card, text=self.tr("chapter_two"), bg=PANEL, fg="#b8b0a0",
                 font=F["F_FOOT"]).pack(side="bottom", pady=(0, 6))

    def _build_timeline(self, parent):
        F = self.F
        tl = tk.Frame(parent, bg=PANEL)
        tl.pack(fill="x", padx=14, pady=(4, 2))
        self._step_defs = [
            (self.tr("step1_title"), self.tr("step1_desc")),
            (self.tr("step2_title"), self.tr("step2_desc")),
            (self.tr("step3_title"), self.tr("step3_desc")),
        ]
        self._step_circle: list = []
        self._step_title: list = []
        self._step_line: list = []
        n = len(self._step_defs)
        for i, (label, desc) in enumerate(self._step_defs):
            row = tk.Frame(tl, bg=PANEL)
            row.pack(fill="x", pady=2)
            col = tk.Frame(row, bg=PANEL)
            col.pack(side="left", padx=(0, 10))
            circ = tk.Label(col, text=str(i + 1), bg=CARD, fg="#8b8378",
                            font=F["F_STAT"], width=2, height=1, relief="flat",
                            highlightthickness=1, highlightbackground="#c9c1ae")
            circ.pack()
            line = None
            if i < n - 1:
                line = tk.Frame(col, width=2, height=22, bg="#e3dccb")
                line.pack()
            self._step_circle.append(circ)
            self._step_line.append(line)
            txt = tk.Frame(row, bg=PANEL)
            txt.pack(side="left", fill="x", expand=True)
            title = tk.Label(txt, text=label, bg=PANEL, fg=MUTED, font=F["F_SUBTITLE"])
            title.pack(anchor="w")
            tk.Label(txt, text=desc, bg=PANEL, fg=MUTED,
                     font=F["F_SMALL"]).pack(anchor="w")
            self._step_title.append(title)

    def _refresh_wizard(self):
        """按 _step_index 重绘时间线（已完成 ✔ / 当前黛蓝 / 未来灰）。"""
        n = len(self._step_defs)
        for i in range(n):
            circ = self._step_circle[i]
            title = self._step_title[i]
            line = self._step_line[i]
            if i < self._step_index:
                circ.config(bg=CARD, fg=OKC, highlightbackground=OKC, text="✔")
                title.config(fg=INK)
                if line:
                    line.config(bg=OKC)
            elif i == self._step_index:
                circ.config(bg=CARD, fg=ACCENT, highlightbackground=ACCENT,
                            text=str(i + 1))
                title.config(fg=INK)
                if line:
                    line.config(bg="#e3dccb")
            else:
                circ.config(bg=CARD, fg="#8b8378", highlightbackground="#c9c1ae",
                            text=str(i + 1))
                title.config(fg=MUTED)
                if line:
                    line.config(bg="#e3dccb")
        self._step_counter.config(
            text=self.tr("step_counter", min(self._step_index + 1, n), n))

    # ------------------------------------------------------------ 状态刷新
    def _refresh_rows(self):
        """刷新各输入源的状态显示 + 时间线进度。"""
        tr = self.tr
        colon = "：" if self.lang == "zh" else ": "

        def _set(lbl, value, set_key):
            if value:
                name = os.path.basename(value)
                if len(name) > 46:
                    name = name[:22] + " … " + name[-20:]
                lbl.config(text="● " + tr(set_key) + colon + name, fg=OKC)
                return True
            return False
        if not _set(self._paper_name, self.docx_path.get(), "st_paper_set"):
            self._paper_name.config(text=tr("st_paper_none"), fg=MUTED)
        if not _set(self._tpl_name, self.school_template.get(), "st_tpl_set"):
            self._tpl_name.config(text=tr("st_tpl_none"), fg=MUTED)
        if not _set(self._latex_name, self.latex_template.get(), "st_latex_set"):
            self._latex_name.config(text=tr("st_latex_none"), fg=MUTED)

        # 问卷 / AI JSON 各自一行（问卷的状态来自内存数据，不是文件路径）
        if self.questionnaire_data:
            self._q_name.config(text="● " + tr("st_q_set"), fg=OKC)
        else:
            self._q_name.config(text=tr("st_q_none"), fg=MUTED)
        if not _set(self._ai_name, self.ai_json.get(), "st_ai_set"):
            self._ai_name.config(text=tr("st_ai_none"), fg=MUTED)

        # 时间线：选了论文即可进入体检
        if self._step_index < 1 and self.docx_path.get():
            self._step_index = 1
        self._refresh_wizard()

    def _set_status(self, text: str, kind: str = "idle"):
        colors = {"idle": MUTED, "run": ACCENT, "ok": OKC, "err": ERRC}
        try:
            self.status_dot.config(fg=colors.get(kind, MUTED))
            self.status_lbl.config(text=text)
        except Exception:
            pass

    def _set_bar(self, state: str, key: str = None):
        """状态栏（左下灯 + 右下文案）；state: idle / run / ok / err。"""
        self._bar_state = state
        if key is None:
            key = {"idle": "bar_idle", "run": "bar_checking", "ok": "bar_check_done",
                   "err": "bar_error"}.get(state, "bar_idle")
        colors = {"idle": OKC, "run": ACCENT, "ok": OKC, "err": ERRC}
        try:
            self.bar_dot.config(fg=colors.get(state, OKC))
            self.bar_right.config(text=self.tr(key))
        except Exception:
            pass

    def _set_report(self, text: str, placeholder: bool = False):
        """写入报告区。

        ``placeholder=True`` 表示这只是"还没跑体检"的引导语 —— 切换语言重建界面时
        要按新语言重新取词，而不是把旧语言的引导语原样搬过去（否则中文界面里会残留
        一段英文提示）。
        """
        self._report_text = text
        self._report_is_placeholder = placeholder
        try:
            self.report.configure(state="normal")
            self.report.delete("1.0", tk.END)
            self.report.insert(tk.END, text)
            self.report.configure(state="disabled")
        except Exception:
            pass

    def _set_busy(self, busy: bool):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for btn in (self._check_btn, self._fix_btn):
            try:
                btn.configure(state=state)
            except Exception:
                pass
        try:
            self.root.update_idletasks()
        except Exception:
            pass

    # ------------------------------------------------------------ 语言切换
    def _set_lang(self, lang: str):
        """切换语言：重建整套界面（低频操作，重建比逐条改文案可靠得多）。"""
        if lang == self.lang or lang not in i18n.LANGS:
            return
        i18n.save_lang(lang)
        self.lang = lang
        self.F.rebuild(lang)
        self._close_popups()
        for w in self.root.winfo_children():
            w.destroy()
        self._last_scale = None
        self._build()
        self._on_resize()

    # ------------------------------------------------------------ 选择文件
    def _pick_docx(self):
        p = filedialog.askopenfilename(title=self.tr("row_paper_title"),
                                       filetypes=[("Word", "*.docx"), ("All files", "*.*")])
        if p:
            self.docx_path.set(p)
            self._refresh_rows()

    def _pick_template(self):
        p = filedialog.askopenfilename(title=self.tr("row_template_title"),
                                       filetypes=[("Word", "*.docx"), ("All files", "*.*")])
        if p:
            self.school_template.set(p)
            self._refresh_rows()

    def _pick_latex(self):
        p = filedialog.askopenfilename(
            title=self.tr("row_latex_title"),
            filetypes=[("LaTeX", "*.tex *.cls *.sty"), ("All files", "*.*")])
        if p:
            self.latex_template.set(p)
            self._refresh_rows()

    def _pick_ai_json(self):
        p = filedialog.askopenfilename(title=self.tr("btn_ai_json"),
                                       filetypes=[("JSON", "*.json"), ("All files", "*.*")])
        if p:
            self.ai_json.set(p)
            self._refresh_rows()

    # ------------------------------------------------------------ 问卷
    def _open_questionnaire(self):
        tr = self.tr
        win = tk.Toplevel(self.root)
        win.title(tr("q_title"))
        win.configure(bg=PAPER)
        self._center(win, 880, 700)
        win.minsize(760, 560)
        self._track(win)

        tk.Label(win, text=tr("q_title"), bg=PAPER, fg=INK,
                 font=self.F["F_DIALOG_TITLE"]).pack(pady=(16, 2))
        tk.Label(win, text=tr("q_hint"), bg=PAPER, fg=MUTED,
                 font=self.F["F_SMALL"]).pack(pady=(0, 8))
        tk.Frame(win, bg=LINE, height=1).pack(fill="x", padx=24)

        inner = _scroll_frame(win, bg=PAPER)
        form = tk.Frame(inner, bg=PAPER)
        form.pack(fill="both", expand=True, padx=24, pady=10)

        vars_: dict = {}
        row = 0
        for field in QUESTIONNAIRE_SCHEMA:
            key = field["key"]
            tk.Label(form, text=i18n.questionnaire_label(self.lang, field), bg=PAPER,
                     fg=INK, font=self.F["F_SMALL"], anchor="w").grid(
                row=row, column=0, sticky="w", padx=(0, 10), pady=4)
            ftype = field.get("type")
            if ftype == "bool":
                # 布尔项用"是/否"下拉，比裸复选框更清楚（默认值照 schema）
                v = tk.BooleanVar(value=bool(field.get("default", False)))
                cb = ttk.Combobox(form, state="readonly", width=18,
                                  font=self.F["F_SMALL"],
                                  values=[tr("bool_yes"), tr("bool_no")])
                cb.set(tr("bool_yes") if v.get() else tr("bool_no"))
                cb.grid(row=row, column=1, sticky="w")
                vars_[key] = ("bool", cb, v)
            elif ftype == "choice":
                v = tk.StringVar(value=str(field.get("default", "")))
                ttk.Combobox(form, textvariable=v, values=field.get("options", []),
                             state="readonly", width=18,
                             font=self.F["F_SMALL"]).grid(row=row, column=1, sticky="w")
                vars_[key] = ("str", v)
            else:
                v = tk.StringVar(value=str(field.get("default", "")))
                ttk.Entry(form, textvariable=v, width=20,
                          font=self.F["F_SMALL"]).grid(row=row, column=1, sticky="w")
                vars_[key] = ("str", v)
            hint = i18n.questionnaire_help(self.lang, field)
            tk.Label(form, text=hint, bg=PAPER, fg="#888888", font=self.F["F_FOOT"],
                     wraplength=330, justify="left", anchor="w").grid(
                row=row, column=2, sticky="w", padx=(12, 0))
            row += 1

        def _save():
            data: dict = {}
            warnings: list = []
            for field in QUESTIONNAIRE_SCHEMA:
                key = field["key"]
                if field.get("type") == "bool":
                    _kind, cb, _v = vars_[key]
                    data[key] = (cb.get() == tr("bool_yes"))
                    continue
                val = vars_[key][1].get().strip()
                if not val:
                    continue
                if field.get("type") in ("float", "int"):
                    try:
                        num = float(val)
                    except ValueError:
                        messagebox.showwarning(
                            tr("q_bad_number_title"),
                            tr("q_bad_number",
                               i18n.questionnaire_label(self.lang, field), val),
                            parent=win)
                        return
                    rng = _FIELD_RANGE.get(key)
                    if rng and not (rng[0] <= num <= rng[1]):
                        warnings.append("%s = %s  (%s–%s)" % (
                            i18n.questionnaire_label(self.lang, field), val,
                            rng[0], rng[1]))
                data[key] = val
            if warnings and not messagebox.askyesno(
                    tr("q_range_title"), tr("q_range_body", "\n".join(warnings)),
                    parent=win):
                return
            self.questionnaire_data = data
            messagebox.showinfo(tr("q_saved_title"), tr("q_saved"), parent=win)
            self._close_one(win)
            self._refresh_rows()

        bar = tk.Frame(win, bg=PAPER)
        bar.pack(fill="x", side="bottom", pady=12)
        ttk.Button(bar, text=tr("q_save"), style="Primary.TButton",
                   command=_save).pack()

    # ------------------------------------------------------------ 关于 / 帮助
    def _about_dialog(self):
        tr = self.tr
        win = tk.Toplevel(self.root)
        win.title(tr("about_title"))
        win.configure(bg=PAPER)
        self._center(win, 580, 660)
        win.resizable(True, True)
        win.minsize(520, 520)
        self._track(win)

        inner = _scroll_frame(win, bg=PAPER)
        padx = 44

        def center_lbl(text, fg, font, pady, nowrap=False):
            opts = dict(bg=PAPER, fg=fg, font=font, justify="center")
            if not nowrap:
                opts["wraplength"] = 420
            lbl = tk.Label(inner, text=text, **opts)
            if not nowrap:
                _auto_wrap(lbl, inner, padx)
            lbl.pack(padx=padx, pady=pady)
            return lbl

        def rule(pady=(0, 14)):
            tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=pady)

        tk.Label(inner, text=tr("app_title"), bg=PAPER, fg=INK,
                 font=self.F["F_TITLE"]).pack(padx=padx, pady=(24, 2))
        tk.Label(inner, text=tr("about_version", self.version), bg=PAPER, fg=MUTED,
                 font=self.F["F_SUBTITLE"]).pack(padx=padx, pady=(0, 12))
        rule()

        center_lbl(tr("about_tagline"), INK, self.F["F_HDR"], (0, 8), nowrap=True)
        center_lbl(tr("about_body"), BODY, self.F["F_BODY"], (0, 14))
        rule()

        center_lbl(tr("about_source_hdr"), ACCENT, self.F["F_HDR"], (0, 6))
        center_lbl(tr("about_source"), BODY, self.F["F_SMALL"], (0, 14))
        rule()

        center_lbl(tr("about_privacy_hdr"), ACCENT, self.F["F_HDR"], (0, 6))
        center_lbl(tr("about_privacy"), BODY, self.F["F_SMALL"], (0, 14))
        rule()

        # 联系方式：只留官网与邮箱（海外版铁律：不出现微信 / 公众号 / 小程序）
        center_lbl(tr("about_contact_hdr"), ACCENT, self.F["F_HDR"], (0, 6))
        center_lbl(i18n.BRAND_EMAIL, INK, self.F["F_BODY"], (0, 2), nowrap=True)
        site = tk.Label(inner, text=tr("footer_site"), bg=PAPER, fg=ACCENT,
                        font=self.F["F_BODY"], cursor="hand2")
        site.bind("<Button-1>", lambda e: webbrowser.open(i18n.BRAND_SITE_URL))
        site.pack(padx=padx, pady=(0, 10))
        center_lbl(tr("about_smartscreen"), MUTED, self.F["F_FOOT"], (0, 14))

        tk.Label(inner, text=tr("about_footer"), bg=PAPER, fg=MUTED,
                 font=self.F["F_FOOT"]).pack(padx=padx, pady=(0, 18))

    def _help_dialog(self):
        tr = self.tr
        win = tk.Toplevel(self.root)
        win.title(tr("help_title"))
        win.configure(bg=PAPER)
        self._center(win, 820, 700)
        win.resizable(True, True)
        win.minsize(680, 520)
        self._track(win)

        inner = _scroll_frame(win, bg=PAPER)
        padx = 40

        def head(text, fg, font, pady):
            lbl = tk.Label(inner, text=text, bg=PAPER, fg=fg, font=font,
                           justify="center", wraplength=660)
            _auto_wrap(lbl, inner, padx)
            lbl.pack(padx=padx, pady=pady)
            return lbl

        def section(text):
            tk.Label(inner, text=text, bg=PAPER, fg=ACCENT, font=self.F["F_HDR"],
                     anchor="w").pack(fill="x", padx=padx, pady=(14, 4))
            tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 4))

        def item(title, body):
            tl = tk.Label(inner, text=title, bg=PAPER, fg=INK,
                          font=self.F["F_SMALL_B"], anchor="w", justify="left",
                          wraplength=660)
            _auto_wrap(tl, inner, padx)
            tl.pack(fill="x", padx=padx, pady=(8, 1))
            bl = tk.Label(inner, text=body, bg=PAPER, fg=BODY,
                          font=self.F["F_SUBTITLE"], anchor="w", justify="left",
                          wraplength=660)
            _auto_wrap(bl, inner, padx)
            bl.pack(fill="x", padx=padx, pady=(0, 4))

        tk.Label(inner, text=tr("help_title"), bg=PAPER, fg=INK,
                 font=self.F["F_TITLE"]).pack(padx=padx, pady=(22, 4))
        head(tr("help_intro"), BODY, self.F["F_BODY"], (0, 12))
        tk.Frame(inner, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 8))

        section(tr("help_sec1"))
        for title, body in tr("help_steps"):
            item(title, body)

        section(tr("help_sec2"))
        for title, body in tr("help_faqs"):
            item(title, body)

        tk.Label(inner, text=tr("about_footer"), bg=PAPER, fg=MUTED,
                 font=self.F["F_FOOT"]).pack(padx=padx, pady=(16, 18))

    # ------------------------------------------------------------ 激活
    def _activate_dialog(self):
        tr = self.tr
        F = self.F
        win = tk.Toplevel(self.root)
        win.title(tr("act_title"))
        win.configure(bg=PAPER)
        self._center(win, 560, 640)
        win.resizable(False, False)
        self._track(win)

        tk.Label(win, text=tr("act_title"), bg=PAPER, fg=INK,
                 font=F["F_DIALOG_TITLE"]).pack(pady=(16, 2))
        tk.Label(win, text=tr("act_subtitle"), bg=PAPER, fg=MUTED,
                 font=F["F_SMALL"], wraplength=440, justify="center").pack(pady=(0, 10))
        tk.Frame(win, bg=LINE, height=1).pack(fill="x", padx=24)

        # ① 在线激活码
        tk.Label(win, text=tr("act_online_label"), bg=PAPER, fg=INK,
                 font=F["F_SMALL_B"], anchor="w").pack(fill="x", padx=24, pady=(12, 2))
        code_var = tk.StringVar()
        ttk.Entry(win, textvariable=code_var, width=44,
                  font=F["F_BODY"]).pack(padx=24, fill="x")
        tk.Label(win, text=tr("act_online_hint"), bg=PAPER, fg=MUTED,
                 font=F["F_FOOT"], anchor="w").pack(fill="x", padx=24, pady=(2, 6))

        def _submit():
            code = code_var.get().strip()
            if not code:
                messagebox.showwarning(tr("act_title"), tr("act_need_code"), parent=win)
                return
            r = lic.activate(code)
            messagebox.showinfo(tr("act_result_title"), r["message"], parent=win)
            if r["ok"]:
                self._close_one(win)

        ttk.Button(win, text=tr("act_online_btn"), style="TButton",
                   command=_submit).pack(anchor="w", padx=24, pady=(0, 8))

        tk.Frame(win, bg=LINE, height=1).pack(fill="x", padx=24, pady=6)

        # ② 本机机器码（发给卖家换离线码）
        tk.Label(win, text=tr("act_machine_label"), bg=PAPER, fg=INK,
                 font=F["F_SMALL_B"], anchor="w").pack(fill="x", padx=24, pady=(6, 2))
        mc_var = tk.StringVar(value=lic._machine_fingerprint())
        ttk.Entry(win, textvariable=mc_var, width=44, state="readonly",
                  font=F["F_MONO"]).pack(padx=24, fill="x")
        tk.Label(win, text=tr("act_machine_hint"), bg=PAPER, fg=MUTED,
                 font=F["F_FOOT"], anchor="w", justify="left",
                 wraplength=470).pack(fill="x", padx=24, pady=(2, 6))

        def _copy_mc():
            try:
                win.clipboard_clear()
                win.clipboard_append(mc_var.get())
                messagebox.showinfo(tr("act_copy_ok_title"), tr("act_copy_ok"),
                                    parent=win)
            except Exception:
                pass

        ttk.Button(win, text=tr("act_copy_btn"), style="TButton",
                   command=_copy_mc).pack(anchor="w", padx=24, pady=(0, 8))

        tk.Frame(win, bg=LINE, height=1).pack(fill="x", padx=24, pady=6)

        # ③ 离线激活码（卖家签发，无需联网）
        tk.Label(win, text=tr("act_offline_label"), bg=PAPER, fg=INK,
                 font=F["F_SMALL_B"], anchor="w").pack(fill="x", padx=24, pady=(6, 2))
        off_var = tk.StringVar()
        ttk.Entry(win, textvariable=off_var, width=44,
                  font=F["F_MONO"]).pack(padx=24, fill="x")
        tk.Label(win, text=tr("act_offline_hint"), bg=PAPER, fg=MUTED,
                 font=F["F_FOOT"], anchor="w").pack(fill="x", padx=24, pady=(2, 6))

        def _submit_offline():
            code = off_var.get().strip()
            if not code:
                messagebox.showwarning(tr("act_title"), tr("act_need_offline"), parent=win)
                return
            r = lic.activate_offline(code)
            messagebox.showinfo(tr("act_result_title"), r["message"], parent=win)
            if r["ok"]:
                self._close_one(win)

        ttk.Button(win, text=tr("act_offline_btn"), style="TButton",
                   command=_submit_offline).pack(anchor="w", padx=24, pady=(0, 14))

    # ------------------------------------------------------------ 目标画像
    def _build_target(self) -> TargetProfile:
        return build_target(
            self.spec_key.get(),
            template=self.school_template.get() or None,
            ai=self.ai_json.get() or None,
            questionnaire=self.questionnaire_data or None,
            latex=self.latex_template.get() or None,
        )

    # ------------------------------------------------------------ 体检
    def _run(self):
        if self._busy:
            return
        docx = self.docx_path.get()
        if not docx or not os.path.isfile(docx):
            messagebox.showwarning(self.tr("msg_missing_paper_title"),
                                   self.tr("msg_missing_paper"))
            return
        tr = self.tr
        self._set_busy(True)
        self._set_status(tr("status_checking"), "run")
        self._set_bar("run")
        try:
            target = self._build_target()
            prof, findings = checker.run_check(docx, target)
            rep = report_mod.build_report(docx, target, prof, findings)
        except DocxReadError as e:
            self._fail(tr("msg_read_error_title"), str(e))
            return
        except (ValueError, FileNotFoundError, OSError) as e:
            self._fail(tr("msg_check_error_title"), str(e))
            return
        except Exception as e:  # noqa: BLE001  兜底，界面绝不能因未知异常僵死
            self._fail(tr("msg_check_error_unexpected"), str(e))
            return
        finally:
            self._set_busy(False)
        self._set_report(rep.markdown)
        self._step_index = max(self._step_index, 2)
        self._refresh_wizard()
        self._set_status(tr("bar_check_done"), "ok")
        self._set_bar("ok")

    def _fail(self, title: str, detail: str):
        self._set_busy(False)
        self._set_status(detail, "err")
        self._set_bar("err")
        messagebox.showerror(title, detail)

    # ------------------------------------------------------------ 一键修正
    def _run_fix(self):
        """一键修正：先算改动（免费）→ 需动手才问授权 → 预览确认 → 落盘。

        顺序与 CLI 一致：**已合规文档不触发门禁**（否则「文档本来就不用改」也会被
        说成「试用已用完」，既误导用户又等于劝人白买码）。
        """
        if self._busy:
            return
        tr = self.tr
        docx = self.docx_path.get()
        if not docx or not os.path.isfile(docx):
            messagebox.showwarning(tr("msg_missing_paper_title"), tr("msg_missing_paper"))
            return

        self._set_busy(True)
        self._set_status(tr("status_fixing"), "run")
        self._set_bar("run", "bar_fixing")
        try:
            target = self._build_target()
            changes = compute_changes(docx, target)
        except DocxReadError as e:
            self._fail(tr("msg_read_error_title"), str(e))
            return
        except (ValueError, FileNotFoundError, OSError) as e:
            self._fail(tr("msg_fix_preview_error_title"), str(e))
            return
        except Exception as e:  # noqa: BLE001
            self._fail(tr("msg_fix_error_unexpected"), str(e))
            return

        if not changes:
            self._set_busy(False)
            self._set_report(tr("msg_no_change"))
            self._set_status(tr("msg_no_change"), "ok")
            self._set_bar("ok")
            messagebox.showinfo(tr("msg_no_change_title"), tr("msg_no_change"))
            return

        # 真要动手改，才做授权裁决（首次免费、之后需激活码）
        gate = lic.require_fix_entitlement()
        if not gate["allowed"]:
            self._set_busy(False)
            self._set_status(gate["message"], "err")
            self._set_bar("err")
            messagebox.showwarning(tr("msg_need_activation_title"), gate["message"])
            return

        preview_text = tr("preview_header") + "\n" + \
            "\n".join("  · " + line for line in changes) + \
            "\n\n" + tr("preview_trial", gate["message"])
        self._set_report(preview_text)
        self._set_busy(False)
        if not messagebox.askyesno(tr("msg_fix_confirm_title"),
                                   preview_text + tr("msg_fix_confirm_body")):
            self._set_bar("idle")
            return

        default_name = os.path.splitext(os.path.basename(docx))[0] + "_fixed.docx"
        out = filedialog.asksaveasfilename(
            title=tr("msg_fix_pick_output"), defaultextension=".docx",
            filetypes=[("Word", "*.docx")], initialfile=default_name,
            initialdir=os.path.dirname(docx) or None)
        if not out:
            self._set_bar("idle")
            return
        # 保存对话框已自带覆盖确认，这里只拦「与原件同一文件」这一种不可逆情形
        if same_file(out, docx):
            messagebox.showerror(tr("msg_cannot_overwrite_title"),
                                 tr("msg_cannot_overwrite"))
            self._set_bar("idle")
            return

        self._set_busy(True)
        try:
            fix_docx(docx, out, target)
            counted = lic.record_fix_used()
        except DocxReadError as e:
            self._fail(tr("msg_read_error_title"), str(e))
            return
        except (ValueError, OSError) as e:
            self._fail(tr("msg_fix_error_title"), str(e))
            return
        finally:
            self._set_busy(False)

        self._step_index = max(self._step_index, 3)
        self._refresh_wizard()
        self._set_status(tr("bar_fix_done"), "ok")
        self._set_bar("ok")
        # 展示**扣减后**的状态（gate['message'] 是扣减前的，会显示成「首次免费试用」）
        messagebox.showinfo(tr("msg_fix_done_title"),
                            "%s\n%s\n\n%s" % (tr("msg_fix_done_before_file"), out,
                                              lic.post_fix_message(gate, counted)))

    # ------------------------------------------------------------ 报告 / 模板
    def _save_report(self):
        tr = self.tr
        content = self._report_text.strip()
        if not content:
            messagebox.showwarning(tr("save_no_report_title"), tr("save_no_report"))
            return
        dst = filedialog.asksaveasfilename(
            title=tr("save_report_dialog"), defaultextension=".md",
            filetypes=[("Markdown", "*.md"), ("All files", "*.*")],
            initialfile="format-check-report.md")
        if not dst:
            return
        try:
            with open(dst, "w", encoding="utf-8") as fh:
                fh.write(content + "\n")
        except OSError as e:
            messagebox.showerror(tr("msg_fix_error_title"), str(e))
            return
        messagebox.showinfo(tr("save_report_done_title"),
                            "%s\n%s" % (tr("save_report_done"), dst))

    def _export_template(self):
        tr = self.tr
        dst = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON", "*.json")],
            initialfile="ai_questionnaire_template.json")
        if not dst:
            return
        try:
            with open(_TEMPLATE_PATH, "r", encoding="utf-8") as src, \
                    open(dst, "w", encoding="utf-8") as out:
                out.write(src.read())
        except OSError as e:
            messagebox.showerror(tr("msg_fix_error_title"), str(e))
            return
        messagebox.showinfo(tr("export_done_title"), tr("export_done", dst))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
