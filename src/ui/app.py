"""Thesis Format Doctor Global —— tkinter 桌面 GUI（离线，论文不出本机）。

界面：象牙白纸底 + 学术深蓝主色。顶栏（品牌 / 标语 / 隐私声明 / 语言·关于·帮助）
→ 两栏卡片（左「Your Documents」/ 右「Review & Fix」）→ 页脚（邮箱 / 官网 / 标语）
→ 状态栏。

v2.2.0 按设计规格书重做视觉层（老板 2026-09-12 交办），**功能一行未动**：

- 视觉层拆成三块，各管一摊：
  - ``theme.py``   —— 颜色 / 字号 / 间距令牌（唯一来源，改色只改这里）
  - ``widgets.py`` —— 圆角卡片 / 圆角按钮 / 步进圆点（tkinter 没有圆角与投影，
    这三个是 Canvas 自绘）
  - 本文件         —— 布局骨架 + 交互状态机
- 旧版的竖排时间线（3 行带长描述）改成**横排 stepper**；1/3 计数、章节角标
  （Ⅰ·SELECT / Ⅱ·PROCESS）按规格删除。
- 问卷 / 配置导入两个常驻行收进 **Advanced options 折叠区**。
- 主按钮改成**上下文相关**：未选论文=置灰、就绪=Check formatting、
  出结果=Fix N issues、改完=Review changes（见 ``_sync_ui``）。

与国内版的差异（老板 2026-09-11 定）：
1. **默认英文**，右上角一键切中文；语言偏好落盘，下次启动记得住。
2. 英文字体走 Georgia（品牌/标题衬线）+ Segoe UI（正文），不套中文字体。
3. **品牌口径只留官网与邮箱**，不出现微信 / 公众号 / 小程序（海外用户不用）。

功能面（不变）：
  选规范 →（可选）学校模板 / LaTeX / 问卷 / 配置导入 → 选论文
  → 运行体检（免费不限次）→ 一键修正（首次免费，之后需激活码）。
GUI 与 CLI 共用同一套 engine，保证逻辑唯一、离线安全。
"""

from __future__ import annotations

import os
import sys
import tkinter as tk
import traceback
import webbrowser
from tkinter import filedialog, messagebox, ttk

from ..engine import checker, report as report_mod
from ..engine.docx_reader import DocxReadError
from ..engine.fixer import compute_changes, fix_docx, same_file
from ..engine.questionnaire import (
    FIELD_RANGES, QUESTIONNAIRE_SCHEMA, TargetProfile, build_target,
)
from ..engine.specs import SPEC_ORDER
from ..license import license as lic
from .. import versioninfo
from . import i18n, iconpath, theme
from .fonts import Fonts
from .widgets import RoundButton, RoundCard, ScrollArea, StepCircle

# ---------------------------------------------------------------------------
# 配色别名 —— 新令牌在 theme.py，这里保留旧名字是为了**没重写的弹窗**（关于 /
# 帮助 / 问卷 / 激活）能一行不改地跟着换成新配色。
# 主色从「黛蓝 + 朱砂」改为「学术深蓝」：朱砂（砖红）按规格**不再作主操作色**，
# 只降级为点缀（GOLD），红色仅保留给真实错误。
# ---------------------------------------------------------------------------
PAPER = theme.BG              # 页面底（象牙纸）
PANEL = theme.SURFACE_SOFT    # 柔和面（信息框 / 状态栏）
CARD = theme.SURFACE          # 卡片面
INK = theme.TEXT              # 主文字
BODY = theme.TEXT             # 正文（弹窗用）
MUTED = theme.TEXT_2          # 次要文字
LINE = theme.BORDER           # 边框 / 细线
ACCENT = theme.PRIMARY        # 主色（学术深蓝）
CINNABAR = theme.GOLD         # 旧「朱砂」→ 降级为点缀金（不再是主按钮色）
CINNABAR_D = "#A5894B"        # 点缀金的 active 态
OKC = theme.SUCCESS           # 成功 / 已完成
ERRC = theme.ERROR            # **仅**真实错误

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
        # 绑在**弹窗的 Toplevel** 上而不是 bind_all：bindtag 含 toplevel，
        # 天然覆盖整个弹窗，且弹窗销毁时绑定自动消失（bind_all 会留下死回调，
        # 之后每次滚轮都去评估一个已删除的 Tcl 命令）。
        # ``winfo_ismapped`` 再兜一层，避免弹窗已关仍在响应。
        try:
            if not canvas.winfo_ismapped():
                return
        except Exception:
            return
        try:
            canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")
        except Exception:
            pass

    try:
        canvas.winfo_toplevel().bind("<MouseWheel>", _wheel, add="+")
    except Exception:
        pass
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
    # 首帧来不及算宽度，延后 30ms 补一次。**必须能取消**：弹窗在 30ms 内被关掉时，
    # 挂起的 after 会去执行一个已被删除的 Tcl 命令（`invalid command name "<lambda>"`），
    # Windows 上 Tcl 的 bgerror 会弹模态框把进程卡住。
    aid = label.after(30, _apply)

    def _cancel(_e=None):
        try:
            label.after_cancel(aid)
        except Exception:
            pass

    label.bind("<Destroy>", _cancel, add="+")


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
        self._report_summary = ""
        self._report_summary_spec = None
        self._report_is_placeholder = True
        self._step_index = 0
        self._busy = False
        self._closing = False        # 根窗口已销毁（见 _on_destroy）
        self._resize_after = None    # 启动后首帧的延迟重排句柄
        # v2.2.0 状态机：phase 决定主按钮文案/可用态与步进位置（见 _sync_ui）
        self._phase = "idle"          # idle / ready / results / fixed / error
        self._last_issues = None      # 最近一次体检的问题数（warn+fail）
        self._checked_path = ""       # 出结论时那份文档 —— 换了文档就作废旧结论
        self._last_output = ""        # 最近一次修正落盘路径（供「查看改动」）
        self._status_spec = (None, "idle", "")   # 状态行最近一次内容（切语言后重渲染）

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
        root.geometry("%dx%d" % (min(_BASE_W, int(_sw * 0.88)), min(1020, int(_sh * 0.94))))
        # minsize 必须**按屏幕夹一次**：写死 1100×720 的话，在 1280×720 这类小屏
        # （或 1366×768 / 150% 缩放）上 WM 会把窗口强行撑到比可用区域还高，
        # side="bottom" 的页脚与状态栏直接落到屏幕外 —— 正好推翻 v2.1.1 修过的 P0。
        root.minsize(max(720, min(theme.WIN_MIN_W, _sw - 60)),
                     max(480, min(theme.WIN_MIN_H, _sh - 120)))
        self._apply_window_icon()
        root.bind("<Configure>", self._on_resize)
        # 关窗/销毁时要把挂着的 after 取消，否则 Tcl 会去执行一个已删除的命令
        # （`invalid command name "<lambda>"`），Windows 上 Tcl 的 bgerror 会弹**模态框
        # 并把进程卡住** —— 表现就是「刚启动就关窗口，进程退不掉」（测试里表现为偶发挂起）。
        root.bind("<Destroy>", self._on_destroy, add="+")
        # Tk 回调里的未捕获异常：默认实现走 messagebox（**模态、阻塞事件循环**），
        # 用户看到的是一个卡死的窗口。改成只记日志 —— 重要错误本来就已经走状态行了。
        root.report_callback_exception = self._on_tk_error
        # 滚轮统一挂一次再按指针位置分发（add="+" 是为了跟弹窗自己的滚轮绑定共存）
        try:
            root.bind_all("<MouseWheel>", self._on_wheel, add="+")
        except Exception:
            pass

        self._build()
        self._resize_after = self.root.after(120, self._on_resize)

    def _on_destroy(self, e=None):
        """根窗口销毁：标记关闭 + 取消挂起的 after（见 __init__ 的说明）。"""
        try:
            if e is not None and getattr(e, "widget", None) is not self.root:
                return
        except Exception:
            pass
        self._closing = True
        aid = getattr(self, "_resize_after", None)
        if aid is not None:
            try:
                self.root.after_cancel(aid)
            except Exception:
                pass
            self._resize_after = None

    def _on_tk_error(self, exc, val, tb):
        """Tk 回调异常的统一出口：写 stderr，绝不弹模态框（否则界面直接卡死）。"""
        try:
            sys.stderr.write("".join(traceback.format_exception(exc, val, tb)))
        except Exception:
            pass

    def _on_wheel(self, e):
        """把滚轮事件转给主区域的滚动区（内部自己判断指针是否在其中）。"""
        area = getattr(self, "_main_scroll", None)
        if area is not None:
            area.handle_wheel(e)

    def _relayout_all(self):
        """内容尺寸可能变了 → 让两张卡与主滚动区**重新测量**。

        必须在 pack/unpack 或长文案切换之后调用：卡片内容区与滚动区内容区的尺寸都被
        钉死，往里加东西不会触发 ``<Configure>``，不主动喊一声就会出现
        「内容涨了但卡片没涨、滚动条也没出现」→ 用户看不到也滚不到
        （Codex 2026-09-12 P0-2）。这个方法是那类问题的**唯一收口**。
        """
        for card in (getattr(self, "_left_card", None),
                     getattr(self, "_right_card", None)):
            if card is not None:
                try:
                    card.refresh()
                except Exception:
                    pass
        area = getattr(self, "_main_scroll", None)
        if area is not None:
            try:
                area.refresh_layout()
            except Exception:
                pass

    # ------------------------------------------------------------ 基础工具
    def tr(self, key: str, *args):
        return i18n.t(self.lang, key, *args)

    def _sep(self) -> str:
        return "｜" if self.lang == "zh" else "|"

    def _apply_window_icon(self):
        """设置标题栏/任务栏图标。**整个函数绝不允许向外抛异常**——图标是锦上添花，
        绝不能因为一个平台上的路径/Tcl 怪癖就让 App 构造失败（那正是「双击打不开」
        那一类 P0 后果）。

        两路**互相独立**的保险（一路失败绝不牵连另一路）：

        ① PNG + ``iconphoto``：跨平台标准做法。``PhotoImage`` 必须挂到实例属性上
           保留引用——只作为临时对象传进去时一旦被 GC 回收，部分 Tk 版本会把图标
           还原成默认（经典坑，v2.0.2 踩过）。
        ② Windows 的 ``.ico`` + ``iconbitmap``：个别环境的 Tcl 没编进 PNG 解码器
           （Tk 安装不完整），① 会静默失败；此时用仓库自带的 ``assets/icon.ico``
           兜底 —— ``iconbitmap`` 读原生 ico，不依赖 PNG 解码器。
        """
        try:
            png = iconpath.find_icon()
            if png:
                try:
                    self._window_icon = tk.PhotoImage(file=png)
                    self.root.iconphoto(True, self._window_icon)
                except Exception:
                    self._window_icon = None
            if sys.platform.startswith("win"):
                ico = iconpath.find_icon_ico()
                if ico:
                    try:
                        self.root.iconbitmap(default=ico)
                    except Exception:
                        pass
        except Exception:
            pass

    def _on_resize(self, _evt=None):
        if getattr(self, "_closing", False):
            return
        # 背景与顶栏是"画上去"的，不吃 pack 的空间分配，所以每次尺寸变化都要自己
        # 重新定位；必须在下面的"缩放没变就早退"之前做，否则宽度变了它们不跟着走。
        self._place_backdrop()
        self._place_header()
        try:
            w = self.root.winfo_width()
        except Exception:
            return
        factor = self.F.scale_for_width(w)
        if self._last_scale is not None and abs(factor - self._last_scale) < 0.02:
            return
        self._last_scale = factor
        self.F.apply_scale(factor)
        # 字体缩放会让顶栏变高 → 主体顶部留白要跟着重算（否则顶栏会和卡片叠上）
        try:
            self._main_scroll.pack_configure(pady=(self._top_gap(), 0))
        except Exception:
            pass

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
        """ttk 样式 —— 主操作色统一为学术深蓝（砖红不再作主操作色，规格第 5 节第 4 条）。

        主/次 CTA 走 ``widgets.RoundButton`` 自绘（要圆角）；这里配置的是弹窗里的
        普通按钮、卡片内辅助动作按钮与下拉框。
        """
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        F = self.F
        style.configure("TButton", font=F["F_BODY"], padding=(12, 7),
                        background=theme.SURFACE, foreground=theme.PRIMARY,
                        bordercolor=theme.BORDER, lightcolor=theme.SURFACE,
                        darkcolor=theme.SURFACE)
        style.map("TButton",
                  background=[("active", theme.PRIMARY_SOFT),
                              ("disabled", theme.PRIMARY_SOFT)],
                  foreground=[("disabled", theme.TEXT_3)])
        style.configure("Action.TButton", font=F["F_BTN"], padding=(12, 10),
                        background=theme.PRIMARY, foreground="#FFFFFF",
                        bordercolor=theme.PRIMARY)
        style.map("Action.TButton",
                  background=[("active", theme.PRIMARY_HOVER),
                              ("disabled", theme.BTN_DISABLED_BG)],
                  foreground=[("disabled", theme.BTN_DISABLED_FG)])
        style.configure("Primary.TButton", font=F["F_BTN"], padding=(12, 10),
                        background=theme.PRIMARY, foreground="#FFFFFF",
                        bordercolor=theme.PRIMARY)
        style.map("Primary.TButton",
                  background=[("active", theme.PRIMARY_HOVER),
                              ("disabled", theme.BTN_DISABLED_BG)],
                  foreground=[("disabled", theme.BTN_DISABLED_FG)])
        # 辅助动作（保存报告 / 导出模板 / 激活）：低调描边，不与主 CTA 抢注意力
        style.configure("Ghost.TButton", font=F["F_BTN_S"], padding=(10, 6),
                        background=theme.SURFACE, foreground=theme.PRIMARY,
                        bordercolor=theme.BORDER, relief="flat")
        style.map("Ghost.TButton",
                  background=[("active", theme.PRIMARY_SOFT),
                              ("disabled", theme.SURFACE)],
                  foreground=[("disabled", theme.TEXT_3)])
        style.configure("TCombobox", font=F["F_BODY"], padding=4)

    # ------------------------------------------------------------ 整体布局
    def _build(self):
        """搭骨架：背景层 → 底部区占位 → 顶栏 → 两栏卡片。

        pack 顺序有讲究：页脚与状态栏用 ``side="bottom"`` **先** pack，底部区优先
        拿到空间，永远不会被主体挤掉（v2.1.1 老板反馈过「看不到页脚」，不能回退）。

        背景层用 ``place`` 铺满窗口且**最先创建** —— Tk 层叠按创建顺序，先建的在下层，
        于是它天然是最底层，卡片/文字/按钮都是它之上的独立控件（规格硬要求）。
        """
        self._build_style()
        self.root.configure(bg=theme.BG)

        self._build_backdrop()
        self._build_footer()
        self._build_statusbar()
        self._build_header()

        # 主体：可滚动区。窗口比内容矮时（1366×768 / 1600×900 笔记本）出滚动条，
        # 绝不把卡片底部裁掉 —— 旧版没有这层，左卡最下面的字段会被圆角矩形切掉。
        # 顶部留白 = 顶栏高度 + HEADER_GAP：顶栏画在背景上不占 pack 空间，而这段留白
        # 下面是背景画布，所以**背景（含天际线）正好从这条带露出来**。
        self._main_scroll = ScrollArea(self.root, bg=theme.BG)
        self._main_scroll.pack(fill="both", expand=True, padx=theme.PAGE_PAD,
                               pady=(self._top_gap(), 0))

        main = tk.Frame(self._main_scroll.inner, bg=theme.BG)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1, uniform="half", minsize=460)
        main.columnconfigure(1, weight=1, uniform="half", minsize=460)
        main.rowconfigure(0, weight=1)

        # 两栏等宽；grid 两边各给半个 CARD_GAP，合计正好一个卡片间距
        self._left_card = RoundCard(main, padx=theme.CARD_PAD_X,
                                    pady=theme.CARD_PAD_Y)
        self._left_card.grid(row=0, column=0, sticky="nsew",
                             padx=(0, theme.CARD_GAP // 2))
        self._right_card = RoundCard(main, padx=theme.CARD_PAD_X,
                                     pady=theme.CARD_PAD_Y)
        self._right_card.grid(row=0, column=1, sticky="nsew",
                              padx=(theme.CARD_GAP // 2, 0))
        self._build_left(self._left_card.inner)
        self._build_right(self._right_card.inner)

        # 重建后恢复界面状态（切语言会走这条路）
        if self._report_is_placeholder or not self._report_text:
            self._set_report(self.tr("report_placeholder"), placeholder=True)
        elif self._report_summary_spec is not None:
            self._set_report(self._report_text,
                             summary_spec=self._report_summary_spec)
        else:
            self._set_report(self._report_text, summary=self._report_summary)
        self._refresh_rows()
        self._set_bar(self._bar_state)
        # 状态行同样按新语言重渲染 —— 不恢复的话，切语言后「发现 18 个格式问题」
        # 这类结论会凭空消失（主按钮却还写着 Fix 18 issues，看起来像自相矛盾）。
        spec, kind, raw = getattr(self, "_status_spec", (None, "idle", ""))
        text = self.tr(spec[0], *(spec[1] or ())) if spec else raw
        if text:
            self._set_status(text, kind, spec=spec)

    # ------------------------------------------------------------ 背景 / 顶栏
    def _build_backdrop(self):
        """最底层装饰背景（视觉规格包 2026-09-12 给的独立底图）。

        实现约束（照抄规格 README 的三条硬要求）：
          * **只作最底层背景**，绝不拿整张 UI 稿当背景、也不许从稿子里裁图当控件；
          * **居中靠上、等比、不变形** —— 所以 1:1 画在 ``anchor="n"`` 水平居中处，
            不做任何缩放（tkinter 只有整数 ``subsample``，缩放会很粗糙）；
          * 够淡、不干扰阅读 —— 底图最暗处 #CAD8E2（亮度 213/255），本身就极淡。

        窗口比底图宽时两侧留空：底图的平坦区已在入库时规范成与 ``theme.BG`` 同色，
        所以**看不出接缝**（这正是把素材底色统一成页面色的原因）。

        素材缺失 / Tk 没编 PNG 解码器 → 降级成纯纸底，功能零影响。
        **但打包漏件不许靠这个降级兜过去**：``tools/verify_bundle.py`` 会在 CI 拦下。
        """
        self._backdrop = tk.Canvas(self.root, bg=theme.BG,
                                   highlightthickness=0, bd=0, width=1, height=1)
        # place：不参与 pack 的空间分配，纯铺底层；先创建 = 层叠最下。
        self._backdrop.place(x=0, y=0, relwidth=1, relheight=1)
        self._backdrop_img = None
        self._backdrop_item = None
        self._backdrop_loaded = False
        path = iconpath.find_backdrop()
        if path:
            try:
                self._backdrop_img = tk.PhotoImage(file=path)   # Tk 8.6 原生读 PNG
                self._backdrop_item = self._backdrop.create_image(
                    0, 0, image=self._backdrop_img, anchor="n")
                self._backdrop_loaded = True
            except Exception:
                self._backdrop_img = None
                self._backdrop_item = None
        self._place_backdrop()

    def _place_backdrop(self):
        """底图水平居中、贴顶（规格 ``background.position = center top``）。"""
        if getattr(self, "_backdrop_item", None) is None:
            return
        try:
            w = self.root.winfo_width()
            if w <= 1:
                return
            self._backdrop.coords(self._backdrop_item, w // 2, 0)
        except Exception:
            pass

    def _top_gap(self) -> int:
        """主体上方该留出的高度 = 顶栏高度 + ``HEADER_GAP``。

        顶栏现在画在背景画布上（不占 pack 空间），所以这段高度必须显式留出来；
        而这段留白的背景正是背景画布 —— 天际线就是从这里露出来的。
        """
        return theme.HEADER_PAD_Y + self._header_height() + theme.HEADER_GAP

    def _header_height(self) -> int:
        try:
            return max(self._hdr_brand.winfo_reqheight(),
                       self._hdr_nav.winfo_reqheight())
        except Exception:
            return 70

    def _place_header(self):
        """顶栏右侧导航随窗口宽度贴右边（左侧品牌固定 PAGE_PAD）。"""
        try:
            w = self.root.winfo_width()
            if w <= 1:
                return
            self._backdrop.coords(self._hdr_nav_win, w - theme.PAGE_PAD,
                                  theme.HEADER_PAD_Y)
        except Exception:
            pass

    def _build_header(self):
        """顶栏：品牌 + 标语 + 隐私声明（左）／语言 ▾ · About · Help（右）。

        与前版的两个差别（都是为背景服务）：
          ① 顶栏不再是一整条**不透明**框架，而是画在背景画布上的两个小组件
             （``create_window``）—— 背景素材的天际线正好压在顶栏这条带上，整条盖住
             就等于白放了一张图；
          ② **删掉顶栏下的通栏细线**：参考稿里没有这条线，留着会横穿天际线。
        """
        F = self.F
        cv = self._backdrop

        brand = tk.Frame(cv, bg=theme.BG)
        tk.Label(brand, text=self.tr("app_title"), bg=theme.BG, fg=theme.PRIMARY,
                 font=F["F_BRAND"]).pack(anchor="w")
        tk.Label(brand, text=self.tr("app_tagline"), bg=theme.BG,
                 fg=theme.SECONDARY, font=F["F_SUB"]).pack(anchor="w", pady=(4, 9))
        # 隐私声明：规格要求从卡片里**移到顶栏**（卡片右上角的 Offline 徽标已删除）
        prow = tk.Frame(brand, bg=theme.BG)
        prow.pack(anchor="w")
        tk.Label(prow, text=self.tr("privacy_line"), bg=theme.BG, fg=theme.PRIMARY,
                 font=F["F_SMALL_B"]).pack(side="left")
        tk.Label(prow, text="   " + self.tr("privacy_sub"), bg=theme.BG,
                 fg=theme.TEXT_3, font=F["F_HELP"]).pack(side="left")

        # 右侧导航：自右向左 pack
        nav = tk.Frame(cv, bg=theme.BG)
        self._link(nav, self.tr("link_help"),
                   self._help_dialog).pack(side="right")
        self._nav_sep(nav)
        self._link(nav, self.tr("link_about"),
                   self._about_dialog).pack(side="right")
        self._nav_sep(nav)
        self._link(nav, self.tr("lang_button"),
                   lambda: self._set_lang(i18n.other_lang(self.lang))).pack(side="right")

        self._hdr_brand, self._hdr_nav = brand, nav
        self._hdr_brand_win = cv.create_window(theme.PAGE_PAD, theme.HEADER_PAD_Y,
                                               window=brand, anchor="nw")
        self._hdr_nav_win = cv.create_window(0, theme.HEADER_PAD_Y,
                                             window=nav, anchor="ne")
        self._place_header()

    def _nav_sep(self, parent):
        tk.Label(parent, text="   |   ", bg=theme.BG, fg=theme.BORDER,
                 font=self.F["F_HELP"]).pack(side="right")

    def _build_footer(self):
        """页脚：左邮箱 / 中官网 / 右标语 —— 规格明确「页脚只此三项」。"""
        F = self.F
        footer = tk.Frame(self.root, bg=theme.BG)
        footer.pack(side="bottom", fill="x", padx=theme.PAGE_PAD,
                    pady=(0, theme.FOOTER_PAD_Y))
        mail = tk.Label(footer, text="✉   " + self.tr("footer_email"), bg=theme.BG,
                        fg=theme.SECONDARY, font=F["F_FOOT"], cursor="hand2")
        mail.pack(side="left")
        mail.bind("<Button-1>",
                  lambda e: webbrowser.open("mailto:" + i18n.BRAND_EMAIL))
        tk.Label(footer, text=self.tr("footer_tagline"), bg=theme.BG,
                 fg=theme.TEXT_3, font=F["F_FOOT"]).pack(side="right")
        site = tk.Label(footer, text="◎   " + self.tr("footer_site"), bg=theme.BG,
                        fg=theme.SECONDARY, font=F["F_FOOT"], cursor="hand2")
        site.bind("<Button-1>", lambda e: webbrowser.open(i18n.BRAND_SITE_URL))
        # fill="x" + expand：吃掉两侧剩余空间，标签文字自然落在中间
        site.pack(side="left", fill="x", expand=True)

    def _build_statusbar(self):
        F = self.F
        tk.Frame(self.root, bg=theme.BORDER, height=1).pack(
            fill="x", side="bottom", padx=theme.PAGE_PAD)
        statusbar = tk.Frame(self.root, bg=theme.BG)
        statusbar.pack(side="bottom", fill="x", padx=theme.PAGE_PAD, pady=(6, 8))
        self.bar_dot = tk.Label(statusbar, text="●", bg=theme.BG, fg=OKC,
                                font=F["F_FOOT"])
        self.bar_dot.pack(side="left", padx=(0, 6))
        tk.Label(statusbar, text=self.tr("bar_left", self.version), bg=theme.BG,
                 fg=theme.TEXT_3, font=F["F_FOOT"]).pack(side="left")
        self.bar_right = tk.Label(statusbar, text="", bg=theme.BG, fg=theme.TEXT_3,
                                  font=F["F_FOOT"])
        self.bar_right.pack(side="right")

    # ------------------------------------------------------------ 左栏
    def _build_left(self, parent):
        """左卡 Your Documents：4 个字段 + 折叠的 Advanced options。

        与旧版的关键差别：字段之间靠**留白 + 一根极细线**分区，不再给每个字段套一层
        带边框的小盒子（规格第 5 节第 1 条「去掉过多的嵌套边框」）。
        """
        F = self.F
        tk.Label(parent, text=self.tr("left_title"), bg=theme.SURFACE,
                 fg=theme.PRIMARY, font=F["F_CARD_HDR"], anchor="w").pack(fill="x")
        desc = tk.Label(parent, text=self.tr("left_desc"), bg=theme.SURFACE,
                        fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                        justify="left")
        desc.pack(fill="x", pady=(6, theme.FIELD_GAP))
        _auto_wrap(desc, parent, theme.CARD_PAD_X)

        # ① 引用规范（必选）
        self._build_field_spec(parent)
        self._hairline(parent)

        # ② 论文（必选）—— 虚线投放区
        self._field_label(parent, "row_paper_title", required=True)
        self._helper(parent, "row_paper_helper")
        drop = RoundCard(parent, radius=9, fill="#FCFDFF",
                         border=theme.SECONDARY, shadow=False, dashed=True,
                         padx=18, pady=13)
        drop.pack(fill="x", pady=(0, theme.FIELD_GAP))
        self._paper_name = tk.Label(drop.inner, text="", bg="#FCFDFF",
                                    fg=theme.TEXT_3, font=F["F_BODY"],
                                    cursor="hand2")
        self._paper_name.pack()
        sub = tk.Label(drop.inner,
                       text=self.tr("btn_select_doc") + "   ·   "
                       + self.tr("paper_support"),
                       bg="#FCFDFF", fg=theme.TEXT_3, font=F["F_HELP"],
                       cursor="hand2")
        sub.pack(pady=(5, 0))
        for w in (drop.inner, self._paper_name, sub):
            w.bind("<Button-1>", lambda e: self._pick_docx())

        self._hairline(parent)
        # ③ 学校模板（可选）
        self._field_label(parent, "row_template_title", required=False)
        self._helper(parent, "row_template_desc")
        self._tpl_name = self._select_field(parent, "row_template_placeholder",
                                            self._pick_template)

        self._hairline(parent)
        # ④ LaTeX 模板（可选）
        self._field_label(parent, "row_latex_title", required=False)
        self._helper(parent, "row_latex_desc")
        self._latex_name = self._select_field(parent, "row_latex_placeholder",
                                              self._pick_latex)

        self._build_advanced(parent)

    def _hairline(self, parent):
        """字段之间的极细分隔线。

        只给**上方** 16px：下方的间距由紧跟着的 ``_field_label`` 提供。
        旧写法两边各给一个 FIELD_GAP，两个字段之间白吃 36px 高度，左卡一下就撑到 700+。
        """
        tk.Frame(parent, bg=theme.HAIRLINE, height=1).pack(fill="x", pady=(16, 0))

    def _field_label(self, parent, key: str, required: bool = False):
        """字段标签 + 小小的 Required/Optional 徽标。

        规格第 5 节第 5 条：必选/可选标签要**小、中性偏蓝、不吓人** ——
        旧版用朱砂红，看着像报错，已改。
        """
        F = self.F
        row = tk.Frame(parent, bg=theme.SURFACE)
        row.pack(fill="x", pady=(6, 0))
        tk.Label(row, text=self.tr(key), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_LABEL"]).pack(side="left")
        badge_text = self.tr("mark_required" if required else "mark_optional")
        badge_color = theme.PRIMARY_SOFT if required else theme.SURFACE_SOFT
        badge_fg = theme.PRIMARY if required else theme.TEXT_2
        tk.Label(row,
                 text="  " + badge_text + "  ",
                 bg=badge_color, fg=badge_fg,
                 font=F["F_HELP"], padx=6, pady=1).pack(side="left", padx=(7, 0))
        return row

    def _helper(self, parent, key: str):
        F = self.F
        lbl = tk.Label(parent, text=self.tr(key), bg=theme.SURFACE,
                       fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                       justify="left")
        lbl.pack(fill="x", pady=(4, 8))
        _auto_wrap(lbl, parent, theme.CARD_PAD_X)
        return lbl

    def _build_field_spec(self, parent):
        """引用规范（必选）：下拉框 + 一行「Other 用法」轻提示。"""
        F = self.F
        self._field_label(parent, "row_spec_title", required=True)
        self._helper(parent, "row_spec_desc")
        self._spec_cb = ttk.Combobox(parent, textvariable=self.spec_key,
                                     values=SPEC_ORDER, state="readonly",
                                     font=F["F_BODY"])
        self._spec_cb.pack(fill="x", pady=(0, 6))
        # 下拉**弹出列表**的字体不受 widget 的 font 影响，得单独喂给 option 库，
        # 否则弹出来还是 Tk 默认字体（观感直接掉档）。
        try:
            self.root.option_add("*TCombobox*Listbox.font", F["F_BODY"])
        except Exception:
            pass
        # 换引用规范会让上一次的体检结论失效（否则结论与所选规范对不上）
        try:
            self.spec_key.trace_add("write", lambda *a: self._on_input_changed())
        except Exception:
            pass
        hint = tk.Label(parent, text=self.tr("chip_style_hint"), bg=theme.SURFACE,
                        fg=theme.GOLD, font=F["F_HELP"], anchor="w",
                        justify="left")
        hint.pack(fill="x")
        _auto_wrap(hint, parent, theme.CARD_PAD_X)

    def _select_field(self, parent, placeholder_key: str, cmd):
        """「下拉框外观 + 文件选择行为」的控件。

        规格画的是 ``Choose a template ▾`` 这样的下拉；我们的真实动作是打开文件
        对话框（本机离线读模板），所以外观照规格做、行为照实际来：整行可点。
        """
        F = self.F
        box = tk.Frame(parent, bg=theme.BG, highlightthickness=1,
                       highlightbackground=theme.FIELD_BORDER,
                       highlightcolor=theme.PRIMARY)
        box.pack(fill="x", pady=(0, 2))
        inner = tk.Frame(box, bg=theme.SURFACE)
        inner.pack(fill="x", padx=11, pady=9)
        lbl = tk.Label(inner, text=self.tr(placeholder_key), bg=theme.SURFACE,
                       fg=theme.TEXT_3, font=F["F_BODY"], anchor="w",
                       cursor="hand2")
        lbl.pack(side="left", fill="x", expand=True)
        tk.Label(inner, text="▾", bg=theme.SURFACE, fg=theme.SECONDARY,
                 font=F["F_BODY"], cursor="hand2").pack(side="right")

        def _click(_e=None):
            cmd()

        for w in (box, inner, lbl):
            w.bind("<Button-1>", _click)
        return lbl

    def _build_advanced(self, parent):
        """Advanced options 折叠区（默认收起）：问卷 / 导入配置 / AI 辅助。

        规格第 5 节第 7 条：技术性功能收进这里。旧版把问卷与 AI 填表 JSON 当作两个
        常驻行摆在明面上，既占版面、也抬高理解成本。
        """
        F = self.F
        self._adv_open = False
        self._hairline(parent)

        hdr = tk.Frame(parent, bg=theme.SURFACE, cursor="hand2")
        hdr.pack(fill="x", pady=(6, 0))
        self._adv_chev = tk.Label(hdr, text="▸  ", bg=theme.SURFACE,
                                  fg=theme.SECONDARY, font=F["F_LABEL"],
                                  cursor="hand2")
        self._adv_chev.pack(side="left")
        tk.Label(hdr, text=self.tr("adv_label"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_LABEL"], cursor="hand2").pack(side="left")

        self._adv_hint = tk.Label(parent, text=self.tr("adv_hint"), bg=theme.SURFACE,
                                  fg=theme.TEXT_3, font=F["F_HELP"], anchor="w",
                                  cursor="hand2")
        self._adv_hint.pack(fill="x", padx=(16, 0), pady=(3, 0))

        self._adv_body = tk.Frame(parent, bg=theme.SURFACE)

        # ① 格式问卷
        r1 = tk.Frame(self._adv_body, bg=theme.SURFACE)
        r1.pack(fill="x", pady=(2, 7))
        c1 = tk.Frame(r1, bg=theme.SURFACE)
        c1.pack(side="left", fill="x", expand=True)
        tk.Label(c1, text=self.tr("row_questionnaire_title"), bg=theme.SURFACE,
                 fg=theme.TEXT, font=F["F_SUBTITLE"], anchor="w").pack(fill="x")
        tk.Label(c1, text=self.tr("row_questionnaire_desc"), bg=theme.SURFACE,
                 fg=theme.TEXT_2, font=F["F_HELP"], anchor="w").pack(fill="x")
        self._q_name = tk.Label(c1, text="", bg=theme.SURFACE, fg=theme.TEXT_3,
                                font=F["F_HELP"], anchor="w")
        self._q_name.pack(fill="x", pady=(3, 0))
        ttk.Button(r1, text=self.tr("btn_questionnaire"), style="Ghost.TButton",
                   command=self._open_questionnaire).pack(side="right")

        # ② 导入配置（AI 填表 JSON）
        r2 = tk.Frame(self._adv_body, bg=theme.SURFACE)
        r2.pack(fill="x", pady=(0, 2))
        c2 = tk.Frame(r2, bg=theme.SURFACE)
        c2.pack(side="left", fill="x", expand=True)
        tk.Label(c2, text=self.tr("row_ai_title"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_SUBTITLE"], anchor="w").pack(fill="x")
        tk.Label(c2, text=self.tr("row_ai_desc"), bg=theme.SURFACE,
                 fg=theme.TEXT_2, font=F["F_HELP"], anchor="w").pack(fill="x")
        self._ai_name = tk.Label(c2, text="", bg=theme.SURFACE, fg=theme.TEXT_3,
                                 font=F["F_HELP"], anchor="w")
        self._ai_name.pack(fill="x", pady=(3, 0))
        ttk.Button(r2, text=self.tr("btn_ai_json"), style="Ghost.TButton",
                   command=self._pick_ai_json).pack(side="right")

        for w in (hdr, self._adv_chev, self._adv_hint):
            w.bind("<Button-1>", lambda e: self._toggle_advanced())

    def _toggle_advanced(self):
        """展开 / 收起 Advanced options（返回展开后的状态，便于测试断言）。"""
        self._adv_open = not self._adv_open
        try:
            if self._adv_open:
                self._adv_body.pack(fill="x", pady=(theme.SECTION_GAP, 0))
                self._adv_chev.config(text="▾  ")
                self._adv_hint.pack_forget()
            else:
                self._adv_body.pack_forget()
                self._adv_chev.config(text="▸  ")
                self._adv_hint.pack(fill="x", padx=(16, 0), pady=(3, 0))
        except Exception:
            pass
        # 展开/收起改变了左卡内容高度 —— 必须主动重新测量，
        # 否则新出现的那两行会被圆角矩形裁掉且滚不到（见 _relayout_all）。
        self._relayout_all()
        return self._adv_open

    # ------------------------------------------------------------ 右栏
    def _build_right(self, parent):
        F = self.F
        tk.Label(parent, text=self.tr("right_title"), bg=theme.SURFACE,
                 fg=theme.PRIMARY, font=F["F_CARD_HDR"], anchor="w").pack(fill="x")
        desc = tk.Label(parent, text=self.tr("right_desc"), bg=theme.SURFACE,
                        fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                        justify="left")
        desc.pack(fill="x", pady=(6, 0))
        _auto_wrap(desc, parent, theme.CARD_PAD_X)

        self._build_stepper(parent)
        self._build_how_card(parent)

        # 状态行：结果与提示都落这里（旧版状态行保留，位置改到按钮上方）
        status_row = tk.Frame(parent, bg=theme.SURFACE)
        status_row.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        self._status_row = status_row
        self.status_dot = tk.Label(status_row, text="●", bg=theme.SURFACE,
                                   fg=theme.TEXT_3, font=F["F_HELP"])
        self.status_dot.pack(side="left", padx=(0, 7))
        self.status_lbl = tk.Label(status_row, text="", bg=theme.SURFACE,
                                   fg=theme.TEXT, font=F["F_STAT"], anchor="w",
                                   justify="left")
        self.status_lbl.pack(side="left", fill="x", expand=True)
        _auto_wrap(self.status_lbl, parent, theme.CARD_PAD_X + 16)

        # 结果明细（通过 / 提示 / 警告 / 不达标 + 报告导出提示）
        self.summary_lbl = tk.Label(parent, text="", bg=theme.SURFACE,
                                    fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                                    justify="left")
        self.summary_lbl.pack(fill="x", pady=(3, 0))
        _auto_wrap(self.summary_lbl, parent, theme.CARD_PAD_X)
        self.summary_hint = tk.Label(parent, text=self.tr("report_summary_hint"),
                                     bg=theme.SURFACE, fg=theme.TEXT_3,
                                     font=F["F_HELP"], anchor="w", justify="left")
        self.summary_hint.pack(fill="x")
        _auto_wrap(self.summary_hint, parent, theme.CARD_PAD_X)

        # 辅助动作（技术性功能，低调）。**必须先建**：下面 `_fix_btn` 的 pack 要以它为锚点。
        aux = tk.Frame(parent, bg=theme.SURFACE)
        aux.pack(fill="x", pady=(10, 0))
        self._aux_row = aux          # 留引用：回归测试要断言次按钮排在它**上面**
        for text, cmd in ((self.tr("btn_save_report"), self._save_report),
                          (self.tr("btn_export_ai"), self._export_template),
                          (self.tr("btn_activate"), self._activate_dialog)):
            ttk.Button(aux, text=text, style="Ghost.TButton",
                       command=cmd).pack(side="left", padx=(0, 6))

        # 主 CTA（视觉绝对主导）+ 次按钮（只在「就绪」态出现）
        self._check_btn = RoundButton(parent, self.tr("btn_check"), self._run,
                                      style="primary", font=F["F_BTN"], height=52)
        self._check_btn.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        # before=aux 必须记进 pack 参数：`set_visible` 是 pack_forget 后再 pack，
        # 不记锚点就会把它排到队尾 —— 次按钮会掉到三个 ghost 按钮**下面**，
        # 破坏"次要动作紧跟主 CTA"的层级（Codex 2026-09-12 P1-1）。
        self._fix_btn = RoundButton(parent, self.tr("btn_fix"), self._run_fix,
                                    style="secondary", font=F["F_BTN_S"], height=52)
        self._fix_btn.pack(fill="x", pady=(8, 0), before=aux)

        self._build_trust_card(parent)

    def _build_stepper(self, parent):
        """横排 ① Prepare → ② Check → ③ Fix（规格把竖排时间线换成横排）。"""
        F = self.F
        wrap = tk.Frame(parent, bg=theme.SURFACE)
        wrap.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        for c in range(3):
            wrap.columnconfigure(c, weight=1, uniform="step")
        self._step_circles: list = []
        self._step_labels: list = []
        for i, key in enumerate(("step1_title", "step2_title", "step3_title")):
            cell = tk.Frame(wrap, bg=theme.SURFACE)
            cell.grid(row=0, column=i, sticky="n")
            circ = StepCircle(cell, i + 1, font=F["F_STEP_N"], bg=theme.SURFACE)
            circ.pack()
            lbl = tk.Label(cell, text=self.tr(key), bg=theme.SURFACE,
                           fg=theme.TEXT_3, font=F["F_HELP"])
            lbl.pack(pady=(5, 0))
            self._step_circles.append(circ)
            self._step_labels.append(lbl)

    def _build_how_card(self, parent):
        """How it works 引导卡：跑过体检后收起，把版面让给结果。"""
        F = self.F
        self._how_card = RoundCard(parent, radius=12, fill=theme.SURFACE_SOFT,
                                   border=theme.INFO_BORDER, shadow=False,
                                   padx=16, pady=13)
        self._how_card.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        tk.Label(self._how_card.inner, text=self.tr("how_title"),
                 bg=theme.SURFACE_SOFT, fg=theme.PRIMARY, font=F["F_LABEL"],
                 anchor="w").pack(fill="x", pady=(0, 7))
        for name, text in i18n.t(self.lang, "how_steps"):
            row = tk.Frame(self._how_card.inner, bg=theme.SURFACE_SOFT)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=name, bg=theme.SURFACE_SOFT, fg=theme.PRIMARY,
                     font=F["F_LABEL"], width=9, anchor="w").pack(side="left")
            tk.Label(row, text="— " + text, bg=theme.SURFACE_SOFT,
                     fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                     justify="left").pack(side="left", fill="x", expand=True)

    def _build_trust_card(self, parent):
        F = self.F
        trust = RoundCard(parent, radius=12, fill=theme.SURFACE_SOFT,
                          border=theme.INFO_BORDER, shadow=False, padx=16, pady=13)
        trust.pack(side="bottom", fill="x", pady=(theme.SECTION_GAP, 0))
        tk.Label(trust.inner, text="✓   " + self.tr("trust_title"),
                 bg=theme.SURFACE_SOFT, fg=theme.PRIMARY, font=F["F_LABEL"],
                 anchor="w").pack(fill="x")
        body = tk.Label(trust.inner, text=self.tr("trust_body"),
                        bg=theme.SURFACE_SOFT, fg=theme.TEXT_2, font=F["F_HELP"],
                        anchor="w", justify="left")
        body.pack(fill="x", pady=(4, 0))
        _auto_wrap(body, trust.inner, 16)

    def _set_step(self, index: int):
        """0=Prepare / 1=Check / 2=Fix / 3=三步全部完成。"""
        self._step_index = max(0, min(int(index), 3))
        for i, circ in enumerate(self._step_circles):
            if i < self._step_index:
                circ.set_state("done")
                self._step_labels[i].config(fg=theme.SUCCESS)
            elif i == self._step_index:
                circ.set_state("active")
                self._step_labels[i].config(fg=theme.PRIMARY)
            else:
                circ.set_state("todo")
                self._step_labels[i].config(fg=theme.TEXT_3)

    def _sync_ui(self):
        """状态机收口：先算状态，再统一做一次重新测量。

        ``_relayout_all()`` 不能省 —— 这里会 pack/unpack ``_how_card``、显隐 ``_fix_btn``，
        都是"内容尺寸变了但不触发 <Configure>"的操作（见该方法 docstring）。
        """
        self._sync_ui_state()
        self._relayout_all()

    def _sync_ui_state(self):
        """主/次按钮与步进的状态机（规格 states 段）。

        ============  ====================================================
        状态          主按钮
        ============  ====================================================
        initial       必填未满足 → 置灰（并说明缺什么）
        ready         Check formatting
        checking      Checking…（置灰，禁止重复提交）
        results       有问题 → Fix N issues；无问题 → 保持 Check
        fixed         Review changes
        ============  ====================================================

        次按钮「Fix issues」只在 ready 态出现：出了结果时主按钮已经承担修正动作，
        再摆一个次按钮只会分散注意力（规格第 5 节第 8 条要求主按钮绝对主导）。
        """
        tr = self.tr
        paper = bool(self.docx_path.get()) and os.path.isfile(self.docx_path.get())
        # fixing 也算"已有结果"：修正过程中不该把引导卡又弹回来（会闪一下）
        results = self._phase in ("results", "fixed", "fixing")

        # How it works 只在还没跑出结果时显示，把版面让给结果
        try:
            if results:
                self._how_card.pack_forget()
            else:
                self._how_card.pack(fill="x", pady=(theme.SECTION_GAP, 0),
                                    before=self._status_row)
        except Exception:
            pass

        if self._busy:
            if not paper:
                self._set_step(0)
            elif self._phase == "fixing":
                self._set_step(2)          # 修正中高亮第 3 步（Fix），不是停在 Check
            else:
                self._set_step(1)
            self._check_btn.set_text(tr("btn_check_busy"))
            self._check_btn.set_enabled(False)
            self._fix_btn.set_visible(False)
            return

        if not paper:
            self._set_step(0)
            self._check_btn.set_text(tr("btn_check"))
            self._check_btn.set_enabled(False)
            self._check_btn.set_action(self._run)
            self._fix_btn.set_visible(False)
            if self._phase in ("idle", "error"):
                self._set_status(tr("state_missing_paper"), "idle",
                                 spec=("state_missing_paper", ()))
            return

        if self._phase == "fixed":
            self._set_step(3)
            self._check_btn.set_text(tr("btn_review"))
            self._check_btn.set_enabled(True)
            self._check_btn.set_action(self._review_changes)
            self._fix_btn.set_visible(False)
            return

        if self._phase == "results":
            self._set_step(2)
            issues = self._last_issues or 0
            if issues:
                self._check_btn.set_text(tr("btn_fix_n", issues))
                self._check_btn.set_action(self._run_fix)
            else:
                self._check_btn.set_text(tr("btn_check"))
                self._check_btn.set_action(self._run)
            self._check_btn.set_enabled(True)
            self._fix_btn.set_visible(False)
            return

        # ready（含出错后可重试）
        self._set_step(1)
        self._check_btn.set_text(tr("btn_check"))
        self._check_btn.set_enabled(True)
        self._check_btn.set_action(self._run)
        self._fix_btn.set_text(tr("btn_fix"))
        self._fix_btn.set_action(self._run_fix)
        self._fix_btn.set_visible(True)

    # ------------------------------------------------------------ 状态刷新
    def _refresh_rows(self):
        """刷新各输入源的状态显示（已选 / 未选）+ 触发界面状态机。"""
        tr = self.tr

        def _sel(lbl, value, none_key):
            if value:
                name = os.path.basename(value)
                if len(name) > 40:
                    name = name[:18] + " … " + name[-18:]
                lbl.config(text="●   " + name, fg=theme.PRIMARY)
            else:
                lbl.config(text=tr(none_key), fg=theme.TEXT_3)

        _sel(self._paper_name, self.docx_path.get(), "st_paper_none")
        _sel(self._tpl_name, self.school_template.get(), "st_tpl_none")
        _sel(self._latex_name, self.latex_template.get(), "st_latex_none")
        # 问卷的状态来自内存数据（不是文件路径），单独判
        if self.questionnaire_data:
            self._q_name.config(text=tr("st_q_set"), fg=theme.SUCCESS)
        else:
            self._q_name.config(text=tr("st_q_none"), fg=theme.TEXT_3)
        _sel(self._ai_name, self.ai_json.get(), "st_ai_none")

        # 选了真实存在的论文 → 从 idle 进入 ready。
        # results / fixed 由动作本身决定，这里绝不能覆盖（否则一刷新就把结果态抹掉）。
        cur = self.docx_path.get()
        # 兜底：换了文档就作废旧结论（换文件走 _on_input_changed，这里防其它入口漏掉）
        if self._checked_path and cur != self._checked_path:
            self._invalidate_results()
        chosen = bool(cur) and os.path.isfile(cur)
        if chosen and self._phase in ("idle", "error"):
            self._phase = "ready"
        self._sync_ui()

    def _invalidate_results(self):
        """作废上一次体检结论与修正产物记录。

        两件真事（Codex 2026-09-12 P1-2）：
        ① 换论文后主按钮仍写 "Fix N issues"，点下去直接改**新文件**；
        ② ``fixed`` 态下「Review changes」会打开**上一份**文件的目录。
        """
        self._checked_path = ""
        self._last_issues = None
        self._last_output = ""
        if self._phase in ("results", "fixed"):
            self._phase = "ready"
        self._set_report(self.tr("report_placeholder"), placeholder=True)

    def _set_status(self, text: str, kind: str = "idle", spec: tuple = None):
        """状态行（结果 / 提示都在这里）。kind: idle / run / ok / warn / err。

        ``spec=(词条key, 位置参数)`` 时，切语言重建后能按新语言**重新渲染**同一个状态；
        不传就原样留存（错误详情这类本来就是动态文本，译不了但也不能丢）。
        """
        self._status_spec = (spec, kind, text)
        colors = {"idle": theme.TEXT_3, "run": theme.PRIMARY, "ok": theme.SUCCESS,
                  "warn": theme.GOLD, "err": theme.ERROR}
        try:
            self.status_dot.config(fg=colors.get(kind, theme.TEXT_3))
            self.status_lbl.config(text=text)
        except Exception:
            pass
        self._relayout_all()

    def _set_bar(self, state: str, key: str = None):
        """状态栏（左下灯 + 右下文案）；state: idle / run / ok / err。"""
        self._bar_state = state
        if key is None:
            key = {"idle": "bar_idle", "run": "bar_checking", "ok": "bar_check_done",
                   "err": "bar_error"}.get(state, "bar_idle")
        colors = {"idle": theme.SUCCESS, "run": theme.PRIMARY, "ok": theme.SUCCESS,
                  "err": theme.ERROR}
        try:
            self.bar_dot.config(fg=colors.get(state, theme.SUCCESS))
            self.bar_right.config(text=self.tr(key))
        except Exception:
            pass

    def _set_report(self, text: str, placeholder: bool = False, summary: str = None,
                    summary_spec: tuple = None):
        """记录完整报告（供「保存报告」导出）+ 在界面只显示**一行精简摘要**。

        ``placeholder=True`` 表示这只是"还没跑体检"的引导语 —— 切换语言重建界面时
        要按新语言重新取词，而不是把旧语言的引导语原样搬过去（否则中文界面里会残留
        一段英文提示）。

        ``summary_spec=(词条key, 位置参数)``：摘要以「词条+参数」形式留存，切语言重建
        后按新语言**重新渲染**（否则界面会残留旧语言的摘要）。

        注：``_report_text``（真实体检报告）来自引擎、与界面语言无关；但 ``msg_no_change``
        与「修正预览」两处传入的本身就是界面文案，切语言后仍为旧语言 —— 影响仅限"导出这份
        提示文本"，可接受。

        v2.1.1：旧版把整篇报告塞进界面下方的 ScrolledText（height=7 / minsize=150），
        既把页面撑高、把页脚挤出可视区，也与「报告应当导出」的产品定位不符。
        现在界面只给摘要，完整报告走导出（右栏「保存报告」/ ``_save_report``）。
        """
        self._report_text = text or ""
        self._report_is_placeholder = placeholder
        if summary_spec is not None:
            self._report_summary_spec = summary_spec
            summary = self._render_summary(summary_spec)
        else:
            self._report_summary_spec = None
            if summary is None:
                summary = text or ""
        self._report_summary = summary
        try:
            self.summary_lbl.config(text=summary)
        except Exception:
            pass
        self._relayout_all()

    def _render_summary(self, spec: tuple) -> str:
        """按当前语言渲染摘要（spec = (词条 key, 位置参数)）。"""
        try:
            key, args = spec
            return self.tr(key, *(args or ()))
        except Exception:
            return self.tr("report_summary_generic")

    def _summary_counts_spec(self, s: dict) -> tuple:
        """report.summary（pass/info/warn/fail）→ 明细行 spec（便于多语言重渲染）。"""
        return ("state_counts_hint",
                (s.get("pass", 0), s.get("info", 0),
                 s.get("warn", 0), s.get("fail", 0)))

    def _set_busy(self, busy: bool):
        """忙态统一交给状态机（RoundButton 自绘，没有 ttk 的 state 选项）。"""
        self._busy = bool(busy)
        self._sync_ui()
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
    def _on_input_changed(self):
        """任何会影响「体检结论」的输入变了 → 作废旧结论 + 刷新界面。

        统一入口，避免逐个选择器漏掉（Codex 2026-09-12 P1-2）。
        """
        if self._checked_path:
            self._invalidate_results()
        self._refresh_rows()

    def _pick_docx(self):
        p = filedialog.askopenfilename(title=self.tr("row_paper_title"),
                                       filetypes=[("Word", "*.docx"), ("All files", "*.*")])
        if p:
            self.docx_path.set(p)
            self._on_input_changed()

    def _pick_template(self):
        p = filedialog.askopenfilename(title=self.tr("row_template_title"),
                                       filetypes=[("Word", "*.docx"), ("All files", "*.*")])
        if p:
            self.school_template.set(p)
            self._on_input_changed()

    def _pick_latex(self):
        p = filedialog.askopenfilename(
            title=self.tr("row_latex_title"),
            filetypes=[("LaTeX", "*.tex *.cls *.sty"), ("All files", "*.*")])
        if p:
            self.latex_template.set(p)
            self._on_input_changed()

    def _pick_ai_json(self):
        p = filedialog.askopenfilename(title=self.tr("btn_ai_json"),
                                       filetypes=[("JSON", "*.json"), ("All files", "*.*")])
        if p:
            self.ai_json.set(p)
            self._on_input_changed()

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
            self._on_input_changed()

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
        """运行格式体检（免费不限次）：出结果后主按钮自动变成 Fix N issues。"""
        if self._busy:
            return
        tr = self.tr
        docx = self.docx_path.get()
        if not docx or not os.path.isfile(docx):
            # 规格 states.error：短的行内提示，不弹模态框
            self._phase = "error"
            self._sync_ui()
            self._set_status(tr("state_missing_paper"), "err",
                             spec=("state_missing_paper", ()))
            return
        self._phase = "ready"
        self._set_busy(True)
        self._set_status(tr("status_checking"), "run")
        self._set_bar("run")
        try:
            target = self._build_target()
            prof, findings = checker.run_check(docx, target)
            rep = report_mod.build_report(docx, target, prof, findings)

            # 结果整理同样放在 try 内 + finally 收尾：这一段若抛异常而逃出去，
            # `_busy` 会永远停在 True（RoundButton 的异常兜底是静默的），
            # 主按钮永久显示 "Checking…"、次按钮永久隐藏 —— 只能重启进程。
            s = rep.summary or {}
            issues = int(s.get("warn", 0)) + int(s.get("fail", 0))
            self._last_issues = issues
            self._checked_path = docx
            self._phase = "results"
            self._set_report(rep.markdown, summary_spec=self._summary_counts_spec(s))
            if issues == 0:
                self._set_status(tr("state_clean"), "ok", spec=("state_clean", ()))
            elif issues == 1:
                self._set_status(tr("state_issue_one"), "warn",
                                 spec=("state_issue_one", ()))
            else:
                self._set_status(tr("state_issues", issues), "warn",
                                 spec=("state_issues", (issues,)))
            self._set_bar("ok")
        except DocxReadError as e:
            self._fail(tr("msg_read_error_title"), str(e))
        except (ValueError, FileNotFoundError, OSError) as e:
            self._fail(tr("msg_check_error_title"), str(e))
        except Exception as e:  # noqa: BLE001  兜底，界面绝不能因未知异常僵死
            self._fail(tr("msg_check_error_unexpected"), str(e))
        finally:
            self._set_busy(False)

    def _fail(self, title: str, detail: str):
        """行内报错 —— 规格 states.error：短提示，绝不整屏红（不弹模态框）。

        完整 detail 同时写进 ``_report_text``，这样「保存报告」能拿到全部信息
        （只留 150 字单行的话，用户既看不到全貌也复制不了）。
        """
        self._phase = "error"
        detail = (detail or "").strip()
        if detail:
            self._set_report("%s\n\n%s" % (title, detail))
        self._set_busy(False)
        first = detail.splitlines() if detail else []
        msg = first[0] if first else ""
        if len(msg) > 150:
            msg = msg[:147] + "…"
        self._set_status("%s — %s" % (title, msg) if msg else title, "err")
        self._set_bar("err")
        # 错误可能发生在滚动区下方，把视口拉回顶部保证用户看得见
        try:
            self._main_scroll.scroll_to_top()
        except Exception:
            pass

    def _review_changes(self):
        """「Review changes」：打开修正结果所在目录，让用户直接看改完的文件。

        没有落盘记录时（例如刚被语言重建清过状态）退回「保存报告」，
        保证这个按钮在任何情况下都有意义。
        """
        out = self._last_output or ""
        if out and os.path.isfile(out):
            try:
                folder = os.path.dirname(out)
                if sys.platform.startswith("win"):
                    os.startfile(folder)          # noqa: S606  Windows 专有
                elif sys.platform == "darwin":
                    import subprocess
                    subprocess.Popen(["open", folder])
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", folder])
                self._set_status(self.tr("state_fixed"), "ok")
                return
            except Exception:
                pass
        self._save_report()

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
            self._phase = "error"
            self._sync_ui()
            self._set_status(tr("state_missing_paper"), "err",
                             spec=("state_missing_paper", ()))
            return

        self._phase = "fixing"
        self._set_busy(True)
        self._set_status(tr("status_fixing"), "run")
        self._set_bar("run", "bar_fixing")
        # 整段「算改动 → 授权裁决 → 预览」都在 try 里，由 finally 收忙态：
        # 任何一处抛异常（例如 gate 结构异常、文案格式化失败）都不能让
        # `_busy` 卡在 True —— 那会让主按钮永久停在 "Checking…"。
        gate_no_go = ""
        preview_text = ""
        try:
            target = self._build_target()
            changes = compute_changes(docx, target)

            if not changes:
                # 已合规：不触发门禁、也不弹模态框（规格 states 要行内提示）
                self._phase = "results"
                self._last_issues = 0
                self._set_report(tr("msg_no_change"),
                                 summary_spec=("msg_no_change", ()))
                self._set_status(tr("msg_no_change"), "ok",
                                 spec=("msg_no_change", ()))
                self._set_bar("ok")
                return

            # 真要动手改，才做授权裁决（首次免费、之后需激活码）
            gate = lic.require_fix_entitlement()
            if not gate.get("allowed"):
                self._phase = "ready"
                gate_no_go = gate.get("message") or tr("msg_need_activation_title")
                self._set_status(gate_no_go, "err")
                self._set_bar("err")
                return

            preview_text = (
                tr("preview_header") + "\n"
                + "\n".join("  · " + line for line in changes)
                + "\n\n" + tr("preview_trial", gate.get("message", "")))
            self._set_report(preview_text,
                             summary_spec=("report_summary_fix", (len(changes),)))
        except DocxReadError as e:
            self._fail(tr("msg_read_error_title"), str(e))
            return
        except (ValueError, FileNotFoundError, OSError) as e:
            self._fail(tr("msg_fix_preview_error_title"), str(e))
            return
        except Exception as e:  # noqa: BLE001
            self._fail(tr("msg_fix_error_unexpected"), str(e))
            return
        finally:
            self._set_busy(False)

        # 授权不足：这是购买决策、不是报错，留在模态框里保证用户一定看见
        # （必须等忙态解除后再弹，否则模态框会挡住还没刷新的按钮状态）
        if gate_no_go:
            messagebox.showwarning(tr("msg_need_activation_title"), gate_no_go)
            return

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
            self._last_output = out
            self._phase = "fixed"
        except DocxReadError as e:
            self._fail(tr("msg_read_error_title"), str(e))
            return
        except (ValueError, OSError) as e:
            self._fail(tr("msg_fix_error_title"), str(e))
            return
        finally:
            self._set_busy(False)

        self._sync_ui()                      # 主按钮 → Review changes，三步全部打勾
        self._set_status(tr("state_fixed"), "ok", spec=("state_fixed", ()))
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
