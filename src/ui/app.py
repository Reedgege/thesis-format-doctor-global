"""Thesis Format Doctor Global —— tkinter 桌面 GUI（离线，论文不出本机）。

界面：象牙白纸底 + 青蓝主色。顶栏（品牌 / 标语 / 隐私声明 / 语言·关于·帮助）
→ 两栏卡片（左「Your Documents」/ 右「Review & Fix」）→ 页脚（邮箱 / 官网 / 标语）
→ 状态栏。

v2.2.0 按设计规格书重做视觉层（老板 2026-09-12 交办），**功能一行未动**：

- 视觉层拆成三块，各管一摊：
  - ``theme.py``   —— 颜色 / 字号 / 间距令牌（唯一来源，改色只改这里）
  - ``widgets.py`` —— 圆角卡片 / 圆角按钮 / 步进圆点（tkinter 没有圆角与投影，
    这三个是 Canvas 自绘）
  - 本文件         —— 布局骨架 + 交互状态机
- 三步走（Prepare / Check / Fix）收成**一条极简进度条**（不再是横排步进圆点 / 竖排时间线）；
  1/3 计数、章节角标（Ⅰ·SELECT / Ⅱ·PROCESS）按规格删除。
- 高级设置（学校模板 / LaTeX / 问卷 / 配置导入）统一收进 **Advanced options 折叠区**，
  主界面只留必填的「引用规范 + 上传 Word」，落实渐进式披露。
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
import ctypes
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
from .modal_shell import ModalShell
from .widgets import (BrandMark, IconBadge, ProgressBar, RoundButton, RoundCard,
                      ScrollArea, draw_icon)

# --- 原生拖拽上传（仅 Windows）：把文档拖进窗口即选中 ---------------------------
if sys.platform == "win32":
    class _FileDropTarget:
        """用 ctypes 接管 Tk 顶层窗口的 WM_DROPFILES，实现「拖文件进窗口即选中」。

        非 Windows（macOS / Linux）Tk 没有这套原生 API，降级为「点选」——投放区仍是
        视觉焦点，只是少了拖拽手势。跨平台拖拽需 tkinterdnd2 这类外部依赖，
        当前不引入（保持零第三方 GUI 依赖）。
        """
        WM_DROPFILES = 0x0233
        GWL_WNDPROC = -4

        def __init__(self, hwnd, callback):
            self._user32 = ctypes.windll.user32
            self._shell32 = ctypes.windll.shell32
            self._hwnd = hwnd
            self._callback = callback
            self._shell32.DragAcceptFiles(hwnd, True)
            # 子类化窗口过程：自己只拦 WM_DROPFILES，其余消息原样转发给 Tk。
            # 必须保留引用（self._wndproc / self._old），否则回调被 GC 后崩溃。
            self._wndproc = wintypes.WNDPROC(self._proc)
            self._old = self._user32.SetWindowLongPtrW(
                hwnd, self.GWL_WNDPROC, self._wndproc)

        def _proc(self, hwnd, msg, wparam, lparam):
            if msg == self.WM_DROPFILES:
                try:
                    self._emit(wparam)
                except Exception:
                    pass
                return 0
            return self._user32.CallWindowProcW(
                self._old, hwnd, msg, wparam, lparam)

        def _emit(self, hdrop):
            count = self._shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
            buf = ctypes.create_unicode_buffer(2048)
            paths = []
            for i in range(count):
                self._shell32.DragQueryFileW(hdrop, i, buf, 2048)
                paths.append(buf.value)
            self._shell32.DragFinish(hdrop)
            if paths:
                self._callback(paths)
else:
    _FileDropTarget = None

# ---------------------------------------------------------------------------
# 配色别名 —— 新令牌在 theme.py，这里保留旧名字是为了**没重写的弹窗**（关于 /
# 帮助 / 问卷 / 激活）能一行不改地跟着换成新配色。
# 主色从「黛蓝 + 朱砂」改为「青蓝」：朱砂（砖红）按规格**不再作主操作色**，
# 只降级为点缀（GOLD），红色仅保留给真实错误。
# ---------------------------------------------------------------------------
PAPER = theme.BG              # 页面底（象牙纸）
PANEL = theme.SURFACE_SOFT    # 柔和面（信息框 / 状态栏）
CARD = theme.SURFACE          # 卡片面
INK = theme.TEXT              # 主文字
BODY = theme.TEXT             # 正文（弹窗用）
MUTED = theme.TEXT_2          # 次要文字
LINE = theme.BORDER           # 边框 / 细线
ACCENT = theme.PRIMARY        # 主色（青蓝）
CINNABAR = theme.GOLD         # 旧「朱砂」→ 降级为点缀金（不再是主按钮色）
CINNABAR_D = "#A5894B"        # 点缀金的 active 态
OKC = theme.SUCCESS           # 成功 / 已完成
ERRC = theme.ERROR            # **仅**真实错误

_BASE_W = 1180         # 设计基准宽度（与 fonts.BASE_WIDTH 一致）

_HERE = os.path.dirname(os.path.abspath(__file__))
_TEMPLATE_PATH = os.path.join(_HERE, "..", "data", "ai_questionnaire_template.json")

_FIELD_RANGE = FIELD_RANGES

# 引用规范下拉的展示名：引擎 key 不变（build_target 仍收 APA），
# 仅把界面上的 APA 显示成设计稿的 "APA 7th edition"。
_SPEC_DISPLAY = {"APA": "APA 7th edition"}


def _spec_display(key: str) -> str:
    return _SPEC_DISPLAY.get(key, key)

# 字段行左侧的小圆徽标用的线描图标（设计稿每个字段前都有一枚，见 widgets.draw_icon）
_FIELD_ICONS = {
    "row_spec_title": "quote",
    "row_paper_title": "doc",
    "row_template_title": "bank",
    "row_latex_title": "code",
}


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

    state = {"wl": None, "busy": False}

    def _refresh_card(w):
        """换行一变，label 的请求高度就变；向上找最近的 RoundCard 让它重新测量。

        卡片内容区（``RoundCard.inner``）的尺寸是钉死的，label 变高**不会**触发父级
        重排 —— 不主动喊一声，宽窗口拖窄后信任卡/说明文字的第二行会被圆角矩形裁掉。
        """
        try:
            while w is not None:
                if isinstance(w, RoundCard):
                    w.refresh()
                    return
                w = w.master
        except Exception:
            pass

    def _apply(_e=None):
        if state["busy"]:
            return
        try:
            w = container.winfo_width() - 2 * padx
            if w <= 80 or w == state["wl"]:
                return
            state["wl"] = w
            label.configure(wraplength=w)
        except Exception:
            return
        # busy 防重入：refresh() 会再次触发 <Configure>，没有它就可能来回重排。
        state["busy"] = True
        try:
            _refresh_card(label)
        finally:
            state["busy"] = False

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
        # 先内后外：嵌套卡片没有自己的布局入口，父卡量到的会是它们的旧请求高度。
        for card in list(getattr(self, "_extra_cards", [])):
            if card is not None:
                try:
                    card.refresh()
                except Exception:
                    pass
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
        """ttk 样式 —— 主操作色统一为青蓝（砖红不再作主操作色，规格第 5 节第 4 条）。

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
        # 嵌套的圆角卡片（Advanced 条 / 信任面板）要跟着一起重测量，
        # 否则父卡按旧高度画圆角矩形、新内容被裁掉（见 _relayout_all）。
        self._extra_cards: list = []

        self._build_backdrop()
        self._build_footer()
        self._build_statusbar()
        self._build_cta_band()
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
        main.columnconfigure(0, weight=1, uniform="half", minsize=380)
        main.columnconfigure(1, weight=1, uniform="half", minsize=380)
        main.rowconfigure(0, weight=1)
        self._main_frame = main
        # 两栏宽度只由窗口决定（见 _sync_columns），否则卡片里的自动换行会和 grid
        # 的列宽分配互相追着跑。
        main.bind("<Configure>", self._sync_columns, add="+")

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
        self._enable_drag_drop()

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
    def _sync_columns(self, e=None):
        """把两栏列宽钉成"窗口算出来的半宽"。

        卡片里的自动换行标签按格宽重排 → 内容请求宽度跟着变。若让 grid 按内容请求
        分配列宽，请求更大的那一列会先吃掉空间、另一列被压窄，换行结果再反过来改
        请求 —— 两个值互推，``update_idletasks`` 永远收敛不了（实测卡死）。
        列宽一旦只跟窗口宽度走，内容请求就再也顶不动布局了。
        """
        try:
            main = self._main_frame
            w = int(e.width) if e is not None else int(main.winfo_width())
            if w <= 2:
                return
            half = max(300, (w - theme.CARD_GAP) // 2)
            for col in (0, 1):
                main.columnconfigure(col, minsize=half)
        except Exception:
            pass

    def _build_backdrop(self):
        """最底层背景：整窗纯色铺满（v2.3.4 起）。

        视觉迭代到 v2.3.4：去掉右上角装饰底图（backdrop.png），改为整窗纯色
        ``theme.BG``。纯色背景更干净、零接缝、零资源依赖，也彻底消除了旧版
        「天际线横带」「底色与卡片对不上」等渲染问题。

        这块画布仍必须存在，承担两件事：
          ① 作为整窗背景层（``place`` 铺满、最先创建 = 最底层）；
          ② 顶栏/页眉的控件用 ``create_window`` 挂在这块画布上（见 ``_place_header``），
             所以画布本身要继续保留，只是不再画任何图片。
        """
        self._backdrop = tk.Canvas(self.root, bg=theme.BG,
                                   highlightthickness=0, bd=0, width=1, height=1)
        # place：不参与 pack 的空间分配，纯铺底层；先创建 = 层叠最下。
        self._backdrop.place(x=0, y=0, relwidth=1, relheight=1)

    def _place_backdrop(self):
        """纯色背景无需重定位（整窗铺满，relwidth/relheight=1 自动跟随窗口）。

        v2.3.4 之前这里负责把装饰底图重新贴到右上角；改为纯色后留作占位，
        保持 ``_on_resize`` 调用的兼容性，不再做任何事。
        """
        return

    def _top_gap(self) -> int:
        """主体上方该留出的高度：顶栏实际高度 + HEADER_GAP。

        顶栏画在背景画布上（不占 pack 空间），这段高度必须显式留出来，否则主体
        容器会压到顶栏上。v2.3.4 去掉装饰底图后，不再需要为「天际线下沿」预留额外高度。
        """
        return theme.HEADER_PAD_Y + self._header_height() + theme.HEADER_GAP

    def _header_height(self) -> int:
        try:
            return max(self._hdr_brand.winfo_reqheight(),
                       self._hdr_privacy.winfo_reqheight(),
                       self._hdr_nav.winfo_reqheight())
        except Exception:
            return 56

    def _place_header(self):
        """顶栏定位：品牌贴左、导航贴右、隐私声明跟在品牌右侧，三者都不吃 pack 空间。

        隐私声明是「有位置就显示、挤到了就整块收起」的可选件：窄窗口下它跟右上导航
        相撞时直接隐藏，绝不允许压字（规格：隐私提示必须安静，不能抢版面）。
        """
        try:
            w = self.root.winfo_width()
            if w <= 1:
                return
            cv = self._backdrop
            pad = theme.PAGE_PAD
            bh = self._hdr_brand.winfo_reqheight()
            nh = self._hdr_nav.winfo_reqheight()
            cv.coords(self._hdr_brand_win, pad, theme.HEADER_PAD_Y)
            nav_y = theme.HEADER_PAD_Y + max(0, (bh - nh) // 2)
            cv.coords(self._hdr_nav_win, w - pad, nav_y)

            px = pad + self._hdr_brand.winfo_reqwidth() + 34
            ph = self._hdr_privacy.winfo_reqheight()
            pw = self._hdr_privacy.winfo_reqwidth()
            nav_x = w - pad - self._hdr_nav.winfo_reqwidth()
            fits = px + pw + 16 <= nav_x
            state = "normal" if fits else "hidden"
            cv.coords(self._hdr_privacy_win, px,
                      theme.HEADER_PAD_Y + max(0, (bh - ph) // 2))
            cv.itemconfigure(self._hdr_privacy_win, state=state)
            cv.itemconfigure(self._hdr_sep_item, state=state)
            cv.coords(self._hdr_sep_item, px - 18, theme.HEADER_PAD_Y + 2,
                      px - 18, theme.HEADER_PAD_Y + bh - 2)
        except Exception:
            pass

    def _build_header(self):
        """顶栏：品牌 + 标语 + 隐私声明（左）／语言 ▾ · About · Help（右）。

        与前版的差别（都是为背景服务）：
          ① 顶栏仍画在背景画布上（``create_window``，不占 pack 空间）—— 背景素材的
             天际线正好压在顶栏这条带上，整条不透明框架盖住就等于白放了一张图；
          ② **删掉顶栏下的通栏细线**：参考稿里没有这条线，留着会横穿天际线；
          ③ 按设计稿补上**品牌 mark**（蓝底圆角方块 + 白色学士帽，``widgets.BrandMark``）
             与隐私声明前的盾牌小徽标；语言选择器改成「当前语言 ▾」。
        """
        F = self.F
        cv = self._backdrop

        # —— 品牌：mark + 标题（单行，左缘 = 页面留白）——
        brand = tk.Frame(cv, bg=theme.BG)
        BrandMark(brand, size=theme.MARK_SIZE, bg=theme.BG).pack(side="left")
        tk.Label(brand, text=self.tr("app_title"), bg=theme.BG, fg=theme.PRIMARY,
                 font=F["F_BRAND"], anchor="w").pack(side="left", padx=(13, 0))

        # —— 隐私声明：单行，保持「安静」（不再压第二行小字）——
        privacy = tk.Frame(cv, bg=theme.BG)
        IconBadge(privacy, "shield", size=theme.BADGE_SIZE, bg=theme.BG).pack(side="left")
        tk.Label(privacy, text=self.tr("privacy_line"), bg=theme.BG, fg=theme.TEXT_2,
                 font=F["F_SMALL_B"]).pack(side="left", padx=(8, 0))

        # 右侧导航：自右向左 pack
        nav = tk.Frame(cv, bg=theme.BG)
        self._link(nav, self.tr("link_help"),
                   self._help_dialog).pack(side="right")
        self._nav_sep(nav)
        self._link(nav, self.tr("link_about"),
                   self._about_dialog).pack(side="right")
        self._nav_sep(nav)
        self._link(nav, self.tr("lang_button"),
                   lambda: self._set_lang(i18n.other_lang(self.lang)),
                   fg=theme.TEXT).pack(side="right")

        self._hdr_brand, self._hdr_nav, self._hdr_privacy = brand, nav, privacy
        self._hdr_brand_win = cv.create_window(theme.PAGE_PAD, theme.HEADER_PAD_Y,
                                               window=brand, anchor="nw")
        self._hdr_nav_win = cv.create_window(0, theme.HEADER_PAD_Y,
                                             window=nav, anchor="ne")
        self._hdr_privacy_win = cv.create_window(0, theme.HEADER_PAD_Y,
                                                 window=privacy, anchor="nw")
        self._hdr_sep_item = cv.create_line(0, 0, 0, 0, fill=theme.BORDER)
        self._place_header()

    def _nav_sep(self, parent):
        tk.Label(parent, text="  |  ", bg=theme.BG, fg=theme.BORDER,
                 font=self.F["F_HELP"]).pack(side="right")

    def _build_footer(self):
        """页脚：左邮箱 / 右官网 —— 简化到两项，保持安静小字。

        规格明确「页脚只此二项」（品牌口径只留官网与邮箱，不出现微信 / 公众号 /
        小程序）。去掉了标语行与分隔线，进一步降噪。
        """
        F = self.F
        footer = tk.Frame(self.root, bg=theme.BG)
        footer.pack(side="bottom", fill="x", padx=theme.PAGE_PAD,
                    pady=(0, theme.FOOTER_PAD_Y))
        mail = tk.Label(footer, text="✉  " + self.tr("footer_email"), bg=theme.BG,
                        fg=theme.TEXT_2, font=F["F_FOOT"], cursor="hand2")
        mail.pack(side="left")
        mail.bind("<Button-1>",
                  lambda e: webbrowser.open("mailto:" + i18n.BRAND_EMAIL))
        site = tk.Label(footer, text="◎   " + self.tr("footer_site"), bg=theme.BG,
                        fg=theme.TEXT_2, font=F["F_FOOT"], cursor="hand2")
        site.pack(side="right")
        site.bind("<Button-1>", lambda e: webbrowser.open(i18n.BRAND_SITE_URL))

    def _build_statusbar(self):
        F = self.F
        tk.Frame(self.root, bg=theme.BORDER, height=1).pack(
            fill="x", side="bottom", padx=theme.PAGE_PAD)
        statusbar = tk.Frame(self.root, bg=theme.BG)
        statusbar.pack(side="bottom", fill="x", padx=theme.PAGE_PAD, pady=(2, 4))
        self.bar_dot = tk.Label(statusbar, text="●", bg=theme.BG, fg=OKC,
                                font=F["F_FOOT"])
        self.bar_dot.pack(side="left", padx=(0, 6))
        tk.Label(statusbar, text=self.tr("bar_left", self.version), bg=theme.BG,
                 fg=theme.TEXT_3, font=F["F_FOOT"]).pack(side="left")
        self.bar_right = tk.Label(statusbar, text="", bg=theme.BG, fg=theme.TEXT_3,
                                  font=F["F_FOOT"])
        self.bar_right.pack(side="right")

    # ------------------------------------------------------------ 底部 CTA 条
    def _build_cta_band(self):
        """底部辅助动作条：保存报告 / 导出 AI 模板。

        重设计（2026-09-26）：主/次 CTA 已移到左卡输入区底部（紧贴操作区），
        底部只保留两个次要的常驻动作，避免主按钮离输入区太远。
        """
        F = self.F
        band = tk.Frame(self.root, bg=theme.BG)
        band.pack(side="bottom", fill="x", padx=theme.PAGE_PAD, pady=(8, 10))
        tk.Frame(band, bg=theme.HAIRLINE, height=1).pack(
            fill="x", pady=(0, 10))
        self._aux_row = tk.Frame(band, bg=theme.BG)
        self._aux_row.pack(fill="x")
        for text, cmd in ((self.tr("btn_save_report"), self._save_report),
                          (self.tr("btn_export_ai"), self._export_template)):
            ttk.Button(self._aux_row, text=text, style="Ghost.TButton",
                       command=cmd).pack(side="left", padx=(0, 6))

    # ------------------------------------------------------------ 左栏
    def _build_left(self, parent):
        """左卡 Your Documents：以「上传论文」为绝对视觉焦点的任务启动器。

        重设计（2026-09-26）：
          ① 论文投放区放到最上方、面积放大，作为整个页面的视觉锚点；
          ② 主操作按钮（Check formatting / Fix issues）紧贴输入区底部，
             不再甩到页面最下方；
          ③ 引用规范、Advanced options 退居二线，形成「先上传 → 再调设置 → 最后执行」
             的自然任务流；
          ④ 投放区支持鼠标悬停反馈（原生文件拖拽需额外依赖，当前保留视觉暗示）。
        """
        F = self.F

        head = tk.Frame(parent, bg=theme.SURFACE)
        head.pack(fill="x")
        IconBadge(head, "doc", size=theme.BADGE_SIZE + 8,
                  bg=theme.SURFACE).pack(side="left")
        htxt = tk.Frame(head, bg=theme.SURFACE)
        htxt.pack(side="left", padx=(10, 0), fill="x", expand=True)
        tk.Label(htxt, text=self.tr("left_title"), bg=theme.SURFACE,
                 fg=theme.TEXT, font=F["F_CARD_HDR"], anchor="w").pack(fill="x")
        desc = tk.Label(htxt, text=self.tr("left_desc"), bg=theme.SURFACE,
                        fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                        justify="left")
        desc.pack(fill="x", pady=(2, 0))
        _auto_wrap(desc, htxt, 0)

        self._spacer(parent)

        # ① 论文（必选）—— 页面视觉焦点：大虚线投放区
        self._field_label(parent, "row_paper_title", required=True)
        self._helper(parent, "row_paper_helper")
        drop = RoundCard(parent, radius=14, fill=theme.PRIMARY_SOFT,
                         border=theme.PRIMARY, shadow=False, dashed=True,
                         padx=18, pady=26, min_height=132)
        self._extra_cards.append(drop)
        drop.pack(fill="x", pady=(4, 0))
        self._paper_drop = drop

        # 居中内容列
        vcol = tk.Frame(drop.inner, bg=theme.PRIMARY_SOFT)
        vcol.pack(expand=True)
        cloud = tk.Canvas(vcol, width=44, height=40, bg=theme.PRIMARY_SOFT,
                          highlightthickness=0, bd=0)
        cloud.pack(pady=(0, 10))
        draw_icon(cloud, "cloud", 22, 20, 40, theme.PRIMARY, 2)

        self._paper_name = tk.Label(vcol, text=self.tr("btn_select_doc"),
                                    bg=theme.PRIMARY_SOFT, fg=theme.PRIMARY,
                                    font=F["F_BODY"], cursor="hand2", anchor="center")
        self._paper_name.pack(fill="x")
        self._drop_hint = tk.Label(vcol, text=self.tr("drop_hint"),
                                   bg=theme.PRIMARY_SOFT, fg=theme.TEXT_2,
                                   font=F["F_HELP"], cursor="hand2", anchor="center")
        self._drop_hint.pack(fill="x", pady=(5, 0))
        self._paper_sub = tk.Label(vcol, text=self.tr("paper_support"),
                                   bg=theme.PRIMARY_SOFT, fg=theme.TEXT_3,
                                   font=F["F_HELP"], cursor="hand2", anchor="center")
        self._paper_sub.pack(fill="x", pady=(3, 0))
        for w in (drop.inner, vcol, self._paper_name, self._drop_hint,
                  self._paper_sub, cloud):
            w.bind("<Button-1>", lambda e: self._pick_docx())
            w.bind("<Enter>", lambda e: self._set_drop_hover(True))
            w.bind("<Leave>", lambda e: self._set_drop_hover(False))

        self._spacer(parent)
        self._hairline(parent)

        # ② 引用规范（必选）—— 上传后的次要设置
        self._build_field_spec(parent)

        self._spacer(parent)
        self._hairline(parent)

        # ③ 高级选项折叠区
        self._build_advanced(parent)

        # ④ 主 / 次 CTA 紧贴输入区
        self._spacer(parent)
        self._check_btn = RoundButton(parent, self.tr("btn_check"), self._run,
                                      style="primary", font=F["F_BTN"],
                                      height=50, icon="search")
        self._check_btn.pack(fill="x", pady=(0, 8))
        self._fix_btn = RoundButton(parent, self.tr("btn_fix"), self._run_fix,
                                    style="secondary", font=F["F_BTN_S"],
                                    height=42, icon="wrench")
        self._fix_btn.pack(fill="x")

    def _hairline(self, parent):
        """字段之间的极细分隔线。

        只给**上方**留白：下方的间距由紧跟着的 ``_field_label`` 提供。
        旧写法两边各给一个 FIELD_GAP，两个字段之间白吃 30+px 高度，左卡一下就撑到 700+。
        """
        tk.Frame(parent, bg=theme.HAIRLINE, height=1).pack(
            fill="x", pady=(theme.HAIRLINE_PAD, 0))

    def _spacer(self, parent):
        """可伸缩留白：窗口拉高时把多出来的高度摊到各区块之间。

        设计稿是"大留白"版式 —— 1440/1920 下卡片内容应当被撑开、区块间距变大，
        而不是在卡片底部空一大块。``height=0`` 的 Frame 对自然高度零贡献
        （1280×720 下布局一个像素都不变），只在容器被拉高时参与分空间。
        """
        try:
            bg = parent.cget("bg")
        except Exception:
            bg = theme.SURFACE
        tk.Frame(parent, bg=bg, height=0).pack(fill="x", expand=True)

    def _field_label(self, parent, key: str, required: bool = False):
        """字段标签行：小圆徽标 + 标签 + Required/Optional 徽标（+ 可选右侧控件）。

        规格第 5 节第 5 条：必选/可选标签要**小、中性偏蓝、不吓人** ——
        旧版用朱砂红，看着像报错，已改。方法返回这一行，调用方可以往右边挂控件
        （Citation style 的下拉就是这么对齐到同一行的）。
        """
        F = self.F
        row = tk.Frame(parent, bg=theme.SURFACE)
        row.pack(fill="x", pady=(theme.FIELD_GAP, 0))
        if _FIELD_ICONS.get(key):
            IconBadge(row, _FIELD_ICONS[key], size=theme.BADGE_SIZE,
                      bg=theme.SURFACE).pack(side="left")
        tk.Label(row, text=self.tr(key), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_LABEL"]).pack(side="left", padx=(8, 0))
        badge_text = self.tr("mark_required" if required else "mark_optional")
        # 必选/可选徽标走中性灰（v2.3.4 视觉收敛：不再用主色，避免满屏青蓝）
        badge_color = theme.SURFACE_SOFT
        badge_fg = theme.TEXT_2 if required else theme.TEXT_3
        tk.Label(row,
                 text="  " + badge_text + "  ",
                 bg=badge_color, fg=badge_fg,
                 font=F["F_HELP"], padx=5).pack(side="left", padx=(6, 0))
        return row

    def _helper(self, parent, key: str):
        F = self.F
        lbl = tk.Label(parent, text=self.tr(key), bg=theme.SURFACE,
                       fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                       justify="left")
        lbl.pack(fill="x", pady=(1, theme.ROW_GAP))
        _auto_wrap(lbl, parent, theme.CARD_PAD_X)
        return lbl

    def _set_drop_hover(self, active: bool):
        """论文投放区悬停反馈：变文案、变光标，暗示可点击/可拖拽。"""
        if not getattr(self, "_paper_drop", None):
            return
        try:
            if active:
                self._drop_hint.config(text=self.tr("drop_active"),
                                         fg=theme.PRIMARY)
                self._paper_drop.inner.config(cursor="hand2")
            else:
                self._drop_hint.config(text=self.tr("drop_hint"),
                                         fg=theme.TEXT_2)
                self._paper_drop.inner.config(cursor="")
        except Exception:
            pass

    def _build_field_spec(self, parent):
        """引用规范（必选）：下拉框 + 一行「Other 用法」轻提示。"""
        F = self.F
        row = self._field_label(parent, "row_spec_title", required=True)
        # 设计稿里 Citation style 的下拉与标签**同行**右对齐（省一整行高度）
        # 设计稿把默认规范显示成 "APA 7th edition"，但引擎 key 仍是 APA
        # （build_target 只认 SPEC_ORDER 里的 key），这里只做展示层映射。
        self._spec_disp = tk.StringVar(value=_spec_display(self.spec_key.get()))
        self._spec_cb = ttk.Combobox(row, textvariable=self._spec_disp,
                                     values=[_spec_display(k) for k in SPEC_ORDER],
                                     state="readonly", font=F["F_BODY"], width=18)
        self._spec_cb.pack(side="right")
        self._spec_cb.bind("<<ComboboxSelected>>", self._on_spec_pick)
        self._helper(parent, "row_spec_desc")
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

    def _on_spec_pick(self, _e=None):
        """下拉展示名 → 引擎 key（选 APA 7th edition 仍写回 APA）。"""
        try:
            disp = self._spec_disp.get()
        except Exception:
            return
        for k in SPEC_ORDER:
            if _spec_display(k) == disp:
                self.spec_key.set(k)
                return

    def _select_field(self, parent, placeholder_key: str, cmd):
        """「下拉框外观 + 打开文件行为」的控件。

        规格画的是 ``Choose a template ▾`` 这样的下拉；我们的真实动作是打开文件
        对话框（本机离线读模板），所以外观照规格做、行为照实际来：整行可点。
        """
        F = self.F
        box = tk.Frame(parent, bg=theme.BG, highlightthickness=1,
                       highlightbackground=theme.FIELD_BORDER,
                       highlightcolor=theme.PRIMARY)
        box.pack(fill="x")
        inner = tk.Frame(box, bg=theme.SURFACE)
        inner.pack(fill="x", padx=10, pady=3)
        lbl = tk.Label(inner, text=self.tr(placeholder_key), bg=theme.SURFACE,
                       fg=theme.TEXT_3, font=F["F_BODY"], anchor="w",
                       cursor="hand2")
        lbl.pack(side="left", fill="x", expand=True)
        tk.Label(inner, text="⌄", bg=theme.SURFACE, fg=theme.SECONDARY,
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

        外观按设计稿做成一条**带边框的圆角条**（齿轮小徽标 + 标题 + 右侧箭头），
        展开的内容长在条内部；条本身也是圆角卡片，所以展开/收起后必须重测量
        （``_relayout_all`` 会连嵌套卡片一起刷新）。
        """
        F = self.F
        self._adv_open = False
        self._spacer(parent)
        self._hairline(parent)

        bar = RoundCard(parent, radius=10, fill=theme.SURFACE,
                        border=theme.BORDER, shadow=False, padx=12, pady=6)
        bar.pack(fill="x", pady=(theme.FIELD_GAP, 0))
        self._adv_card = bar
        self._extra_cards.append(bar)

        hdr = tk.Frame(bar.inner, bg=theme.SURFACE, cursor="hand2")
        hdr.pack(fill="x")
        self._adv_chev = tk.Label(hdr, text="▸", bg=theme.SURFACE,
                                  fg=theme.SECONDARY, font=F["F_LABEL"],
                                  cursor="hand2")
        self._adv_chev.pack(side="right")
        IconBadge(hdr, "gear", size=theme.BADGE_SIZE,
                  bg=theme.SURFACE).pack(side="left")
        tk.Label(hdr, text=self.tr("adv_label"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_LABEL"], cursor="hand2").pack(side="left", padx=(8, 0))

        self._adv_hint = tk.Label(bar.inner, text=self.tr("adv_hint"),
                                  bg=theme.SURFACE, fg=theme.TEXT_3,
                                  font=F["F_HELP"], anchor="w", cursor="hand2")
        self._adv_hint.pack(fill="x", padx=(theme.BADGE_SIZE + 8, 0), pady=(2, 0))

        self._adv_body = tk.Frame(bar.inner, bg=theme.SURFACE)

        # —— 高级格式设置（默认折叠，渐进式披露）——
        # ① 学校模板（可选）
        self._field_label(self._adv_body, "row_template_title", required=False)
        self._helper(self._adv_body, "row_template_desc")
        self._tpl_name = self._select_field(self._adv_body, "row_template_placeholder",
                                            self._pick_template)
        self._spacer(self._adv_body)
        self._hairline(self._adv_body)
        # ② LaTeX 模板（可选）
        self._field_label(self._adv_body, "row_latex_title", required=False)
        self._helper(self._adv_body, "row_latex_desc")
        self._latex_name = self._select_field(self._adv_body, "row_latex_placeholder",
                                              self._pick_latex)
        self._spacer(self._adv_body)
        self._hairline(self._adv_body)
        # ③ 格式问卷
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
                self._adv_chev.config(text="▾")
                self._adv_hint.pack_forget()
            else:
                self._adv_body.pack_forget()
                self._adv_chev.config(text="▸")
                self._adv_hint.pack(fill="x", padx=(theme.BADGE_SIZE + 8, 0),
                                    pady=(2, 0))
        except Exception:
            pass
        # 展开/收起改变了左卡内容高度 —— 必须主动重新测量，
        # 否则新出现的那两行会被圆角矩形裁掉且滚不到（见 _relayout_all）。
        self._relayout_all()
        return self._adv_open

    # ------------------------------------------------------------ 右栏
    def _build_right(self, parent):
        """右卡 Review & Fix：空状态显示 1-2-3 步骤引导，有输入后切换为进度条。

        重设计（2026-09-26）：解决「右栏空白」问题 —— 未选论文时给出清晰步骤图，
        选论文后自然过渡到进度条与状态提示；信任卡和商业区始终保留在底部。
        """
        F = self.F
        head = tk.Frame(parent, bg=theme.SURFACE)
        head.pack(fill="x")
        trow = tk.Frame(head, bg=theme.SURFACE)
        trow.pack(side="left")
        IconBadge(trow, "sliders", size=theme.BADGE_SIZE + 8,
                  bg=theme.SURFACE).pack(side="left")
        tk.Label(trow, text=self.tr("right_title"), bg=theme.SURFACE,
                 fg=theme.TEXT, font=F["F_CARD_HDR"], anchor="w").pack(
                     side="left", padx=(10, 0))
        desc = tk.Label(parent, text=self.tr("right_desc"), bg=theme.SURFACE,
                        fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                        justify="left")
        desc.pack(fill="x", pady=(2, 0))
        _auto_wrap(desc, parent, theme.CARD_PAD_X)

        self._spacer(parent)

        # 空状态：1-2-3 步骤引导（未选论文时显示）
        self._empty_state = tk.Frame(parent, bg=theme.SURFACE)
        self._empty_state.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        tk.Label(self._empty_state, text=self.tr("how_title"), bg=theme.SURFACE,
                 fg=theme.TEXT, font=F["F_LABEL"], anchor="w").pack(fill="x")
        for i, (title, body) in enumerate(self.tr("how_steps"), 1):
            self._build_step_row(self._empty_state, i, title, body)

        # 极简进度条：选论文后显示
        self._build_progress(parent)
        self._progress.pack_forget()

        # 状态行：结果与提示都落这里
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
        self.summary_lbl.pack_forget()
        self.summary_hint.pack_forget()

        self._spacer(parent)
        self._build_trust_card(parent)
        self._build_commercial(parent)

    def _build_step_row(self, parent, index, title, body):
        """空状态步骤行：带序号圆点的 1-2-3 引导。"""
        F = self.F
        row = tk.Frame(parent, bg=theme.SURFACE)
        row.pack(fill="x", pady=(12, 0))
        num = tk.Canvas(row, width=24, height=24, bg=theme.SURFACE,
                        highlightthickness=0, bd=0)
        num.pack(side="left", padx=(0, 10))
        num.create_oval(1, 1, 23, 23, fill=theme.PRIMARY_SOFT,
                        outline=theme.PRIMARY, width=1.5)
        num.create_text(12, 12, text=str(index), fill=theme.PRIMARY,
                        font=F["F_STEP_N"])
        col = tk.Frame(row, bg=theme.SURFACE)
        col.pack(side="left", fill="x", expand=True)
        tk.Label(col, text=title, bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_LABEL"], anchor="w").pack(fill="x")
        lbl = tk.Label(col, text=body, bg=theme.SURFACE, fg=theme.TEXT_2,
                       font=F["F_HELP"], anchor="w", justify="left")
        lbl.pack(fill="x")
        _auto_wrap(lbl, col, 0)

    def _build_progress(self, parent):
        """极简进度条（替代旧版横排步进器）：Prepare → Check → Fix。"""
        F = self.F
        self._progress = ProgressBar(parent, steps=3, bg=theme.SURFACE,
                                     font=F["F_HELP"], height=34)
        self._progress.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        self._progress.set_labels([self.tr("step1_title"), self.tr("step2_title"),
                                   self.tr("step3_title")])

    def _build_trust_card(self, parent):
        F = self.F
        trust = RoundCard(parent, radius=12, fill=theme.SURFACE_SOFT,
                          border=theme.INFO_BORDER, shadow=False, padx=14, pady=10)
        self._trust_card = trust
        self._extra_cards.append(trust)
        # 顺排（不用 side="bottom"）：右栏内容比左栏短时，多出来的高度要落在**卡片底部**，
        # 而不是卡在"ghost 按钮"和"信任卡"中间留一块空洞。
        trust.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        row = tk.Frame(trust.inner, bg=theme.SURFACE_SOFT)
        row.pack(fill="x")
        IconBadge(row, "shield", size=theme.BADGE_SIZE, bg=theme.SURFACE_SOFT,
                  fill=theme.PRIMARY_SOFT).pack(side="left")
        tk.Label(row, text=self.tr("trust_title"), bg=theme.SURFACE_SOFT,
                 fg=theme.PRIMARY_HOVER, font=F["F_LABEL"],
                 anchor="w").pack(side="left", padx=(8, 0))
        cap = tk.Canvas(trust.inner, width=30, height=26, bg=theme.SURFACE_SOFT,
                        highlightthickness=0, bd=0)
        cap.pack(side="right", padx=(8, 0))
        draw_icon(cap, "cap", 15, 13, 26, theme.PRIMARY_SOFT)
        body = tk.Label(trust.inner, text=self.tr("trust_body"),
                        bg=theme.SURFACE_SOFT, fg=theme.TEXT_2, font=F["F_HELP"],
                        anchor="w", justify="left")
        body.pack(fill="x", pady=(3, 0), padx=(theme.BADGE_SIZE + 8, 0))
        _auto_wrap(body, trust.inner, theme.BADGE_SIZE + 8 + 38)

    # ------------------------------------------------------------ 商业区
    def _build_commercial(self, parent):
        """商业区：主页只留一行 ``Trial · X free fixes left`` + ``Upgrade ›``。

        规格 COMMERCIAL UI 明确「主页不要铺四张定价卡」：四个方案收进
        ``_upgrade_dialog`` 的定价弹窗，点 Upgrade 才展开。试用/授权状态全部来自
        ``license.status()``（**只读**），UI 不碰任何计数与权益逻辑。
        """
        F = self.F
        row = tk.Frame(parent, bg=theme.SURFACE)
        row.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        self._commercial_row = row
        self._trial_lbl = tk.Label(row, text="", bg=theme.SURFACE,
                                   fg=theme.TEXT_2, font=F["F_HELP"], anchor="w")
        self._trial_lbl.pack(side="left")
        self._upgrade_btn = RoundButton(
            row, text=self.tr("btn_upgrade"), command=self._upgrade_dialog,
            style="secondary", font=F["F_BTN_S"], height=28, padx=12)
        self._upgrade_btn.pack(side="right")
        self._refresh_commercial()

    def _refresh_commercial(self):
        """按 license 的**只读**状态刷新试用行文案（任何异常都退回中性文案）。"""
        tr = self.tr
        text = tr("trial_none")
        try:
            st = lic.status()
            limit = int(st.get("trial_limit") or 1)
            if st.get("activated"):
                text = tr("trial_licensed")
            elif not st.get("trial_fix_used"):
                left = max(0, limit)
                text = tr("trial_fix_one", left) if left == 1 else tr("trial_fix_many", left)
            else:
                text = tr("trial_none")
        except Exception:
            pass
        try:
            self._trial_lbl.config(text=text)
        except Exception:
            pass

    def _set_step(self, index: int):
        """0=Prepare / 1=Check / 2=Fix / 3=三步全部完成。

        直接驱动底部/右栏的极简进度条：``frac = index / 3``（index=3 → 满格）。
        """
        self._step_index = max(0, min(int(index), 3))
        if getattr(self, "_progress", None) is not None:
            try:
                self._progress.set_step(self._step_index)
            except Exception:
                pass

    def _sync_ui(self):
        """状态机收口：先算状态，再统一做一次重新测量。

        ``_relayout_all()`` 不能省 —— 这里会显隐 ``_fix_btn``、展开/收起 Advanced、
        切换结果明细，都是"内容尺寸变了但不触发 <Configure>"的操作（见该方法 docstring）。
        """
        self._sync_ui_state()
        self._relayout_all()

    def _sync_ui_state(self):
        """主/次按钮、进度条、空状态的状态机（规格 states 段）。

        ============  ====================================================
        状态          主按钮
        ============  ====================================================
        initial       必填未满足 → 置灰（并说明缺什么）
        ready         Check formatting
        checking      Checking…（置灰，禁止重复提交）
        results       有问题 → Fix N issues；无问题 → 保持 Check
        fixed         Review changes
        ============  ====================================================

        主/次按钮已移到左卡输入区底部（紧贴操作区）；右卡空状态与进度条互斥显示。
        """
        tr = self.tr
        paper = bool(self.docx_path.get()) and os.path.isfile(self.docx_path.get())
        # fixing 也算"已有结果"：修正过程中不该把引导卡又弹回来（会闪一下）
        results = self._phase in ("results", "fixed", "fixing")

        # 右卡：空状态（未选论文）与进度条（已选论文）互斥显示
        try:
            if paper:
                self._empty_state.pack_forget()
                self._progress.pack(fill="x", pady=(theme.SECTION_GAP, 0))
            else:
                self._empty_state.pack(fill="x", pady=(theme.SECTION_GAP, 0))
                self._progress.pack_forget()
        except Exception:
            pass

        # 结果明细与引导卡相反：有报告才出现（见 _build_right 的说明）
        # 结果明细与引导卡相反：有报告才出现（见 _build_right 的说明）
        try:
            if results:
                self.summary_lbl.pack(fill="x", pady=(3, 0))
                self.summary_hint.pack(fill="x")
            else:
                self.summary_lbl.pack_forget()
                self.summary_hint.pack_forget()
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
            # 空态也保持主按钮可用：点一下给出友好行内提示（见 _run 的缺稿分支），
            # 比一上来就置灰更有「引导下一步」的明确感（v2.3.4）。
            self._check_btn.set_enabled(True)
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
                # i18n 里 ``*_n`` 词条自带单复数回退（N=1 → "Fix 1 issue"）
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

        # 投放区主行文案：未选时显示「Select a document」，已选时显示文件名 + 大小，
        # 选了非 .docx 时显示错误提示。
        paper = self.docx_path.get()
        if paper and not paper.lower().endswith(".docx"):
            self._paper_name.config(text=tr("st_paper_error"), fg=theme.ERROR)
            self._paper_sub.config(text=tr("paper_support"), fg=theme.TEXT_3)
        elif paper and os.path.isfile(paper):
            name = os.path.basename(paper)
            if len(name) > 40:
                name = name[:18] + " … " + name[-18:]
            self._paper_name.config(text=name, fg=theme.PRIMARY)
            try:
                size = os.path.getsize(paper)
                size_text = f"{size / 1024 / 1024:.2f} MB" if size > 1024 * 1024 else f"{size / 1024:.0f} KB"
            except Exception:
                size_text = tr("paper_support")
            self._paper_sub.config(text=size_text, fg=theme.TEXT_2)
        else:
            self._paper_name.config(text=tr("btn_select_doc"), fg=theme.PRIMARY)
            self._paper_sub.config(text=tr("paper_support"), fg=theme.TEXT_3)
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

    # ------------------------------------------------------------ 拖拽上传
    def _enable_drag_drop(self):
        """开启原生拖拽上传（仅 Windows）。失败静默降级为点选。"""
        if sys.platform != "win32" or _FileDropTarget is None:
            return
        try:
            hwnd = self.root.winfo_id()
            self._drop_target = _FileDropTarget(hwnd, self._on_dropped_files)
        except Exception:
            self._drop_target = None

    def _on_dropped_files(self, paths):
        """拖进来的文件：取第一个；非 .docx 走错误提示（由 _refresh_rows 呈现）。"""
        if not paths:
            return
        self.docx_path.set(paths[0])
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
    def _upgrade_dialog(self):
        """升级 / 定价弹窗（规格 COMMERCIAL UI：主页只留一行，点开才展示四个方案）。

        视觉一律走 ``modal_shell.ModalShell`` + ``widgets.RoundCard`` —— 规格硬要求
        「同一功能不得造第二套样式」，弹窗不许自创标题栏/配色。

        定价与权益逻辑在 ``license`` 层，**本弹窗只做呈现**：四个方案卡照
        ``04_COMPONENTS/pricing-*.svg`` 的版式（名称 / 副题 / 价格 / 说明 / 选择按钮），
        价格给占位符（规格稿本身就是 ``$—``），点「选择方案」接到既有的激活 / 联系流程，
        绝不在这里发明价格或改动任何权益判断。
        """
        tr = self.tr
        F = self.F
        win = ModalShell(self.root, tr("up_title"), width=640, height=840)
        self._track(win)

        body = win.content
        tk.Label(body, text=tr("up_subtitle"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_SMALL"], anchor="w").pack(fill="x")

        # 规格 modal-shell.svg 的两条摘要信息条（One-time / 方案组）
        one = RoundCard(body, radius=9, fill=theme.SURFACE_SOFT,
                        border=theme.INFO_BORDER, shadow=False, padx=14, pady=8)
        one.pack(fill="x", pady=(12, 0))
        tk.Label(one.inner, text=tr("up_plan_onetime"), bg=theme.SURFACE_SOFT,
                 fg=theme.PRIMARY_HOVER, font=F["F_LABEL"], anchor="w").pack(fill="x")
        tk.Label(one.inner, text=tr("up_onetime_desc"), bg=theme.SURFACE_SOFT,
                 fg=theme.TEXT_2, font=F["F_HELP"], anchor="w").pack(fill="x")

        group = RoundCard(body, radius=9, fill=theme.SURFACE_SOFT,
                          border=theme.INFO_BORDER, shadow=False, padx=14, pady=8)
        group.pack(fill="x", pady=(8, 0))
        tk.Label(group.inner, text=tr("up_group_title"), bg=theme.SURFACE_SOFT,
                 fg=theme.PRIMARY_HOVER, font=F["F_LABEL"], anchor="w").pack(fill="x")
        tk.Label(group.inner, text=tr("up_group_desc"), bg=theme.SURFACE_SOFT,
                 fg=theme.TEXT_2, font=F["F_HELP"], anchor="w").pack(fill="x")

        grid = tk.Frame(body, bg=theme.SURFACE)
        grid.pack(fill="both", expand=True, pady=(12, 0))
        grid.columnconfigure(0, weight=1, uniform="plan")
        grid.columnconfigure(1, weight=1, uniform="plan")

        plans = (
            ("up_plan_onetime", "up_plan_onetime_sub", "up_plan_onetime_desc"),
            ("up_plan_weekly", "up_plan_weekly_sub", "up_plan_weekly_desc"),
            ("up_plan_monthly", "up_plan_monthly_sub", "up_plan_monthly_desc"),
            ("up_plan_lifetime", "up_plan_lifetime_sub", "up_plan_lifetime_desc"),
        )
        for i, (name_key, sub_key, desc_key) in enumerate(plans):
            card = RoundCard(grid, radius=12, fill=theme.SURFACE,
                             border=theme.BORDER, shadow=False, padx=14, pady=12)
            card.grid(row=i // 2, column=i % 2, sticky="nsew",
                      padx=(0, 8) if i % 2 == 0 else (8, 0),
                      pady=(0, 10))

            def _choose(_e=None):
                self._close_one(win)
                self._activate_dialog()

            tk.Label(card.inner, text=tr(name_key), bg=theme.SURFACE,
                     fg=theme.PRIMARY_HOVER, font=F["F_LABEL"], anchor="w").pack(fill="x")
            tk.Label(card.inner, text=tr(sub_key), bg=theme.SURFACE,
                     fg=theme.TEXT_2, font=F["F_HELP"], anchor="w").pack(fill="x")
            tk.Label(card.inner, text=tr("up_price_tbd"), bg=theme.SURFACE,
                     fg=theme.PRIMARY, font=F["F_TITLE"], anchor="w").pack(
                         fill="x", pady=(6, 0))
            tk.Label(card.inner, text=tr(desc_key), bg=theme.SURFACE,
                     fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                     justify="left", wraplength=230).pack(fill="x", pady=(2, 8))
            RoundButton(card.inner, text=tr("up_choose"), command=_choose,
                        style="secondary", font=F["F_BTN_S"],
                        height=30, padx=10).pack(fill="x")

        note = tk.Label(body, text=tr("up_note"), bg=theme.SURFACE,
                        fg=theme.TEXT_3, font=F["F_HELP"], anchor="w",
                        justify="left")
        note.pack(fill="x", pady=(6, 0))
        # 说明行跟着弹窗宽度换行：不设 wraplength 会被 590px 的内容区横向裁掉。
        _auto_wrap(note, body, 0)

        win.add_primary_action(tr("up_view_plans"), self._view_plans)
        win.add_cancel_action(tr("up_close"))
        win.add_secondary_action(tr("up_have_code"), self._activate_dialog)
        # 内容比初始高度高时按真实需要长高，保证底部按钮与说明不被裁掉
        # （上限 = 屏幕可用高度，避免小屏弹窗探出可视区）。
        self._fit_modal(win, 640)

    def _fit_modal(self, win, width: int):
        """把弹窗调到内容的真实高度，并保证**整窗留在屏幕工作区内**。

        两个坑：① ``_inner`` 的请求高度只在首帧可靠，必须 ``update_idletasks``
        后再量；② ``winfo_rootx/rooty`` 是**客户区**坐标，含标题栏/边框偏移，
        直接拿来再设一次 ``geometry`` 会越跑越偏（旧写法就是这么把弹窗推到任务栏
        底下的）。这里用 ``winfo_x/y``（窗口外框坐标）定位，并按桌面底部留白夹一次。
        """
        try:
            win.update_idletasks()
            need = int(win._inner.winfo_reqheight()) + 6
            sh = int(self.root.winfo_screenheight())
            # 桌面保留区（任务栏）：1080p 下约 40-48px，取 56 留余量。
            bottom_limit = sh - 56
            cap = max(320, bottom_limit - 8)
            height = min(max(need, 1), cap)
            dy = win.winfo_rooty() - win.winfo_y()
            y = min(win.winfo_y(), bottom_limit - height - dy)
            y = max(8, y)
            win.geometry("%dx%d+%d+%d" % (width, height, win.winfo_x(), y))
        except Exception:
            pass

    def _view_plans(self):
        """「View plans ›」：打开官网（与页脚官网同一入口，不引入新逻辑）。"""
        try:
            webbrowser.open(i18n.BRAND_SITE_URL)
        except Exception:
            pass

    def _activate_dialog(self):
        """激活弹窗 —— 视觉与升级弹窗同一套 modal-shell 语言（规格禁止弹窗自创样式）。

        **只换外壳**：``lic.activate`` / ``lic.activate_offline`` / ``_machine_fingerprint``
        与三个处理函数逐字未动，颜色全部走 theme 令牌。
        """
        tr = self.tr
        F = self.F
        win = ModalShell(self.root, tr("act_title"), width=600, height=660)
        self._track(win)
        body = win.content

        tk.Label(body, text=tr("act_subtitle"), bg=theme.SURFACE,
                 fg=theme.TEXT_2, font=F["F_SMALL"], anchor="w", justify="left",
                 wraplength=500).pack(fill="x")

        def _section(label_key: str, icon: str):
            """一节 = 柔和面圆角卡（与主页信息面板同款），返回卡片内容区。"""
            card = RoundCard(body, radius=12, fill=theme.SURFACE_SOFT,
                             border=theme.INFO_BORDER, shadow=False,
                             padx=14, pady=10)
            card.pack(fill="x", pady=(10, 0))
            head = tk.Frame(card.inner, bg=theme.SURFACE_SOFT)
            head.pack(fill="x")
            IconBadge(head, icon, size=theme.BADGE_SIZE, bg=theme.SURFACE_SOFT,
                      fill=theme.PRIMARY_SOFT).pack(side="left")
            tk.Label(head, text=tr(label_key), bg=theme.SURFACE_SOFT,
                     fg=theme.PRIMARY_HOVER, font=F["F_SMALL_B"], anchor="w",
                     justify="left").pack(side="left", padx=(8, 0))
            return card.inner

        def _hint(parent, key):
            tk.Label(parent, text=tr(key), bg=theme.SURFACE_SOFT,
                     fg=theme.TEXT_3, font=F["F_FOOT"], anchor="w",
                     justify="left").pack(fill="x", pady=(3, 6))

        # ① 在线激活码
        sec = _section("act_online_label", "globe")
        code_var = tk.StringVar()
        ttk.Entry(sec, textvariable=code_var, width=40,
                  font=F["F_BODY"]).pack(fill="x", pady=(8, 0))
        _hint(sec, "act_online_hint")

        def _submit():
            code = code_var.get().strip()
            if not code:
                messagebox.showwarning(tr("act_title"), tr("act_need_code"), parent=win)
                return
            r = lic.activate(code)
            messagebox.showinfo(tr("act_result_title"), r["message"], parent=win)
            if r["ok"]:
                self._close_one(win)

        RoundButton(sec, text=tr("act_online_btn"), command=_submit,
                    style="secondary", font=F["F_BTN_S"], height=34,
                    padx=14).pack(anchor="w")

        # ② 本机机器码（发给卖家换离线码）
        sec = _section("act_machine_label", "mail")
        mc_var = tk.StringVar(value=lic._machine_fingerprint())
        ttk.Entry(sec, textvariable=mc_var, width=40, state="readonly",
                  font=F["F_MONO"]).pack(fill="x", pady=(8, 0))
        _hint(sec, "act_machine_hint")

        def _copy_mc():
            try:
                win.clipboard_clear()
                win.clipboard_append(mc_var.get())
                messagebox.showinfo(tr("act_copy_ok_title"), tr("act_copy_ok"),
                                    parent=win)
            except Exception:
                pass

        RoundButton(sec, text=tr("act_copy_btn"), command=_copy_mc,
                    style="secondary", font=F["F_BTN_S"], height=34,
                    padx=14).pack(anchor="w")

        # ③ 离线激活码（卖家签发，无需联网）
        sec = _section("act_offline_label", "shield")
        off_var = tk.StringVar()
        ttk.Entry(sec, textvariable=off_var, width=40,
                  font=F["F_MONO"]).pack(fill="x", pady=(8, 0))
        _hint(sec, "act_offline_hint")

        def _submit_offline():
            code = off_var.get().strip()
            if not code:
                messagebox.showwarning(tr("act_title"), tr("act_need_offline"), parent=win)
                return
            r = lic.activate_offline(code)
            messagebox.showinfo(tr("act_result_title"), r["message"], parent=win)
            if r["ok"]:
                self._close_one(win)

        RoundButton(sec, text=tr("act_offline_btn"), command=_submit_offline,
                    style="secondary", font=F["F_BTN_S"], height=34,
                    padx=14).pack(anchor="w")

        win.add_cancel_action(tr("up_close"))

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
    root.withdraw()          # 消掉启动一闪：先隐藏主窗，等 UI 全部建完再显示
    App(root)
    root.deiconify()         # UI 构建完毕，正式显示
    root.mainloop()


if __name__ == "__main__":
    main()
