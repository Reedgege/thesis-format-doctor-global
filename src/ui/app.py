"""PaperFormat Pro —— tkinter 桌面 GUI（离线，论文不出本机）。

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
- 高级设置（学校模板 / 问卷 / 配置导入）统一收进 **Advanced options 折叠区**，
  主界面只留必填的「引用规范 + 上传 Word」，落实渐进式披露。
- 主按钮改成**上下文相关**：未选论文=置灰、就绪=Check formatting、
  出结果=Fix N issues、改完=Review changes（见 ``_sync_ui``）。

与国内版的差异（老板 2026-09-11 定）：
1. **默认英文**，右上角一键切中文；语言偏好落盘，下次启动记得住。
2. 英文字体走 Georgia（品牌/标题衬线）+ Segoe UI（正文），不套中文字体。
3. **品牌口径只留官网与邮箱**，不出现微信 / 公众号 / 小程序（海外用户不用）。

功能面（不变）：
  选规范 →（可选）学校模板 / 问卷 / 配置导入 → 选论文
  → 运行体检（免费不限次）→ 一键修正（首次免费，之后需激活码）。
GUI 与 CLI 共用同一套 engine，保证逻辑唯一、离线安全。
"""

from __future__ import annotations

import os
import sys
import ctypes
import json
import re
import queue
import threading
import tkinter as tk
import traceback
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from ..engine import checker, report as report_mod
from ..engine.docx_reader import DocxReadError, PAPER_SIZES
from ..engine.fixer import compute_changes, fix_docx, same_file, _compute_changes
from ..engine.questionnaire import (
    FIELD_RANGES, QUESTIONNAIRE_SCHEMA, TargetProfile, build_target,
    apply_page_overrides, PAGE_NUMBER_POSITIONS, target_from_spec,
)
from ..engine.specs import (SPEC_ORDER, spec_display, REFERENCE_STYLE_KEYS,
                            REFERENCE_STYLE_CHOICES, spec_is_baseline)
from ..license import license as lic
from .. import versioninfo
from . import i18n, iconpath, recents, theme
from . import dpi as _dpi_mod
from .fonts import Fonts
from .modal_shell import ModalShell
from .widgets import (BrandMark, IconBadge, ProgressBar, RoundButton, RoundCard,
                      ScrollArea, draw_icon)


# (启动诊断代码已随 v2.3.39 移除：空白问题已在真机用诊断日志确认根治，
#  不再写 %TEMP%/tfd_startup_diag.log)


class _PausedPaint:
    """重绘闸门（配合 ``App._freeze_paint`` / ``_thaw_paint``，见那里的长注释）。

    用法：``with self._paused_paint(): ...布局变动...`` —— 块内不上屏，出块时
    一次性画成最终画面。**必须**在 ``finally`` 里解冻：中途抛异常若忘了开回
    WM_SETREDRAW，窗口会永远空白，那是比花屏严重得多的 P0。
    """

    __slots__ = ("_app",)

    def __init__(self, app):
        self._app = app

    def __enter__(self):
        try:
            self._app._freeze_paint(True)
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            self._app._freeze_paint(False)
        except Exception:
            try:
                self._app._thaw_paint()
            except Exception:
                pass
        return False


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
_TEMPLATE_PATH = os.path.join(_HERE, "..", "data", "ai_questionnaire_template.yaml")

_FIELD_RANGE = FIELD_RANGES


def _opt_val(opts, disp):
    """下拉「显示名 → 引擎值」反查（显示名是给人看的，绝不把文字反解析成值）。

    ``opts`` 是 [(value, display), ...]；找不到返回 None（= 不设定）。
    """
    for value, text in opts or ():
        if text == disp:
            return value
    return None

# 引用规范下拉的展示名：引擎 key 已是分层后的 APA-Student / APA-Professional，
# 这里只把界面展示名写清楚（学生版 / 专业版），其余风格用 key 本身。
# 展示名统一取自 engine.specs.spec_display（与报告共用单一来源）。
_spec_display = spec_display


def _columns_display_value(columns, none_txt, one_txt, two_txt):
    """栏数 → 复核弹窗下拉的显示名（纯函数，便于 headless 单测）。

    🔴 v2.3.102（老板 2026-10-24 截图）：``None`` = 该来源**不规定**栏数
    （如 IEEE-Journal 按裁定「不碰版式」）→ 必须显示「—」。
    绝不能默认成 "Single column"：既在说谎，又会在确认时把 None 静默升级成
    「强制单栏」，双栏期刊论文一检查就报"应为单栏"。
    """
    if columns is None:
        return none_txt
    try:
        return two_txt if int(columns) == 2 else one_txt
    except (TypeError, ValueError):
        return none_txt


def _columns_edit_value(display, none_txt, two_txt):
    """下拉显示名 → ``apply_page_overrides`` 的原始值（纯函数）。

    「—」必须回**空串**：``apply_page_overrides`` 只把空串当"未改动"；
    传 "1" 会把 ``target.columns`` 从 None 改写成 1 并把 resolution 升级为 user。
    """
    if display == none_txt:
        return ""
    return "2" if display == two_txt else "1"

# 字段行左侧的小圆徽标：v2.3.17 起由 IconBadge 内部自动查找对应 PNG 图片；
# 这里的名字仍是设计稿线描图标名，映射表在 src/ui/icon_loader.ICON_NAME_MAP。
# row_paper_title 单独用 "paper"，以便与左卡大标题 "doc" -> your_documents 区分。
_FIELD_ICONS = {
    # v2.3.62 改版把规范行的键从 row_spec_title 改成 source_title，图标映射没跟上
    # → 「格式来源」行的引号图标静默丢失（2026-10-20 老板发现）。quote -> citation_style PNG。
    "source_title": "quote",
    "row_paper_title": "paper",
    "row_template_title": "bank",
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
    # v2.3.80：与 ScrollArea 同口径 —— 显式 1 unit == 1px。不设时 canvas 的
    # 1 unit = 视口 1/10，下面 ``yview_scroll(-e.delta, "units")`` 的「每格约 120px」
    # 会实际变成一格 3000+px（直接到底）。
    canvas = tk.Canvas(parent, bg=bg, highlightthickness=0, yscrollincrement=1)
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
            # v2.3.22：按 delta 像素滚动（旧版每格只滚 1px，滚轮极不灵敏）。
            canvas.yview_scroll(-e.delta, "units")
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

        v2.3.52：光刷新卡片**不够**，必须连外层滚动区一起收口。卡片把画布请求高度
        顶到 530 之后，外层 grid 的行高是由``ScrollArea`` 钉在画布 window item 上的
        那个快照决定的；而 ``ScrollArea._relayout`` 只在 :meth:`refresh_layout` 被调
        时才重算。首屏自动换行重排就发生在这条路径上（此前的
        ``_schedule_first_layout`` 早已跑完），于是快照停在 517 → 行高 517 →
        右卡（sticky="new"，高度全看自然高度）底部被裁 13px。真机取证：
        自然 530 / 实际 517，补一次 area 刷新即回到 530。

        两段都只在"宽度真的变了"（``_apply`` 里``w == state["wl"]`` 已提前返回）
        且窗口已映射时才跑，不会退化成构建期那种跨卡级联。
        """
        card = None
        area = None
        try:
            while w is not None:
                if card is None and isinstance(w, RoundCard):
                    card = w
                elif area is None and isinstance(w, ScrollArea):
                    area = w
                if card is not None and area is not None:
                    break
                w = w.master
        except Exception:
            pass
        try:
            if card is not None:
                card.refresh()
        except Exception:
            pass
        try:
            if area is not None:
                area.refresh_layout()
        except Exception:
            pass

    def _apply(_e=None):
        if state["busy"]:
            return
        # v2.3.50：窗口还没真正显示（main() 先 withdraw 再建 UI，构建期几何是 1px 占位值）
        # 此刻算出来的 wraplength 毫无意义，还会和 refresh 互相追着跑、跨多张 RoundCard
        # 在同一次 update_idletasks 里无限级联，把启动从图标修复后的 ~10s 又拖回去。
        # 没映射之前直接跳过，等 deiconify 后的真实 <Configure> / _schedule_first_layout
        # 再统一收口，一次收敛。
        try:
            if not container.winfo_viewable():
                return
        except Exception:
            pass
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


# 月份缩写写死成英文，避免 strftime("%b") 依赖操作系统区域设置
# （中文区域的机器上 "%b" 会渲染成中文月份，破坏英文界面的纯净度）。
_MONTHS_EN = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def _fmt_expiry(value):
    """把 license 层返回的 ``expires_at`` 格式化成西方习惯的 ``Nov 1, 2026``。

    月份缩写硬编码为英文（见 ``_MONTHS_EN``），不依赖系统区域设置，
    保证英文界面在任何操作系统上都不会冒出中文月份。

    兼容多种来源格式：Unix 时间戳（int/float）、ISO 8601（含 ``Z`` 后缀）、纯日期。
    解析失败一律返回 ``None``，调用方退化成不带日期的授权文案。
    """
    if not value:
        return None
    if isinstance(value, (int, float)):
        try:
            v = float(value)
            # 毫秒时间戳（Unix epoch ms）：秒级上限约 1.7e9，值 > 1e11 必为毫秒，先除 1000，
            # 否则 datetime.fromtimestamp 会因数值过大抛异常被吞 → 状态行退化成无日期。
            if v > 1e11:
                v = v / 1000.0
            dt = datetime.fromtimestamp(v)
            return "%s %d, %d" % (_MONTHS_EN[dt.month - 1], dt.day, dt.year)
        except Exception:
            return None
    s = str(value).strip()
    if not s:
        return None
    iso = s.replace("Z", "+00:00")
    for parser in (lambda: datetime.fromisoformat(iso),
                   lambda: datetime.strptime(s, "%Y-%m-%d"),
                   lambda: datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")):
        try:
            dt = parser()
            return "%s %d, %d" % (_MONTHS_EN[dt.month - 1], dt.day, dt.year)
        except Exception:
            continue
    return None


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
        # v2.3.32 重绘闸门：嵌套深度 + 是否已关闸（见 _freeze_paint 的长注释）
        self._paint_depth = 0
        self._paint_frozen = False
        self._closing = False        # 根窗口已销毁（见 _on_destroy）
        self._resize_after = None    # 启动后首帧的延迟重排句柄
        # 首帧真实尺寸未就绪标记：deiconify 后 WM 尚未报告客户区尺寸，winfo_width
        # 仍是 1×1，此刻跑 _relayout_all 会把卡片按 1px 残框画成首屏（启动空白根因，
        # v2.3.36 的 _flush_frame 反而把残帧提前画上了屏）。等首次拿到真实宽度再
        # 统一重排一次，见 _on_resize 里的 _first_layout_done。
        self._first_layout_done = False
        # v2.2.0 状态机：phase 决定主按钮文案/可用态与步进位置（见 _sync_ui）
        self._phase = "idle"          # idle / ready / results / fixed / error
        self._last_issues = None      # 最近一次体检的问题数（warn+fail）
        self._checked_path = ""       # 出结论时那份文档 —— 换了文档就作废旧结论
        self._last_output = ""        # 最近一次修正落盘路径（供「查看改动」）
        self._status_spec = (None, "idle", "")   # 状态行最近一次内容（切语言后重渲染）

        # 输入源
        self.spec_key = tk.StringVar(value="APA-Student")
        self.school_template = tk.StringVar(value="")
        self.ai_json = tk.StringVar(value="")
        self.docx_path = tk.StringVar(value="")
        self.questionnaire_data: dict = {}
        # 6 选 1 来源：前 4 项带内置规范（spec_key 非空），后 2 项走无规范基线
        self.source_mode = tk.StringVar(value="src_apa_student")
        # 学校模板子下拉：引用 / 参考文献风格（默认留空=不检查引用）
        self.school_ref_style = tk.StringVar(value="")
        # 高级区是否可用（选内置规范时整片置灰）
        self._adv_enabled = True

        # v2.3.78：最近使用（手填问卷模板 / AI 导入文件）的控件与缓存
        self._manual_recent_items: list = []   # 与下拉项一一对应的问卷 dict 列表
        self._ai_recent_items: list = []       # 与下拉项一一对应的文件路径列表
        self._restore_last_session()           # 启动自动恢复上一次（落盘在 recents.json）

        _sw = root.winfo_screenwidth()
        _sh = root.winfo_screenheight()
        # 可用高度 = 屏幕高 − 任务栏预留。预留 64px（与 _fit_modal 同口径，覆盖任务栏
        # + 标题栏 + 边距），并在可用区内居中。否则小屏（1280×720 / 1366×768 / 150%
        # 缩放）上 0.94 这种系数算出的初始高度几乎顶满，WM 再往下偏一点，底部页脚/
        # 状态栏就被任务栏吃掉 —— 正是 v2.1.1 修过的那个 P0 的现世报。
        _taskbar_reserve = 64
        _avail = max(360, _sh - _taskbar_reserve)
        root.title("PaperFormat Pro")
        _w = min(_BASE_W, int(_sw * 0.88))
        _h = min(1020, _avail - 16)
        _h = max(480, _h)
        _x = max(8, (_sw - _w) // 2)
        _y = max(8, (_avail - _h) // 2)
        root.geometry("%dx%d+%d+%d" % (_w, _h, _x, _y))
        # v2.3.82：记下算出来的起始几何。main() 要先把窗口挪到屏幕外完成首帧绘制
        # （治启动花屏），画完再挪回这里 —— 挪回时得知道"这里"是哪。
        self._start_geometry = "%dx%d+%d+%d" % (_w, _h, _x, _y)
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
        # v2.3.48：窗口真正映射到屏幕时（deiconify → <Map>）再设一次图标，这是任务栏
        # 图标生效的最稳时机；此前靠 winfo_viewable() 轮询守卫会在编译后 exe 上永久拦住
        # 补设，导致任务栏无图标。该回调幂等无害，最小化/还原再次映射时也会重设。
        try:
            root.bind("<Map>", self._on_window_mapped)
        except Exception:
            pass
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
        # v2.3.32：整个过程关掉重绘。量卡片要一连跑好几次 update_idletasks，
        # 每一次都会把"还没走完"的中间画面推上屏（旧像素还残留在让出来的区域里），
        # 屏幕上就同时出现新旧两套内容 = 花屏。闸门一关，中间态一帧都不上屏。
        with self._paused_paint():
            self._relayout_pass()
            # v2.3.52：补一次**收敛**重排。首屏右卡会被裁掉 13px（真机取证：
            # 自然高度 530、实际只给到 517），因为 ScrollArea._relayout 量
            # ``natural = inner.winfo_reqheight()`` 时，卡片内的自动换行标签还没按
            # 最终列宽重排完，读到的是偏小的旧值 → 把画布 window item 钉死在 517 →
            # 外层 grid 的行也就只有 517 → 右卡（sticky="new"，高度全看自然高度）
            # 被压掉 13px，多半是信任卡/商业区那几行，用户看着像"右卡内容被遮挡"。
            #
            # 判据不能用"卡片实际高< 卡片请求高"：那一刻卡片自己的请求高也还是旧值
            # （实测 517 vs 517），会误判成没被裁而跳过补排（栽过一次的坑）。
            # 真正可靠的是**滚动区钉死的高度 vs 它内容此刻的自然高度** —— 前者是
            # 上一次量到的快照，后者会随后续重排涨上来，差值 >0 就说明还欠一次。
            # 只在欠的时候才补那一次，避免每次都多跑一遍 update_idletasks
            # （那正是 v2.3.49 启动卡顿 30s 的来源）。
            if self._area_needs_more_height():
                self._relayout_pass()

    def _relayout_pass(self):
        """一轮重新测量：先内后外（嵌套卡片没有自己的布局入口）。"""
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

    def _area_needs_more_height(self):
        """滚动区钉死的高度是否已经小于其内容此刻的自然高度（欠一次重排）。

        收敛判据（v2.3.52）。``ScrollArea._relayout`` 会把画布 window item 的高度
        钉成 ``max(内容自然高度, 视口高)``；若它是在自动换行重排完成**之前**量的，
        就会钉在偏小的旧值上，之后内容自然高度涨上来、这个快照却不再更新 →
        外层 grid 的行高偏小 → 右卡底部被裁。两者出现差值就是欠排。

        留 1px 容差避开 Tk 整数取整抖动；取不到就返回 False（不额外重排，宁可
        少排也不要引入新的启动开销）。
        """
        area = getattr(self, "_main_scroll", None)
        if area is None:
            return False
        try:
            pinned = area.inner.winfo_height()
            natural = area.inner.winfo_reqheight()
        except Exception:
            return False
        return natural > pinned + 1

    # ------------------------------------------------------------ 基础工具
    def tr(self, key: str, *args):
        return i18n.t(self.lang, key, *args)

    def _sep(self) -> str:
        return "｜" if self.lang == "zh" else "|"

    def _apply_window_icon(self):
        """设置标题栏/任务栏图标。**整个函数绝不允许向外抛异常**——图标是锦上添花，
        绝不能因为一个平台上的路径/Tcl 怪癖就让 App 构造失败（那正是「双击打不开」
        那一类 P0 后果）。

        两条路都**故意推迟到窗口真正显示之后**再执行（见 ``_schedule_icon_apply``）：

        ① PNG + ``iconphoto``：跨平台标准做法。``PhotoImage`` 必须挂到实例属性上
           保留引用——只作为临时对象传进去时一旦被 GC 回收，部分 Tk 版本会把图标
           还原成默认（经典坑，v2.0.2 踩过）。
        ② Windows 的 ``.ico`` + Win32 API：部分 Tcl/Nuitka 打包后 ``iconbitmap``
           会静默不设置图标句柄（实测 ``WM_GETICON`` 仍返回 0），导致左上角仍是
           Windows 默认程序图标。这里绕过 Tcl，用 ``LoadImageW`` 读原生 ``.ico``，
           再 ``WM_SETICON`` 直接写当前窗口句柄，并 ``SetClassLongPtr`` 同步窗口类
           图标，让后续弹出的 Toplevel 也带上图标。
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
                        self._set_win_icon_from_ico(ico)
                    except Exception:
                        pass
        except Exception:
            pass
        # v2.3.47：Tk 的 iconphoto / Win32 的 WM_SETICON 都必须在窗口 **已映射**
        # 之后才生效。在 withdraw() 状态下设置会被 Tk/Windows 直接丢弃——这正是
        # 真机实测（v2.3.46 及更早）任务栏/标题栏一直没有图标的原因：
        # WM_GETICON 与 ClassLongPtr(HICON) 全为 0，与用哪版图标文件无关。
        # 因此这里再排一次「窗口显示后」的补设。
        # 注意：v2.3.47 那次补设里用 winfo_viewable() 做「等映射」守卫，结果在编译后
        # exe 上该判断始终返回 False，把补设永久拦死，反而让任务栏图标彻底没了
        # （见 _reapply_window_icon / _on_window_mapped 的 v2.3.48 修正）。
        self._schedule_icon_apply()

    def _schedule_icon_apply(self):
        """窗口显示后补设图标（见 ``_apply_window_icon`` 末尾的说明）。

        ``deiconify()`` 之后 Tk 才真正映射窗口；再叠加一次 ``update_idletasks()``
        确保 WM 已收到映射消息。这里用 ``after`` 延后若干毫秒执行，避开映射与
        首次布局的竞争。所有失败静默——图标只是锦上添花。
        """
        for delay in (0, 60, 300, 1200):
            try:
                self.root.after(delay, self._reapply_window_icon)
            except Exception:
                pass

    def _reapply_window_icon(self):
        """窗口已显示后重新设置图标（幂等；失败不影响任何功能）。

        **不再**用 ``winfo_viewable()`` 守卫去「等窗口映射」——本应用里该判断在编译后
        exe 上会**始终返回 False**（真机/源码探针均复现），导致补设被永久拦截、任务栏
        图标永远设不上（v2.3.47 的回退反而引入此回归）。图标设置本身是幂等且无害的：
        未映射时 WM 忽略、映射后后续重试即生效。真正的「映射后才生效」保证改由
        ``<Map>`` 事件绑定（``_on_window_mapped``）提供——窗口一旦真正可见，该回调必然
        触发，此刻 ``WM_SETICON`` 一定能写进去。
        """
        try:
            if getattr(self, "_closing", False):
                return
            self._apply_window_icon_once()
        except Exception:
            pass

    def _on_window_mapped(self, _evt=None):
        """窗口真正被映射到屏幕时回调：此刻 ``WM_SETICON`` 才会被 Windows 采纳，是
        任务栏图标生效的最可靠时机（比轮询 ``winfo_viewable()`` 稳得多）。幂等、静默。

        绑定见 ``App.__init__``：deiconify() 触发 <Map>，此时窗口已可见，设置必生效。
        窗口被最小化再还原、或重新 deiconify 时 <Map> 会再次触发，重设幂等无害。

        ⚠️ v2.3.49 修复启动卡 30s 的根因：``<Map>`` 事件会沿 bindtag 链**冒泡**到根
        窗口，于是构建期每个子控件被 pack/grid 映射时，都会触发根上这个 ``<Map>``
        绑定，导致 ``_apply_window_icon_once`` 被调用约 100 次（每次含 tk.PhotoImage
        + Win32 LoadImageW，约 0.3s），累计 30s+。只有「根窗口自身」的 ``<Map>``
        才是真可见，必须忽略所有 ``event.widget is not self.root`` 的冒泡事件。
        """
        # 忽略子控件冒泡上来的 <Map>：只认根窗口自己的映射事件。
        if _evt is not None and getattr(_evt, "widget", None) is not self.root:
            return
        try:
            if getattr(self, "_closing", False):
                return
            self._apply_window_icon_once()
        except Exception:
            pass

    def _apply_window_icon_once(self):
        """真正执行设置（``_reapply_window_icon`` 的实际动作，不递归调度）。

        ``_window_icon`` 只创建一次并缓存：根窗口的 ``<Map>`` 在最小化/还原、重新
        deiconify 时会再次触发，复用已加载的 PhotoImage 即可，避免重复读文件
        （见 ``_on_window_mapped`` 的 v2.3.49 启动卡顿修复说明）。
        """
        try:
            png = iconpath.find_icon()
            if png:
                try:
                    if getattr(self, "_window_icon", None) is None:
                        self._window_icon = tk.PhotoImage(file=png)
                    self.root.iconphoto(True, self._window_icon)
                except Exception:
                    pass
            if sys.platform.startswith("win"):
                ico = iconpath.find_icon_ico()
                if ico:
                    try:
                        self._set_win_icon_from_ico(ico)
                    except Exception:
                        pass
        except Exception:
            pass

    def _set_win_icon_from_ico(self, ico_path: str):
        """Windows：用 Win32 API 把 ``.ico`` 设成标题栏/任务栏/类图标。

        必须在窗口 **已映射** 之后调用（见 ``_reapply_window_icon``）：窗口还在
        withdraw 状态时 ``WM_SETICON`` 会被丢掉，真机实测 ``WM_GETICON`` 恒为 0。

        任何异常都向上抛，由调用方吞掉，绝不影响启动。
        """
        user32 = ctypes.windll.user32
        hwnd = self.root.winfo_id()
        if not hwnd:
            return
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x0010
        LR_DEFAULTSIZE = 0x0040
        WM_SETICON = 0x0080
        WM_GETICON = 0x007F
        GCLP_HICON = -14
        GCLP_HICONSM = -34
        ICON_SMALL = 0
        ICON_BIG = 1
        # 不用 LR_SHARED：共享句柄不销毁会在多次补设时泄漏，且共享句柄交给
        # 窗口类后可能被系统释放。改为私有句柄，补设前先释放自己上一轮设的。
        flags = LR_LOADFROMFILE | LR_DEFAULTSIZE

        # 先清掉上一轮自己设的句柄，避免重复补设时泄漏（旧句柄会一直留着）。
        for wparam, cls_idx in ((ICON_BIG, GCLP_HICON), (ICON_SMALL, GCLP_HICONSM)):
            try:
                old = user32.GetClassLongPtrW(hwnd, cls_idx)
                if old:
                    cur = user32.SendMessageW(hwnd, WM_GETICON, wparam, 0)
                    if cur == old:
                        user32.DestroyIcon(old)
            except Exception:
                pass

        hsm = user32.LoadImageW(None, ico_path, IMAGE_ICON, 16, 16, flags)
        hlg = user32.LoadImageW(None, ico_path, IMAGE_ICON, 32, 32, flags)
        if not hsm and not hlg:
            return
        if hsm:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hsm)
            user32.SetClassLongPtrW(hwnd, GCLP_HICONSM, hsm)
        if hlg:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hlg)
            user32.SetClassLongPtrW(hwnd, GCLP_HICON, hlg)

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
        # 首帧：窗口拿到真实客户区尺寸后，才第一次真正重排卡片内容。此前（1×1 时）
        # _relayout_all 已把卡片画成残框，这里补一次真实尺寸下的重排把首屏纠正过来，
        # 不依赖用户"点一下上边缘"（那只是靠点击触发 <Configure> 才碰巧修好）。
        if not self._first_layout_done and w > 2:
            self._first_layout_done = True
            try:
                self._relayout_all()
            except Exception:
                pass
        factor = self.F.scale_for_width(w)
        if self._last_scale is not None and abs(factor - self._last_scale) < 0.02:
            # scale 没变也要把主体顶留白重设成当前 _top_gap()：顶栏高度在重建/切语言
            # 后的首帧常尚未测量（winfo_reqheight 返回 1），旧代码只在 scale 变化时
            # 才重设 pady，于是内容一直压在顶栏底下、错位到必须来一次真正 resize 才
            # 纠正——这正是用户反馈的「切语言界面飘、最大化再还原就好」的根因
            # （最大化→还原改了窗口宽度 → scale 变 → 此分支才被触发 → 留白才纠正）。
            # 切语言后窗口宽度不变、scale 不变，所以必须在早退前补这一刀（v2.3.8）。
            try:
                self._main_scroll.pack_configure(pady=(self._top_gap(), 0))
            except Exception:
                pass
            return
        self._last_scale = factor
        self.F.apply_scale(factor)
        # 字体缩放会让顶栏变高 → 主体顶部留白要跟着重算（否则顶栏会和卡片叠上）
        try:
            self._main_scroll.pack_configure(pady=(self._top_gap(), 0))
        except Exception:
            pass

    def _schedule_first_layout(self):
        """首帧布局：轮询直到窗口拿到真实客户区宽度(w>2)再 relayout + 强制整窗重绘。

        不依赖 <Configure> 何时派发（缩放屏/Nuitka 下它常迟到、甚至要等点击才来）；
        也不在尺寸未就绪(1x1)时提前画残帧。每 30ms 查一次，拿到真实宽后立即纠正首屏，
        并把画面用 _flush_frame(RedrawWindow UPDATENOW) 强制画出来。

        v2.3.82：这一帧画完就调用 ``on_first_frame_ready`` 注册的回调（main() 用它把
        窗口从屏幕外移回最终位置）—— 客户看到的第一眼就是成品。
        """
        if getattr(self, "_first_layout_flushed", False):
            return
        try:
            w = self.root.winfo_width()
        except Exception:
            w = 0
        if w <= 2:
            self.root.after(30, self._schedule_first_layout)
            return
        try:
            if not self._first_layout_done:
                self._first_layout_done = True
                self._relayout_all()
        except Exception:
            pass
        self._first_layout_flushed = True
        self._flush_frame()
        # v2.3.82：首帧画完的钩子（main() 用它把窗口从屏幕外挪回来）。
        # 只跑一次，异常不影响后续。
        _cb = getattr(self, "_first_frame_cb", None)
        if _cb is not None:
            self._first_frame_cb = None
            try:
                _cb()
            except Exception:
                pass

    def on_first_frame_ready(self, cb):
        """注册"首帧已画完"回调（只跑一次）。用于启动时把窗口移入最终位置。"""
        self._first_frame_cb = cb

    # ------------------------------------------ 启动花屏：三条走不通的路（v2.3.82）
    #
    # 老板 2026-10-21：「打开软件的时候，启动时，刚进入界面会有一阵花屏」。
    # 取证脚本与结论见 main() 里的注释。这里只记三条**实测过、不能走**的路，
    # 免得以后有人（包括我）再花时间重试：
    #
    # ① 隐藏状态下先排版（"先排好再显示"）
    #    withdraw + update_idletasks() 会让 Tk 把请求的几何**提前兑现**到内部模型：
    #    deiconify 之后 Tk 以为自己已经是 1126x640、不再向系统申请尺寸，而真实 HWND
    #    只有 1x1 —— 窗口**永久停在 1x1**（A/B 实测：不预排则会正常变成 1126x640）。
    #    窗口真实客户区尺寸只能等 map 之后由系统给，抢不得。
    #
    # ② 用 ``-alpha 0`` 当遮罩（画完再淡入）
    #    Tk 在 alpha=0 状态下**不产生任何重绘**；把 alpha 调回 1.0 也不补画，
    #    实测 4 秒后屏幕上依然是桌面（窗口永久空白）。比花屏严重得多。
    #
    # ③ 复用 v2.3.32 的 WM_SETREDRAW 重绘闸门
    #    ``_flush_frame()`` 开头那条 ``_thaw_paint()`` 保险丝会在首帧 flush 时立刻把
    #    闸门清零（实测 depth 1→0），闸门等于没关。那条保险丝是故意的（防"窗口永远
    #    画不出来"），不能为这个需求拆。
    #
    # 结论：只剩"在屏幕外完成映射与首帧绘制、画完再移入"这一条路（见 main()）。

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

    def _link(self, parent, text, cmd, fg=ACCENT, tooltip=None):
        lbl = tk.Label(parent, text=text, bg=PAPER, fg=fg,
                       font=self.F["F_SUBTITLE"], cursor="hand2")
        lbl.bind("<Button-1>", lambda e: cmd())
        lbl.bind("<Enter>", lambda e: lbl.config(fg="#33465c"))
        lbl.bind("<Leave>", lambda e: lbl.config(fg=fg))
        if tooltip:
            self._attach_tooltip(lbl, tooltip)
        return lbl

    # ------------------------------------------------------------ 极简 tooltip
    def _attach_tooltip(self, widget, text):
        """悬停显示一行说明；不影响 _link 既有的颜色 hover。"""
        tip = {"w": None}

        def show(_e):
            if tip["w"] or not text:
                return
            x = widget.winfo_rootx() + 12
            y = widget.winfo_rooty() + widget.winfo_height() + 6
            tw = tk.Toplevel(widget)
            tw.wm_overrideredirect(True)
            tw.attributes("-topmost", True)
            tk.Label(tw, text=text, bg="#2C2C2A", fg="#FFFFFF",
                     font=self.F["F_SMALL"], padx=7, pady=4).pack()
            tw.wm_geometry("+%d+%d" % (x, y))
            tip["w"] = tw

        def hide(_e):
            if tip["w"]:
                tip["w"].destroy()
                tip["w"] = None

        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)

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
        # 下拉框（选择框）：静止薄荷灰描边 + 主色箭头，聚焦时描边转主色
        style.configure("TCombobox", font=F["F_BODY"], padding=4,
                        background=theme.SURFACE,
                        fieldbackground=theme.SURFACE,
                        foreground=theme.TEXT,
                        bordercolor=theme.SELECT_BORDER,
                        arrowcolor=theme.SELECT_ARROW)
        style.map("TCombobox",
                  bordercolor=[("focus", theme.SELECT_BORDER_FOCUS),
                               ("active", theme.SELECT_BORDER_FOCUS)],
                  fieldbackground=[("focus", theme.SURFACE),
                                   ("active", theme.SURFACE_SOFT)])
        # 滚动条：极细、浅灰、无上下三角箭头；装得下时本就隐藏，溢出时由
        # ScrollArea 在 hover/滚动时才显形（静止隐去）。clam 主题下 width 与去箭头
        # 布局才生效。
        style.configure("TScrollbar",
                        background=theme.SCROLLBAR_THUMB,
                        troughcolor=theme.SCROLLBAR_TROUGH,
                        bordercolor=theme.SCROLLBAR_TROUGH,
                        arrowcolor=theme.SCROLLBAR_ARROW,
                        lightcolor=theme.SCROLLBAR_THUMB,
                        darkcolor=theme.SCROLLBAR_THUMB,
                        gripcount=0)
        style.map("TScrollbar",
                  background=[("active", theme.SCROLLBAR_THUMB_HOVER),
                              ("pressed", theme.SCROLLBAR_THUMB_HOVER)])
        # v2.3.15：竖向滚动条专门定宽 + 去掉箭头元素，避免"老式 Windows 滑块"。
        style.configure("Vertical.TScrollbar", width=8)
        style.layout("Vertical.TScrollbar",
                     [("Scrollbar.trough",
                       {"children": [("Scrollbar.thumb", {"expand": 1})]})])
        # 下拉弹出的列表：白底深字，选中行主色高亮（option_add 全局生效，一次即可；
        # 语言切换时 _rebuild 处的同名 option_add 会再兜底）。
        self.root.option_add("*TCombobox*Listbox.background", theme.SURFACE)
        self.root.option_add("*TCombobox*Listbox.foreground", theme.TEXT)
        self.root.option_add("*TCombobox*Listbox.selectBackground",
                             theme.SELECT_LIST_SEL_BG)
        self.root.option_add("*TCombobox*Listbox.selectForeground",
                             theme.SELECT_LIST_SEL_FG)

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
        self._build_header()

        # 主体：可滚动区。窗口比内容矮时（1366×768 / 1600×900 笔记本）出滚动条，
        # 绝不把卡片底部裁掉 —— 旧版没有这层，左卡最下面的字段会被圆角矩形切掉。
        # 顶部留白 = 顶栏高度 + HEADER_GAP：顶栏画在背景上不占 pack 空间，而这段留白
        # 下面是背景画布，所以**背景（含天际线）正好从这条带露出来**。
        self._main_scroll = ScrollArea(self.root, bg=theme.BG)
        self._main_scroll.pack(fill="both", expand=True, padx=theme.PAGE_PAD,
                               pady=(self._top_gap(), 0))

        main = tk.Frame(self._main_scroll.inner, bg=theme.BG)
        # v2.3.16：网格容器右侧留 12px 余白（外部 padx），给自动隐藏滚动条落位，
        # 既让两卡保持等宽，又让滚动条压在留白里、不挡右卡右侧的 Upgrade 按钮 / 图标
        # （老板实测"右边的字被覆盖"）。两列仍按窗口半宽均分，等宽断言不受影响。
        main.pack(fill="both", expand=True, padx=(0, 12))
        main.columnconfigure(0, weight=1, uniform="half", minsize=380)
        main.columnconfigure(1, weight=1, uniform="half", minsize=380)
        main.rowconfigure(0, weight=1)
        self._main_frame = main
        # 两栏宽度只由窗口决定（见 _sync_columns），否则卡片里的自动换行会和 grid
        # 的列宽分配互相追着跑。
        main.bind("<Configure>", self._sync_columns, add="+")

        # 两栏等宽；grid 两边各给半个 CARD_GAP，合计正好一个卡片间距
        # sticky="new"（上+左右、不向下拉伸）：左卡展开 Advanced 时自身会涨高，
        # 右卡只按自己的自然高度停在顶部，绝不被左卡"拽"下去（老板 2026-09-26 反馈
        # 右卡跟着左卡下拉很难看）。两卡都不再撑满整窗高度，避免空出一大块死板卡片。
        # 主卡片去投影、留干净的 1px 边框（v2.3.10：老板反馈投影几乎看不见、
        # 反而显得边框"缺了点什么"）。边框色统一用 theme.BORDER（#D2DEDD），
        # 和卡片内字段细线之外的所有分区分隔保持一致。
        self._left_card = RoundCard(main, padx=theme.CARD_PAD_X,
                                    pady=theme.CARD_PAD_Y, shadow=False)
        self._left_card.grid(row=0, column=0, sticky="new",
                             padx=(0, theme.CARD_GAP // 2))
        self._right_card = RoundCard(main, padx=theme.CARD_PAD_X,
                                     pady=theme.CARD_PAD_Y, shadow=False)
        self._right_card.grid(row=0, column=1, sticky="new",
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
        # 全量重建后强制所有圆角卡片（含 Advanced 折叠条）统一重新测量并补全边框：
        # 首帧量到的可能是 1px 内容高度，圆角矩形会被画成残框/被裁切；等布局落定再
        # refresh 一次，确保外框完整（见 RoundCard.refresh 的 P0-2 说明）。
        self.root.after_idle(lambda: self._closing or self._relayout_all())
        # 启动即核对在线激活码吊销状态（后台线程，非阻塞；撤销即时生效，方向 A）
        # v2.3.92：改成「worker 只塞队列 + 主线程 after 轮询」。原先 worker 线程直接调
        # self.root.after() —— Tk 不是线程安全的，after() 内部会走 tk.createcommand，
        # 在非主线程上偶发 RuntimeError: main thread is not in main loop（实测能随机把
        # 后续用例的 setup 打成 ERROR）。现在 after() 只从主线程调用。
        self._start_launch_check()

    # ------------------------------------------------------------ 授权：启动一次性核验（授权 + 版本）
    def _start_launch_check(self):
        """起后台线程做启动核验，并在主线程挂一个队列轮询器（幂等，可多次调用）。"""
        if getattr(self, "_launch_q", None) is None:
            self._launch_q = queue.Queue()
        threading.Thread(target=self._refresh_license_at_launch, daemon=True).start()
        if not getattr(self, "_launch_polling", False):
            self._launch_polling = True
            try:
                self.root.after(250, self._poll_launch_check)
            except Exception:
                self._launch_polling = False

    def _poll_launch_check(self):
        """主线程轮询：把后台核验结果取出来交给 UI 线程处理（5 态状态机）。

        只有拿到结果才停止轮询 —— 后台失败（断网/异常）时线程直接 return，
        这里不会卡死，也不会漏掉「就是没有结果」这件事。
        """
        if getattr(self, "_closing", False):
            self._launch_polling = False
            return
        q = getattr(self, "_launch_q", None)
        got = None
        if q is not None:
            try:
                got = q.get_nowait()
            except queue.Empty:
                got = None
        if got is not None:
            self._launch_polling = False
            try:
                self._on_launch_check(got)
            except Exception:
                pass
            return
        try:
            self.root.after(250, self._poll_launch_check)
        except Exception:
            self._launch_polling = False

    def _refresh_license_at_launch(self):
        """后台线程：启动即向中台核对授权 + 版本（fail-soft，不阻塞首屏）。

        ⚠️ 本方法跑在 worker 线程上，**只许做网络与计算，绝不碰 Tk**
        （不调 after、不读控件、不弹窗）；结果塞队列，由主线程的
        ``_poll_launch_check`` 取走。
        """
        try:
            res = lic.refresh_license_on_launch()
        except Exception:
            return
        q = getattr(self, "_launch_q", None)
        if q is None:
            return
        q.put(res)

    def _on_launch_check(self, res):
        """启动核验结果分状态处理（5 态状态机）。

        - locked（吊销/过期）→ 提醒 + 激活窗口（合并升级/激活窗，含激活码输入框）
        - network_unreachable → 软提示「未能联网」，不锁、可继续用
        - 有更新且未失效 → 软横幅（可关，记住忽略版本）
        - 有效 + 最新 → 真静默（什么都不做）
        """
        tr = self.tr
        version = res.get("version") or {}
        if res.get("locked"):
            self._set_status(tr("gate_revoked"), "err")
            self._set_bar("err")
            # 🔴 必须先刷新商业区（老板 2026-10-25 报）：状态行与 Upgrade 按钮是在
            # 界面搭建时按**本地缓存**渲染的（那时心跳还没跑），不刷新就会一直显示
            # "Licensed · unlimited checks"，而实际码已被后台停用。刷新后状态行变成
            # "已过期或已撤销"，Upgrade 按钮也重新出现（客户有再次购买的入口）。
            self._refresh_commercial()
            # 直接打开合并的升级/激活窗（含激活码输入框）
            self._upgrade_dialog()
            return
        if res.get("network_unreachable"):
            self._set_status(tr("launch_offline"), "warn")
        if version.get("update_available") and not res.get("locked"):
            self._show_update_banner(version.get("latest_version"),
                                     version.get("download_url"))

    # ------------------------------------------------------------ 启动版本更新软横幅
    def _show_update_banner(self, latest, url):
        """启动发现新版本：非模态软横幅（可关，记住忽略版本）。"""
        if not latest:
            return
        # 已忽略过这个版本 → 不再烦
        if self._dismissed_version() == latest:
            return
        tr = self.tr
        try:
            win = tk.Toplevel(self.root)
            win.withdraw()                 # 先隐藏，定位后再显示，避免闪一下
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            win.configure(bg=theme.SURFACE)
            pad = 12
            tk.Label(win, text="⬆  " + tr("upd_title"), bg=theme.SURFACE,
                     fg=theme.PRIMARY, font=self.F["F_LABEL"]).pack(
                side="top", anchor="w", padx=pad, pady=(pad, 2))
            body = tr("upd_body") % (latest, self.version)
            tk.Label(win, text=body, bg=theme.SURFACE, fg=theme.TEXT_2,
                     font=self.F["F_SUBTITLE"], anchor="w", justify="left",
                     wraplength=360).pack(side="top", anchor="w", padx=pad, pady=(0, 8))
            row = tk.Frame(win, bg=theme.SURFACE)
            row.pack(side="top", anchor="w", padx=pad, pady=(0, pad))
            RoundButton(row, text=tr("upd_btn"), command=lambda: self._open_update(url),
                        style="primary", font=self.F["F_BTN_S"], height=30,
                        padx=14).pack(side="left")
            tk.Button(row, text=tr("upd_later"), bg=theme.SURFACE,
                      fg=theme.TEXT_3, font=self.F["F_FOOT"], bd=0,
                      command=lambda: self._dismiss_update(win, latest)).pack(
                side="left", padx=(10, 0))
            # 定位到主窗顶部居中
            self.root.update_idletasks()
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            rw = self.root.winfo_width()
            win.update_idletasks()
            ww = win.winfo_reqwidth()
            win.geometry("+%d+%d" % (rx + max(0, (rw - ww) // 2), ry + 8))
            win.deiconify()
        except Exception:
            pass

    def _open_update(self, url):
        try:
            webbrowser.open(url or i18n.BRAND_BUY_URL)
        except Exception:
            pass

    def _dismiss_update(self, win, latest):
        try:
            self._set_dismissed_version(latest)
        except Exception:
            pass
        try:
            win.destroy()
        except Exception:
            pass

    def _dismissed_version(self):
        try:
            p = os.path.join(os.path.expanduser("~"), ".paperformatpro",
                             "update_skip.json")
            if os.path.isfile(p):
                import json as _json
                with open(p, "r", encoding="utf-8") as fh:
                    return _json.load(fh).get("dismissed_version")
        except Exception:
            pass
        return None

    def _set_dismissed_version(self, v):
        try:
            import json as _json
            p = os.path.join(os.path.expanduser("~"), ".paperformatpro",
                             "update_skip.json")
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "w", encoding="utf-8") as fh:
                _json.dump({"dismissed_version": v}, fh)
        except Exception:
            pass

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
        """顶栏：品牌 + 标语 + 隐私声明（左）／语言（Chinese / 英文） · About · Help（右）。

        与前版的差别（都是为背景服务）：
          ① 顶栏仍画在背景画布上（``create_window``，不占 pack 空间）—— 背景素材的
             天际线正好压在顶栏这条带上，整条不透明框架盖住就等于白放了一张图；
          ② **删掉顶栏下的通栏细线**：参考稿里没有这条线，留着会横穿天际线；
          ③ 按设计稿补上**品牌 mark**（``widgets.BrandMark`` 加载 ``iconpath.find_icon()``
             取到的 PaperFormat Pro 新版圆形徽章 PNG，与窗口/任务栏图标同源）
             与隐私声明前的盾牌小徽标；语言选择器显示目标语言名（英文界面显 Chinese、中文界面显 英文），悬停提示仅切换界面语言。
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
        # v2.3.17：顶部隐私提示改用用户准备的「盾牌+电脑」PNG 图标。
        IconBadge(privacy, "thesis_local_privacy", size=theme.BADGE_SIZE,
                  bg=theme.BG).pack(side="left")
        tk.Label(privacy, text=self.tr("privacy_line"), bg=theme.BG, fg=theme.TEXT_2,
                 font=F["F_SMALL_B"]).pack(side="left", padx=(8, 0))

        # 右侧导航：自右向左 pack
        nav = tk.Frame(cv, bg=theme.BG)
        self._link(nav, self.tr("link_help"),
                   self._help_dialog).pack(side="right")
        self._nav_sep(nav)
        # v2.3.60：反馈入口紧挨 Help（老板 2026-10-04 定，低摩擦、易发现）
        self._link(nav, self.tr("link_feedback"),
                   self._feedback_dialog).pack(side="right")
        self._nav_sep(nav)
        self._link(nav, self.tr("link_about"),
                   self._about_dialog).pack(side="right")
        self._nav_sep(nav)
        self._link(nav, self.tr("lang_button"),
                   lambda: self._set_lang(i18n.other_lang(self.lang)),
                   fg=theme.TEXT,
                   tooltip=self.tr("lang_button_tip")).pack(side="right")

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
        """页脚：极简，最底部只留官网 —— 邮箱已收进 Help / About。

        v2.3.15：状态栏左侧已是「PaperFormat Pro · 版本」、右侧是就绪隐私提示，
        页脚进一步收敛为"只剩官网"一行小字，整页更有商业成品感（不再像开发工具）。
        品牌口径铁律：只出现官网与邮箱，绝不出现微信 / 公众号 / 小程序。
        """
        F = self.F
        footer = tk.Frame(self.root, bg=theme.BG)
        footer.pack(side="bottom", fill="x", padx=theme.PAGE_PAD,
                    pady=(0, theme.FOOTER_PAD_Y))
        site = tk.Label(footer, text=self.tr("footer_site"), bg=theme.BG,
                        fg=theme.TEXT_3, font=F["F_FOOT"], cursor="hand2")
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
        IconBadge(head, "doc", size=theme.BADGE_SIZE + 6,
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

        # ① 论文（必选）—— 页面视觉焦点：实线投放区（压扁版，老板 2026-09-26 嫌太高）
        self._field_label(parent, "row_paper_title", required=True)
        self._helper(parent, "row_paper_helper")
        drop = RoundCard(parent, radius=14, fill=theme.PRIMARY_SOFT,
                         border=theme.BORDER, shadow=False, dashed=False,
                         padx=16, pady=10, min_height=78)
        self._extra_cards.append(drop)
        drop.pack(fill="x", pady=(4, 0))
        self._paper_drop = drop

        # 居中内容列（v2.3.26：去掉手绘云朵上传图标，避免观感廉价；纯文字引导更干净、更专业）
        vcol = tk.Frame(drop.inner, bg=theme.PRIMARY_SOFT)
        vcol.pack(expand=True)

        self._paper_name = tk.Label(vcol, text=self.tr("btn_select_doc"),
                                    bg=theme.PRIMARY_SOFT, fg=theme.PRIMARY,
                                    font=F["F_BODY"], cursor="hand2", anchor="center")
        self._paper_name.pack(fill="x", pady=(2, 0))
        self._drop_hint = tk.Label(vcol, text=self.tr("drop_hint"),
                                   bg=theme.PRIMARY_SOFT, fg=theme.TEXT_2,
                                   font=F["F_HELP"], cursor="hand2", anchor="center")
        self._drop_hint.pack(fill="x", pady=(2, 0))
        self._paper_sub = tk.Label(vcol, text=self.tr("paper_support"),
                                   bg=theme.PRIMARY_SOFT, fg=theme.TEXT_3,
                                   font=F["F_HELP"], cursor="hand2", anchor="center")
        self._paper_sub.pack(fill="x", pady=(1, 0))
        # v2.3.19：论文投放区的「✕」清除按钮（默认隐藏，选中论文后显示，置于卡片右上角）。
        # 用 place 固定右上角，不影响 vcol 居中布局；返回 "break" 阻止冒泡到整卡（避免同时触发选文件）。
        self._paper_clear = tk.Label(drop.inner, text="✕", bg=theme.PRIMARY_SOFT,
                                     fg=theme.TEXT_3, font=F["F_LABEL"], cursor="hand2")
        self._paper_clear.place(relx=1.0, rely=0.0, x=-10, y=8, anchor="ne")
        self._paper_clear.bind("<Button-1>",
                               lambda e: (self._clear_input("paper"), "break")[-1])
        self._paper_clear.bind("<Enter>", lambda e: self._paper_clear.config(fg=theme.ERROR))
        self._paper_clear.bind("<Leave>", lambda e: self._paper_clear.config(fg=theme.TEXT_3))
        for w in (drop.inner, vcol, self._paper_name, self._drop_hint,
                  self._paper_sub):
            w.bind("<Button-1>", lambda e: self._pick_docx())
            w.bind("<Enter>", lambda e: self._set_drop_hover(True))
            w.bind("<Leave>", lambda e: self._set_drop_hover(False))

        self._spacer(parent)
        self._hairline(parent)

        # ② 格式来源（必选）：6 选 1 互斥
        self._build_source_selector(parent)

        self._spacer(parent)
        self._hairline(parent)

        # ③ 高级选项折叠区：学校模板 / 自定义（选内置规范时整片置灰）
        self._build_advanced(parent)
        # 初始化高级区可用态（与当前来源一致：默认内置规范 → 置灰折叠）
        self._apply_source_mode(self.source_mode.get(), refresh=False)

        # ④ 导出 AI 模板：属于「导入配置」的一部分，放在配置区下方（v2.3.19 按钮重组：
        # 左卡=输入/配置，右卡=操作/结果）。它不是主操作，用次级按钮，不抢视觉主角。
        self._spacer(parent)
        self._export_btn = RoundButton(parent, self.tr("btn_export_ai"),
                                       self._export_template, style="secondary",
                                       font=F["F_BTN_S"], height=42)
        self._export_btn.pack(fill="x")
        # 它只在来源 = AI 导入时才有意义（6/7 项互斥），建完按当前来源收一次口
        self._sync_export_visible()

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
            # v2.3.19：Citation style 的引号图标视觉偏重，单独缩小一点，
            # 让它跟其它字段行图标（paper/bank/code/gear）更协调。
            badge_size = (theme.BADGE_SIZE - 4
                          if key == "source_title" else theme.BADGE_SIZE)
            IconBadge(row, _FIELD_ICONS[key], size=badge_size,
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

    def _build_source_selector(self, parent):
        """格式来源（必选）：6 选 1 互斥。

        前 4 项带内置规范（spec_key 非空，自动套用其规则）；后 2 项
        （学校模板 / 自定义）走无规范基线，页面与引用完全由用户素材决定。
        选内置规范时高级区整片置灰并折叠；选学校模板 / 自定义时高级区自动展开。
        """
        F = self.F
        self._field_label(parent, "source_title", required=True)
        # 6 选 1 互斥来源表：(spec_key 或 None, i18n 显示键)
        self._source_map = [
            ("APA-Student", "src_apa_student"),
            ("APA-Professional", "src_apa_pro"),
            ("MLA", "src_mla"),
            ("IEEE-Conference", "src_ieee"),
            ("IEEE-Journal", "src_ieee_journal"),
            (None, "src_school"),
            (None, "src_manual"),
            (None, "src_ai"),
        ]
        disp_values = [self.tr(k) for _, k in self._source_map]
        self._source_disp = tk.StringVar(value=self.tr(self.source_mode.get()))
        self._source_cb = ttk.Combobox(parent, textvariable=self._source_disp,
                                       values=disp_values, state="readonly",
                                       font=F["F_BODY"], width=30)
        self._source_cb.pack(fill="x", pady=(0, theme.ROW_GAP))
        self._source_cb.bind("<<ComboboxSelected>>", self._on_source_pick)
        self._helper(parent, "source_helper")
        # 下拉**弹出列表**的字体不受 widget 的 font 影响，得单独喂给 option 库，
        # 否则弹出来还是 Tk 默认字体（观感直接掉档）。
        try:
            self.root.option_add("*TCombobox*Listbox.font", F["F_BODY"])
        except Exception:
            pass
        # 来源变化会让上一次的体检结论失效（否则结论与所选来源对不上）。
        # 但**正常交互路径**（_on_source_pick）自己会在同一道重绘闸门里完成
        # 「折叠/展开 + 作废旧结论 + 文本刷新 + 重测量」，这里若再跑一遍
        # _on_input_changed，同一次选择就会触发两轮重排、两次上屏 →
        # 中间帧出现「卡片被撕开 / 同一行字出现两遍」（花屏同类现象）。
        # 故用 _applying_source 重入标志挡住重复的一遍；外部程序化改 source_mode
        # 仍能享受这条兜底刷新。
        try:
            self.source_mode.trace_add("write", lambda *a: self._on_source_trace())
        except Exception:
            pass

    def _on_source_trace(self):
        """source_mode 变化的兜底刷新（正常交互路径已被 _applying_source 挡住）。"""
        if getattr(self, "_applying_source", False):
            return
        self._on_input_changed()

    def _on_source_pick(self, _e=None):
        """下拉展示名 → 来源模式键（再驱动 spec_key 与高级区可用态）。"""
        try:
            disp = self._source_disp.get()
        except Exception:
            return
        mode_key = None
        for _, k in self._source_map:
            if self.tr(k) == disp:
                mode_key = k
                break
        if mode_key is None:
            return
        # 置重入标志：.set() 会触发 trace，但完整收口交给下面的 _apply_source_mode
        # （一次操作只应该重排/上屏一次）。
        self._applying_source = True
        try:
            self.source_mode.set(mode_key)
            self._apply_source_mode(mode_key, refresh=True)
        finally:
            self._applying_source = False

    def _apply_source_mode(self, mode_key, refresh=True):
        """按来源模式设置 spec_key + 高级区可用 / 展开态 + 显示哪一块子面板。

        - 内置规范（前 4 项）：spec_key 落盘、高级区整片置灰并折叠；
        - 学校模板 / 自定义：spec_key 清空、高级区展开且可用，**只展开对应那一块**
          （v2.3.65 起两块互斥显示，不再同时摊开）。

        refresh=False 用于初始构建期（不触发 _relayout_all，避免访问尚未建好的控件）；
        用户交互时 refresh=True，再做完整重测量与刷新。
        """
        spec_key = None
        for sk, k in self._source_map:
            if k == mode_key:
                spec_key = sk
                break
        # 单一真相：以"应用哪一个来源"为准，顺手把 source_mode 对齐。
        # 旧版只有 _on_source_pick 会写 source_mode，程序化调用（测试 / 恢复状态）时
        # 两者会不一致 —— _build_target 与「导出按钮显隐」就可能读到旧来源。
        if self.source_mode.get() != mode_key:
            was = getattr(self, "_applying_source", False)
            self._applying_source = True
            try:
                self.source_mode.set(mode_key)
            finally:
                self._applying_source = was
        # 下拉框显示的文案也得跟上（它和 source_mode 是"显示名 ↔ 键"两套值）。
        # 真实点击时由 _on_source_pick 从选中项反推，程序化切换（恢复状态 / 测试）
        # 不补这一步，就会出现"面板换成 AI 导入、下拉还写着 APA"的错位。
        try:
            disp = self.tr(mode_key)
            if getattr(self, "_source_disp", None) is not None and self._source_disp.get() != disp:
                self._source_disp.set(disp)
        except Exception:
            pass
        self.spec_key.set(spec_key if spec_key is not None else "")
        self._adv_enabled = spec_key is None
        is_school = mode_key == "src_school"
        is_manual = mode_key == "src_manual"
        is_ai = mode_key == "src_ai"

        def _flip():
            """折叠/展开高级区并刷新置灰态（纯显隐，不含测量）。"""
            if spec_key is not None:
                if getattr(self, "_adv_open", False):
                    self._adv_body.pack_forget()
                    self._adv_chev.config(text="▸")
                    self._adv_open = False
            else:
                if not getattr(self, "_adv_open", False):
                    self._adv_body.pack(fill="x", pady=(theme.SECTION_GAP, 0))
                    self._adv_chev.config(text="▾")
                    self._adv_open = True
            # v2.3.65：来源是 7 选 1、**互斥**的 → 子面板也只显示被选中的那一块。
            # 旧版学校模板与自定义同时摊开，客户选了自定义还看见「University template」，
            # 自然要问"手动输入为什么还要学校模板"。三块（学校模板 / 手填问卷 / AI 导入）
            # 一一对应三个来源，同一时刻最多露一块。
            for frame, show, pad in (
                    (getattr(self, "_school_frame", None), is_school, (4, 10)),
                    (getattr(self, "_manual_frame", None), is_manual, (4, 10)),
                    (getattr(self, "_ai_frame", None), is_ai, (4, 2))):
                if frame is None:
                    continue
                if show:
                    frame.pack(fill="x", pady=pad)
                else:
                    frame.pack_forget()
            # 「Export AI template」只属于 AI 导入这条路（老板：只在 AI 导入时显示）
            self._sync_export_visible()
            self._paint_advanced_state()

        if not refresh:
            # 初始构建期：尚未建好全部控件，只做显隐，不触发测量。
            _flip()
            return
        # 整段（显隐 + 作废旧结论 + 文本刷新 + 重测量）关在同一道重绘闸门里：
        # 切来源会折叠/展开高级区，与 _toggle_advanced 属同一类花屏场景
        # （实测中间帧会出现「卡片被撕开 / 同一行字出现两遍」）。
        with self._paused_paint():
            _flip()
            # 来源变了 → 上一次的体检结论不再可信（与旧 trace 行为一致，但只做一次）
            if getattr(self, "_checked_path", ""):
                try:
                    self._invalidate_results()
                except Exception:
                    pass
            self._refresh_rows()      # 内部 _sync_ui → _relayout_all（可嵌套闸门）
        # 出闸门后一次完整刷新，最终画面立刻上屏（不等事件循环）。
        self._flush_frame()

    def _sync_export_visible(self):
        """「Export AI template」按钮只在来源 = Import configuration (AI) 时出现。

        它属于 AI 导入这条路的配套设施（导出模板 → 交给 AI 填 → 再导回来），
        选内置规范 / 学校模板 / 手填问卷时留在界面上只会干扰。
        """
        btn = getattr(self, "_export_btn", None)
        if btn is None:
            return
        try:
            shown = btn.winfo_manager() == "pack"
        except Exception:
            # 切换语言会整体重建界面：此刻 _export_btn 还指着已被销毁的旧控件
            # （winfo_manager 会抛 TclError）。等新按钮建好后 _build_left 末尾再收口。
            return
        if self.source_mode.get() == "src_ai":
            if not shown:
                btn.pack(fill="x")
        elif shown:
            btn.pack_forget()

    def _paint_advanced_state(self):
        """高级区可用态外观：可用=正常色，禁用=灰色（标题 + 箭头）。"""
        if self._adv_enabled:
            self._adv_title.config(fg=theme.TEXT)
            self._adv_chev.config(fg=theme.SECONDARY)
        else:
            self._adv_title.config(fg=theme.TEXT_3)
            self._adv_chev.config(fg=theme.TEXT_3)

    def _select_field(self, parent, placeholder_key: str, cmd, on_clear=None):
        """「下拉框外观 + 打开文件行为」的控件。

        规格画的是 ``Choose a template ▾`` 这样的下拉；我们的真实动作是打开文件
        对话框（本机离线读模板），所以外观照规格做、行为照实际来：整行可点。
        v2.3.19：新增 on_clear 回调——传入则在行尾渲染一个小「✕」清除按钮
        （默认隐藏，选中文件后由 ``_refresh_rows`` 显示）。点 ✕ 即清除该项选择，
        导入错了不必重开软件，点叉号即可重新导入。返回 ``(lbl, clear_btn)``，
        clear_btn 未启用时为 ``None``。
        """
        F = self.F
        box = tk.Frame(parent, bg=theme.BG, highlightthickness=1,
                       highlightbackground=theme.FIELD_BORDER,
                       highlightcolor=theme.PRIMARY)
        box.pack(fill="x")
        inner = tk.Frame(box, bg=theme.SURFACE)
        inner.pack(fill="x", padx=10, pady=3)
        # v2.3.33：根治「长文件名盖住 ✕」。拆成左右两框——
        #   左框（expand）吸收剩余宽度，文件名 label 填在其中、文字被框宽裁住，
        #   物理上不可能溢出到右框；右框放「✕ / ⌄」固定控件（自然宽度，永远在右侧）。
        # 这样无论文件名多长都不会盖住清除按钮；下面的动态省略号只是锦上添花。
        left = tk.Frame(inner, bg=theme.SURFACE)
        left.pack(side="left", fill="x")
        right = tk.Frame(inner, bg=theme.SURFACE)
        right.pack(side="right")
        lbl = tk.Label(left, text=self.tr(placeholder_key), bg=theme.SURFACE,
                       fg=theme.TEXT_3, font=F["F_BODY"], anchor="w",
                       cursor="hand2")
        lbl.pack(side="left", fill="x", anchor="w")
        # v2.3.33：存完整名，按左框真实像素宽动态截断加省略号（随窗口缩放 / Advanced
        # 展开自动重算），保证文件名永远不溢出左框、更不盖住右侧 ✕/⌄。
        lbl._full_text = ""
        lbl.bind("<Configure>", lambda e: self._fit_field_label(lbl))
        # v2.3.19：小「✕」清除按钮（行尾，默认隐藏；选中后由 _refresh_rows 显示）。
        # v2.3.29：做成带边框的小芯片、放在下拉箭头「⌄」左侧，hover 高亮——此前
        # bg=SURFACE 与行同色、又挤在箭头右侧，实测几乎看不见 → 用户报"导入后没有叉号"。
        clear_btn = None
        if on_clear:
            clear_btn = tk.Label(right, text="✕", bg=theme.SURFACE, fg=theme.TEXT_2,
                                 font=F["F_HELP"], cursor="hand2",
                                 highlightbackground=theme.BORDER, highlightthickness=1,
                                 padx=3, pady=0)
            clear_btn.pack(side="left", padx=(4, 0))
            clear_btn.pack_forget()
            clear_btn._chevron = None   # 占位，下面绑定到真正的下拉箭头
            # 返回 "break" 阻止事件冒泡到整行（否则会同时触发「打开文件」）
            clear_btn.bind("<Button-1>", lambda e: (on_clear(), "break")[-1])
            clear_btn.bind("<Enter>",
                           lambda e: clear_btn.config(bg=theme.PRIMARY_SOFT,
                                                      fg=theme.ERROR))
            clear_btn.bind("<Leave>",
                           lambda e: clear_btn.config(bg=theme.SURFACE,
                                                      fg=theme.TEXT_2))
        chevron = tk.Label(right, text="⌄", bg=theme.SURFACE, fg=theme.SECONDARY,
                           font=F["F_BODY"], cursor="hand2")
        chevron.pack(side="left")
        if clear_btn is not None:
            clear_btn._chevron = chevron

        def _click(_e=None):
            cmd()

        for w in (box, inner, lbl):
            w.bind("<Button-1>", _click)
        return lbl, clear_btn

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

        bar = RoundCard(parent, radius=theme.CARD_RADIUS, fill=theme.SURFACE,
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
        self._adv_title = tk.Label(hdr, text=self.tr("adv_label"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_LABEL"], cursor="hand2")
        self._adv_title.pack(side="left", padx=(8, 0))

        self._adv_body = tk.Frame(bar.inner, bg=theme.SURFACE)

        # —— 高级格式设置（默认折叠，渐进式披露）——
        # ① 学校模板：导入 .docx 模板，自动读其真实样式；子下拉选引用风格（默认不检查）。
        # ② 自定义：手填问卷 / 导入 AI 填表 JSON。
        # 选内置规范时整片置灰并折叠（_apply_source_mode 控制）；选学校模板 / 自定义时展开可用。
        self._build_adv_school(self._adv_body)
        self._build_adv_custom(self._adv_body)

        for w in (hdr, self._adv_chev, self._adv_title):
            w.bind("<Button-1>", lambda e: self._toggle_advanced())

    def _build_adv_school(self, parent):
        """高级区 ①：学校模板面板（两行）。

        行1 导入 .docx 模板文件（读取真实样式）；行2 引用 / 参考文献风格子下拉，
        默认「不检查引用」。来源为学校模板时由 _apply_source_mode 自动展开。
        """
        tr = self.tr
        F = self.F
        sec = tk.Frame(parent, bg=theme.SURFACE)
        sec.pack(fill="x", pady=(4, 10))
        self._school_frame = sec
        h = tk.Frame(sec, bg=theme.SURFACE)
        h.pack(fill="x")
        # v2.3.62 重做面板时误用了 "doc"（→ your_documents 文档图标，与左卡大标题撞图标）；
        # 学校模板行按设计稿一直用 "bank"（→ university_template 圆顶大楼图标）。
        IconBadge(h, "bank", size=theme.BADGE_SIZE, bg=theme.SURFACE).pack(side="left")
        tk.Label(h, text=tr("adv_school_title"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_SUBTITLE"], anchor="w").pack(side="left", padx=(8, 0))
        tk.Label(sec, text=tr("adv_school_desc"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_HELP"], anchor="w", justify="left").pack(
            fill="x", pady=(1, 6))
        # 行1：导入模板文件
        self._tpl_name, self._tpl_clear = self._select_field(
            sec, "row_template_placeholder", self._pick_template,
            on_clear=lambda: self._clear_input("template"))
        # 行2：引用 / 参考文献风格
        r2 = tk.Frame(sec, bg=theme.SURFACE)
        r2.pack(fill="x", pady=(6, 0))
        tk.Label(r2, text=tr("school_ref_title"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_HELP"], anchor="w").pack(side="left")
        self._school_ref_cb = ttk.Combobox(
            r2, textvariable=self.school_ref_style,
            values=[tr("school_ref_none"), "APA", "MLA", "IEEE"],
            state="readonly", font=F["F_BODY"], width=22)
        self._school_ref_cb.pack(side="right")
        # 默认「不检查引用」（用 ComboboxSelected 绑定而非 trace_add，避免构建期
        # 程序化 set 误触发 _refresh_rows 访问尚未建好的控件）
        self.school_ref_style.set(tr("school_ref_none"))
        try:
            self.root.option_add("*TCombobox*Listbox.font", F["F_BODY"])
        except Exception:
            pass
        # 引用风格改变 → 旧结论失效 + 刷新（仅用户交互时触发）
        self._school_ref_cb.bind(
            "<<ComboboxSelected>>", lambda e: self._on_input_changed())

    def _build_adv_custom(self, parent):
        """高级区 ②③：手填问卷 / AI 导入——**各自是一个独立来源**，互斥显示。

        v2.3.65：来源下拉是 7 选 1、互斥的。旧版把「手填问卷 + AI 导入」塞进同一个
        Custom 选项里，客户选了 Custom 反而同时看见学校模板、问卷、AI 三样入口，
        逻辑上说不通（老板原话：「手动输入为什么还要学校模板」「这七个选项是互斥的」）。
        现在拆成两块独立面板，由 ``_apply_source_mode`` 按来源只显示对应的那一块：

        - 学校模板（``_build_adv_school``）：模板行 + 参考文献格式下拉（**只有这项有**）
        - 手填问卷（本函数前半）：问卷行
        - AI 导入（本函数后半）：导入配置行（+ 左卡底部的「Export AI template」）

        两行都保留 v2.3.64 的两处修正：不再罗列格式明细、按钮先入队避免被压窄裁字。
        """
        tr = self.tr
        F = self.F

        # —— ② 手填问卷 ——
        sec = tk.Frame(parent, bg=theme.SURFACE)
        sec.pack(fill="x", pady=(4, 10))
        self._manual_frame = sec
        r1 = tk.Frame(sec, bg=theme.SURFACE)
        r1.pack(fill="x")
        # 按钮**先**入队。Tk 的 pack 在横向空间不够时是"压缩最后入队的控件"，
        # 而 RoundButton 是自绘 canvas：一旦被压窄，居中绘制的文字就从两侧被裁
        # （真机截图实测 "Import" 只剩 "mpor"）。先让按钮占住自己的宽度，
        # 剩余空间再分给左侧文字列 —— 无论文案多长，按钮字都完整。
        RoundButton(r1, text=tr("btn_questionnaire"),
                    command=self._open_questionnaire, style="secondary",
                    font=F["F_BTN_S"], height=28, padx=12).pack(
                        side="right", padx=(10, 0))
        c1 = tk.Frame(r1, bg=theme.SURFACE)
        c1.pack(side="left", fill="x", expand=True)
        IconBadge(c1, "questionnaire", size=theme.BADGE_SIZE, bg=theme.SURFACE).pack(side="left")
        tk.Label(c1, text=tr("row_questionnaire_title"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_SUBTITLE"], anchor="w").pack(
                     side="left", padx=(8, 0), fill="x", expand=True)
        # 只留状态行（已填项数 / 未填写），明细留给复核弹窗去看
        self._q_name = tk.Label(c1, text="", bg=theme.SURFACE, fg=theme.TEXT_3,
                                font=F["F_HELP"], anchor="w", justify="left")
        self._q_name.pack(fill="x", pady=(2, 0))
        _auto_wrap(self._q_name, c1, 0)
        # v2.3.78：最近填过的问卷 = 可点即填入的「模板」下拉
        r_rec = tk.Frame(sec, bg=theme.SURFACE)
        r_rec.pack(fill="x", pady=(4, 0))
        tk.Label(r_rec, text=tr("recents_manual_label"), bg=theme.SURFACE,
                 fg=theme.TEXT_3, font=F["F_HELP"], anchor="w").pack(
                     side="left", padx=(0, 6))
        self._manual_recent = ttk.Combobox(r_rec, state="readonly", height=6,
                                           font=F["F_HELP"], width=20)
        self._manual_recent.pack(side="left", fill="x", expand=True)
        self._manual_recent.bind("<<ComboboxSelected>>", self._on_manual_recent_pick)

        # —— ③ AI 导入配置 ——
        sec2 = tk.Frame(parent, bg=theme.SURFACE)
        sec2.pack(fill="x", pady=(4, 2))
        self._ai_frame = sec2
        r2 = tk.Frame(sec2, bg=theme.SURFACE)
        r2.pack(fill="x")
        RoundButton(r2, text=tr("btn_ai_import"),
                    command=self._pick_ai_config, style="secondary",
                    font=F["F_BTN_S"], height=28, padx=12).pack(
                        side="right", padx=(10, 0))
        # ✕ 清除按钮排在上面那个按钮的左侧（后入队的 side="right" 更靠内），
        # 选中配置后由 _refresh_rows 显示
        self._ai_clear = tk.Label(r2, text="✕", bg=theme.SURFACE, fg=theme.TEXT_3,
                                  font=F["F_HELP"], cursor="hand2")
        self._ai_clear.pack(side="right", padx=(0, 4))
        self._ai_clear.pack_forget()
        self._ai_clear.bind("<Button-1>", lambda e: (self._clear_input("ai"), "break")[-1])
        self._ai_clear.bind("<Enter>", lambda e: self._ai_clear.config(fg=theme.ERROR))
        self._ai_clear.bind("<Leave>", lambda e: self._ai_clear.config(fg=theme.TEXT_3))
        c2 = tk.Frame(r2, bg=theme.SURFACE)
        c2.pack(side="left", fill="x", expand=True)
        IconBadge(c2, "import_config", size=theme.BADGE_SIZE, bg=theme.SURFACE).pack(side="left")
        tk.Label(c2, text=tr("row_ai_title"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_SUBTITLE"], anchor="w").pack(
                     side="left", padx=(8, 0), fill="x", expand=True)
        self._ai_name = tk.Label(c2, text="", bg=theme.SURFACE, fg=theme.TEXT_3,
                                 font=F["F_HELP"], anchor="w", justify="left")
        self._ai_name.pack(fill="x", pady=(2, 0))
        _auto_wrap(self._ai_name, c2, 0)
        # v2.3.78：最近导入的 AI 配置文件 = 可点即重新导入的下拉
        r_rec2 = tk.Frame(sec2, bg=theme.SURFACE)
        r_rec2.pack(fill="x", pady=(4, 0))
        tk.Label(r_rec2, text=tr("recents_ai_label"), bg=theme.SURFACE,
                 fg=theme.TEXT_3, font=F["F_HELP"], anchor="w").pack(
                     side="left", padx=(0, 6))
        self._ai_recent = ttk.Combobox(r_rec2, state="readonly", height=6,
                                       font=F["F_HELP"], width=20)
        self._ai_recent.pack(side="left", fill="x", expand=True)
        self._ai_recent.bind("<<ComboboxSelected>>", self._on_ai_recent_pick)

        # 下拉建好后立即用磁盘上的 recents 填充（语言切换/记忆变更后都会再刷）
        self._refresh_recent_combos()

    def _toggle_advanced(self):
        """展开 / 收起 Advanced options（返回展开后的状态，便于测试断言）。"""
        # 选内置规范时高级区整片置灰，用户点击无效（来源切回学校/自定义才解锁）
        if not getattr(self, "_adv_enabled", True):
            return self._adv_open
        # v2.3.32：整段（显隐 + 重新测量）都关在闸门里 —— 实测展开/收起各有
        # 数百毫秒的中间帧，屏幕上会出现"同一行字出现两遍、卡片被撕开"的花屏。
        with self._paused_paint():
            self._adv_open = not self._adv_open
            try:
                if self._adv_open:
                    self._adv_body.pack(fill="x", pady=(theme.SECTION_GAP, 0))
                    self._adv_chev.config(text="▾")
                else:
                    self._adv_body.pack_forget()
                    self._adv_chev.config(text="▸")
            except Exception:
                pass
            # 展开/收起改变了左卡内容高度 —— 必须主动重新测量，
            # 否则新出现的那两行会被圆角矩形裁掉且滚不到（见 _relayout_all）。
            self._relayout_all()
        # 出闸门后补一次完整刷新，确保最终画面立刻上屏（不等事件循环）。
        self._flush_frame()
        return self._adv_open

    # ------------------------------------------------------------ 右栏
    def _build_right(self, parent):
        """右卡 Review & Fix：1-2-3 引导**常驻**，进度条按状态显隐，二者不互斥替换。

        重设计（2026-09-26 / 2.3.25 修正）：解决「右栏空白」+「右卡不稳定」两个问题。
        1-2-3 步骤图永远钉在右卡顶部（根治 v2.3.23 之前"导入后 1-2-3 掉到右下角"），
        进度条仅在已选论文后显示在 1-2-3 之下（空态隐藏，保持右栏干净、不挂无用步进器）；
        信任卡与商业区始终保留在底部。无论空态/已选论文/已出结果，顶部引导都固定在原地，
        符合"内容不折叠"的界面底线。
        """
        F = self.F
        head = tk.Frame(parent, bg=theme.SURFACE)
        head.pack(fill="x")
        trow = tk.Frame(head, bg=theme.SURFACE)
        trow.pack(side="left")
        IconBadge(trow, "sliders", size=theme.BADGE_SIZE + 6,
                  bg=theme.SURFACE).pack(side="left")
        tk.Label(trow, text=self.tr("right_title"), bg=theme.SURFACE,
                 fg=theme.TEXT, font=F["F_CARD_HDR"], anchor="w").pack(
                     side="left", padx=(10, 0))
        desc = tk.Label(parent, text=self.tr("right_desc"), bg=theme.SURFACE,
                        fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                        justify="left")
        desc.pack(fill="x", pady=(2, 0))
        _auto_wrap(desc, parent, theme.CARD_PAD_X)

        # v2.3.16：常驻「下一步」引导条——随状态机更新，明确告诉用户现在该做什么
        # （老板反馈"向导性不够好"）。放在右卡标题下方，比把说明散落各处更聚焦。
        self._guide_lbl = tk.Label(parent, text="", bg=theme.SURFACE_SOFT,
                                  fg=theme.PRIMARY, font=F["F_HELP"], anchor="w",
                                  justify="left")
        self._guide_lbl.pack(fill="x", pady=(8, 0))
        _auto_wrap(self._guide_lbl, parent, theme.CARD_PAD_X)

        self._spacer(parent)

        # 空状态：1-2-3 步骤引导（未选论文时显示）
        self._empty_state = tk.Frame(parent, bg=theme.SURFACE)
        self._empty_state.pack(fill="x", pady=(theme.SECTION_GAP, 0))
        tk.Label(self._empty_state, text=self.tr("how_title"), bg=theme.SURFACE,
                 fg=theme.TEXT, font=F["F_LABEL"], anchor="w").pack(fill="x")
        for i, (title, body) in enumerate(self.tr("how_steps"), 1):
            self._build_step_row(self._empty_state, i, title, body)

        # 极简进度条：默认隐藏，由 _sync_ui_state 在「已选论文」态才 pack（与
        # test_ui_redesign 约定：空态不挂步进器），避免首帧闪一下
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

        # v2.3.19 按钮重组：操作按钮组（Check / Fix / Save）放在右卡偏下位置。
        # 理由：用户点击时视线落在按钮「上方」的内容（结果/状态）上，按钮在下方才符合
        # 「先看结果 → 再操作」的动线；左卡=输入/配置，右卡=操作/结果。
        # 用固定容器 _ops_group 兜住操作按钮：_check_btn 常驻、_save_btn 初始 pack_forget()，
        # 在 set_visible(True) 时会 super().pack() 重排到「父容器末尾」——若直接挂在右卡
        # 主体 parent 上，重排后会掉到下方 _commercial_row（Trial + Upgrade）之下，复辟
        # P1-1 旧坑（次按钮被挤到 CTA 行下面）。故按钮只在 _ops_group 内重排，组本身位置
        # 固定，永不污染右卡主体 packing 顺序。
        self._ops_group = tk.Frame(parent, bg=theme.SURFACE)
        self._ops_group.pack(fill="x", pady=(0, 8))
        self._check_btn = RoundButton(self._ops_group, self.tr("btn_check"), self._run,
                                      style="primary", font=F["F_BTN"],
                                      height=50, icon="search")
        self._check_btn.pack(fill="x", pady=(0, 8))
        self._save_btn = RoundButton(self._ops_group, self.tr("btn_save_report"),
                                     self._save_report, style="secondary",
                                     font=F["F_BTN_S"], height=42)
        self._save_btn.pack_forget()

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
        # v2.3.27：高度 34 → 46。draw() 把步骤标签画到 y≈27 往下（底边≈41px），
        # 原 34px 画布装不下自己的标签 → 标签被裁、右卡重测时只按 34px 量高，
        # 导入论文后右卡"涨了但不够"、进度条标签压到状态行。46px 让画布高度与真实
        # 内容一致，右卡才能涨到正确高度、标签完整显示。
        self._progress = ProgressBar(parent, steps=3, bg=theme.SURFACE,
                                     font=F["F_HELP"], height=46)
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
        # v2.3.17：用用户准备的「盾牌+学士帽」PNG 替换原来的 shield 自绘小徽标，
        # 并去掉右侧冗余的单独 cap 图标（built_academic 图片已包含学士帽元素）。
        IconBadge(row, "built_academic", size=theme.BADGE_SIZE, bg=theme.SURFACE_SOFT,
                  fill=theme.PRIMARY_SOFT).pack(side="left")
        tk.Label(row, text=self.tr("trust_title"), bg=theme.SURFACE_SOFT,
                 fg=theme.TEXT_2, font=F["F_LABEL"],
                 anchor="w").pack(side="left", padx=(8, 0))
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
                                   fg=theme.TEXT_2, font=F["F_HELP"], anchor="w",
                                   justify="left")
        # fill+expand：状态文案吃掉 Upgrade 按钮之外的剩余宽度，按钮宽度由 pack 保证
        # 不被挤压（配合 _bind_commercial_wrap 的动态换行，见其 docstring）。
        self._trial_lbl.pack(side="left", fill="x", expand=True)
        self._upgrade_btn = RoundButton(
            row, text=self.tr("btn_upgrade"), command=self._upgrade_dialog,
            style="secondary", font=F["F_BTN_S"], height=28, padx=12)
        self._upgrade_btn.pack(side="right")
        self._bind_commercial_wrap()
        self._refresh_commercial()

    def _refresh_commercial(self):
        """按 license 的**只读**状态刷新状态行文案与升级按钮显隐（任何异常都退回中性文案）。

        核心分档只读 ``lic.status()``（与授权停用回归测试 ``test_revoked_state_sync`` 一致，
        该测试 monkeypatch 的就是 ``lic.status``）；四档授权（times/quarterly/annual/lifetime）
        的剩余次数 / 剩余天数 / 预警 / 耗尽等**数值细节**再补一份 ``lic.license_summary()``。

        规则：
        - revoked：显示撤销文案 + 露 Upgrade；
        - 未激活：试用剩余（limit - used）/ 已用完 + 露 Upgrade；
        - lifetime：已授权·不限次，隐藏 Upgrade；
        - times 次卡：显示剩余次数；≤1 次或已用完 → 露 Upgrade（加购）；
        - quarterly/annual/monthly 定时卡：显示到期日；≤7 天或已过期 → 露 Upgrade（续费/再买）。
        """
        tr = self.tr
        text = tr("trial_none")
        upgrade_visible = True
        try:
            st = lic.status()
            summ = lic.license_summary()  # 四档数值细节（剩余次数/天数/预警/耗尽）
            limit = int(st.get("trial_limit") or 1)
            if st.get("revoked"):
                text = tr("trial_revoked")
                upgrade_visible = True
            elif st.get("activated"):
                kind = st.get("kind") or lic.KIND_LIFETIME
                label = summ.get("kind_label") or tr("trial_licensed")
                if kind == lic.KIND_LIFETIME:
                    text = tr("trial_licensed")
                    upgrade_visible = False
                elif kind == lic.KIND_TIMES:
                    left = summ.get("uses_left")
                    if left is not None and left <= 0:
                        text = tr("lic_times_exhausted", label)
                        upgrade_visible = True
                    elif left is not None and left <= 1:
                        text = tr("lic_times_left_one", label, left)
                        upgrade_visible = True  # 低额度加购提示
                    else:
                        text = tr("lic_times_left", label, left) if left is not None else tr("trial_licensed")
                        upgrade_visible = False
                elif kind in (lic.KIND_QUARTERLY, lic.KIND_ANNUAL, "monthly"):
                    exp = _fmt_expiry(st.get("expires_at"))
                    if summ.get("exhausted"):
                        text = tr("lic_exp_expired", label)
                        upgrade_visible = True
                    elif summ.get("warn"):
                        text = tr("lic_exp_soon", label, exp) if exp else tr("lic_exp_soon_nd", label)
                        upgrade_visible = True  # 近到期软提醒（续费/再买）
                    else:
                        text = tr("lic_exp", label, exp) if exp else tr("lic_exp_nd", label)
                        upgrade_visible = False
                else:
                    text = tr("trial_licensed")
                    upgrade_visible = False
            elif not st.get("trial_fix_used"):
                left = max(0, limit)
                text = tr("trial_fix_one", left) if left == 1 else tr("trial_fix_many", left)
                upgrade_visible = True
            else:
                text = tr("trial_none")
                upgrade_visible = True
        except Exception:
            pass
        try:
            self._trial_lbl.config(text=text)
        except Exception:
            pass
        try:
            self._upgrade_btn.set_visible(upgrade_visible)
        except Exception:
            pass

    def _bind_commercial_wrap(self):
        """商业区状态文案动态换行：换行上限 = 行宽 − Upgrade 按钮的**适配宽** − 间隙。

        ``_trial_lbl`` 若干态文案很长：revoked 曾达 66 字符；四档里
        ``lic_exp_soon``「%s · expires %s (renew soon)」、``lic_exp_expired``
        「%s expired — renew or buy again」、``lic_times_exhausted``
        「%s used up — buy another pack」铺开也都有 ~40+ 字符。若不给 label
        收口 wraplength，``side="left"`` 的长文案按自然宽度请求，会把右侧
        ``side="right"`` 的 Upgrade 按钮挤窄 —— 按钮文字 "Upgrade to
        continue →" 被裁得只剩开头（2026-10-25 用户截图取证；到期/次数用完态同理）。

        修法：把 wraplength 钉到「行宽 − 按钮适配宽(``_fit_width``) − 间隙」，
        按钮宽度由 pack 保证永不被挤压，长文案在剩余宽度里自动换行。**必须按按钮
        适配宽预留，而非标签实占宽**——标签实占宽本身已含被挤掉的缺口，照它算会
        自我循环、按钮照旧被挤（首版踩过，已用真渲染冒烟取证）。按钮隐藏
        （lifetime 已授权态）时不预留，文案吃满整行。防重入 / 延后首帧 / 销毁取消
        沿用 ``_auto_wrap`` 的成熟套路。
        """
        lbl = self._trial_lbl
        row = self._commercial_row
        btn = self._upgrade_btn
        state = {"wl": None, "busy": False}

        def _apply(_e=None):
            if state["busy"]:
                return
            try:
                row_w = row.winfo_width()
            except Exception:
                return
            # 预留宽度按按钮的**意图显隐** ``_visible`` 判断（由 set_visible 维护）——
            # winfo_ismapped 在布局时序里会瞬时返回 0，不可靠（首版踩过）。
            try:
                reserve = (btn._fit_width() + 12) if getattr(btn, "_visible", True) else 0
            except Exception:
                reserve = 0
            w = row_w - reserve
            # 只挡构建期 1px 占位（w 会变负）；不放 winfo_viewable 早退——label 的
            # <Configure> 常常只在搭建期触发一次且那时尚未映射，早退会让我们永远接不到
            # 真实行宽（按钮一直被挤）。拿到真实行宽就设，after(30) 再校正 deiconify 后
            # 的最终宽度，后续 resize / 按钮显隐都由 label <Configure> 接着收口。
            if w <= 20 or w == state["wl"]:
                return
            state["wl"] = w
            state["busy"] = True
            try:
                lbl.configure(wraplength=max(40, w))
            finally:
                state["busy"] = False

        lbl.bind("<Configure>", _apply, add="+")
        aid = lbl.after(30, _apply)

        def _cancel(_e=None):
            try:
                lbl.after_cancel(aid)
            except Exception:
                pass
        lbl.bind("<Destroy>", _cancel, add="+")

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

        ``_relayout_all()`` 不能省 —— 这里会随状态显隐结果明细、展开/收起 Advanced、
        切换结果明细，都是"内容尺寸变了但不触发 <Configure>"的操作（见该方法 docstring）。

        v2.3.62：整段（按钮/进度条/结果明细的 pack-forget + 引导条换文案 + 重新测量）
        关进**同一道**重绘闸门，出闸门只画一次最终画面。旧版没关闸，而中间调用的
        ``_set_status`` / ``_set_report`` 各自又会 ``_relayout_all()`` 一次（每次
        自带 thaw→RedrawWindow），于是一次状态切换要连画两三帧**互不相同的中间态**
        —— 进度条压在引导条上、结果明细先出现再被挤走，就是老板报的"花屏/跳动"。
        闸门可嵌套（深度计数），外面若已经关了闸，这里只加深一层、不提前解冻。
        """
        with self._paused_paint():
            self._sync_ui_state()
            self._update_guidance()
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

        主/次按钮已移到右卡操作区底部（Check 在上、Fix 在 results 态显现）；
        右卡空状态与进度条互斥显示，结果与保存报告同条件显隐。
        """
        tr = self.tr
        paper = bool(self.docx_path.get()) and os.path.isfile(self.docx_path.get())
        # fixing 也算"已有结果"：修正过程中不该把引导卡又弹回来（会闪一下）
        results = self._phase in ("results", "fixed", "fixing")

        # 右卡：1-2-3 引导**常驻**（用户要求"内容不折叠、右卡不稳定"），进度条按状态显隐。
        # 1-2-3 永远插在「状态行」之前、标题/引导条之后，顺序钉死，导入/删除论文时
        # 顶部引导永不跑位（根治 v2.3.23 之前"1-2-3 掉到右下角"问题）；进度条仅在已选
        # 论文后显示在 1-2-3 之下（空态隐藏，避免右栏空白时还挂着 Prepare→Check→Fix，
        # 与 test_ui_redesign 约定一致）。
        try:
            # v2.3.29：只在 manager 状态真变化时才 pack/forget —— 否则每次 _sync_ui 都
            # 重新 pack 一遍 _progress，触发它的 <Configure>→draw() 重画（delete all 再
            # 重绘），一次检查里被调多次就肉眼可见地闪（用户报"花屏"）。
            if self._empty_state.winfo_manager() != "pack":
                self._empty_state.pack(fill="x", pady=(theme.SECTION_GAP, 0),
                                       before=self._status_row)
            if paper:
                if self._progress.winfo_manager() != "pack":
                    self._progress.pack(fill="x", pady=(theme.SECTION_GAP, 0),
                                       before=self._status_row)
            else:
                if self._progress.winfo_manager() == "pack":
                    self._progress.pack_forget()
        except Exception:
            pass

        # 结果明细与引导卡相反：有报告才出现（见 _build_right 的说明）。
        # v2.3.19：保存报告按钮与结果同显隐（同属「操作产出」，放在右卡）。
        # ⚠️ 同上的 packing 顺序坑：结果明细也要 ``before=self._ops_group`` 钉回
        # 状态行之下、操作按钮组之上的原位，否则重新显示时会掉到右卡最底部。
        try:
            # v2.3.29：同上，只在状态真变化时才 pack/forget，避免冗余 <Configure> 重画。
            if results:
                if self.summary_lbl.winfo_manager() != "pack":
                    self.summary_lbl.pack(fill="x", pady=(3, 0),
                                          before=self._ops_group)
                if self.summary_hint.winfo_manager() != "pack":
                    self.summary_hint.pack(fill="x", before=self._ops_group)
                if self._save_btn.winfo_manager() != "pack":
                    self._save_btn.pack(fill="x", pady=(8, 0))
            else:
                if self.summary_lbl.winfo_manager() == "pack":
                    self.summary_lbl.pack_forget()
                if self.summary_hint.winfo_manager() == "pack":
                    self.summary_hint.pack_forget()
                if self._save_btn.winfo_manager() == "pack":
                    self._save_btn.pack_forget()
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
            return

        if not paper:
            self._set_step(0)
            self._check_btn.set_text(tr("btn_check"))
            # 空态也保持主按钮可用：点一下给出友好行内提示（见 _run 的缺稿分支），
            # 比一上来就置灰更有「引导下一步」的明确感（v2.3.4）。
            self._check_btn.set_enabled(True)
            self._check_btn.set_action(self._run)
            if self._phase in ("idle", "error"):
                self._set_status(tr("state_missing_paper"), "idle",
                                 spec=("state_missing_paper", ()))
            return

        if self._phase == "fixed":
            self._set_step(3)
            self._check_btn.set_text(tr("btn_review"))
            self._check_btn.set_enabled(True)
            self._check_btn.set_action(self._review_changes)
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
            return

        # ready（含出错后可重试）
        self._set_step(1)
        self._check_btn.set_text(tr("btn_check"))
        self._check_btn.set_enabled(True)
        self._check_btn.set_action(self._run)

    # ------------------------------------------------------------ 引导 / 授权映射
    def _update_guidance(self):
        """右卡常驻「下一步」引导条：随状态机更新（v2.3.16，治"向导性不够好"）。"""
        tr = self.tr
        try:
            if self._phase == "fixed":
                text = tr("guide_fixed")
            elif self._phase == "results":
                issues = self._last_issues or 0
                text = tr("guide_ok") if not issues else tr("guide_fix", issues)
            elif self._busy:
                text = tr("guide_checking")
            elif self._phase == "ready" or (self.docx_path.get()
                                            and os.path.isfile(self.docx_path.get())):
                text = tr("guide_check")
            else:
                text = tr("guide_pick")
            self._guide_lbl.config(text=text)
        except Exception:
            pass

    def _gate_message(self, gate):
        """把授权层返回的 ``gate`` 按 ``reason`` 映射成本地图文案。

        v2.3.16 修复：旧代码直接 ``gate.get("message")``，而 license 层所有 message 都是
        中文硬编码，英文版下就显示中文（老板实测"试用提示词出现中文"）。这里只按 reason
        取 i18n，绝不直接显示后端中文串；reason 未知时退回到中性标题文案。
        """
        tr = self.tr
        reason = (gate or {}).get("reason")
        if reason == "trial_exhausted":
            return tr("gate_trial_exhausted")
        if reason == "needs_reactivation":
            return tr("gate_needs_reactivation")
        if reason == "revoked":
            return tr("gate_revoked")
        if reason == "state_error":
            return tr("gate_state_error")
        if reason == "activated":
            return tr("gate_activated")
        if reason == "trial":
            return tr("gate_trial", lic.TRIAL_FIX_LIMIT)
        return tr("msg_need_activation_title")

    # ------------------------------------------------------------ 状态刷新
    def _refresh_rows(self):
        """刷新各输入源的状态显示（已选 / 未选）+ 触发界面状态机。"""
        tr = self.tr

        def _sel(lbl, value, none_key):
            if value:
                name = os.path.basename(value)
                full = "●   " + name
                lbl._full_text = full
                lbl.config(fg=theme.PRIMARY)
                # v2.3.33：按实际像素宽动态截断，避免长文件名盖住右侧 ✕/⌄
                self._fit_field_label(lbl)
            else:
                lbl._full_text = ""
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
        # 问卷 / AI 两行只显示「填没填」（v2.3.64：不再罗列会检查哪些格式项，
        # 明细留给复核弹窗）。问卷给出**已填项数**：只写 "Filled in" 客户不知道
        # 是填了两项还是全填了；只列格式项又列不全还挤掉按钮。条数两头都顾上。
        _q_data = self.questionnaire_data or {}
        _n_q = len(_q_data)
        _hf_levels = _q_data.get("heading_formats") or {}
        if _hf_levels:
            # 每级标题明细按「一项一级」计数：整块只算一项会让客户以为漏填了
            _n_q += len(_hf_levels) - 1
        if _n_q:
            self._q_name.config(text=tr("st_q_set_n", _n_q), fg=theme.SUCCESS)
        else:
            self._q_name.config(text=tr("st_q_none"), fg=theme.TEXT_3)
        # AI 配置：已导入时给「已导入 · 文件名」——只说"已导入"看不出是哪个文件，
        # 只显示文件名又看不出状态。文件名按字符截断：超长名会把 c2 的请求宽度顶大，
        # 反过来挤扁右侧按钮（v2.3.64 修的就是这个连锁反应）。
        _ai_file = self.ai_json.get()
        if _ai_file:
            _nm = os.path.basename(_ai_file)
            if len(_nm) > 32:
                _nm = _nm[:14] + " … " + _nm[-14:]
            self._ai_name.config(text=tr("st_ai_set", _nm), fg=theme.SUCCESS)
        else:
            self._ai_name.config(text=tr("st_ai_none"), fg=theme.TEXT_3)

        # 选了真实存在的论文 → 从 idle 进入 ready。
        # results / fixed 由动作本身决定，这里绝不能覆盖（否则一刷新就把结果态抹掉）。
        cur = self.docx_path.get()
        # 兜底：换了文档就作废旧结论（换文件走 _on_input_changed，这里防其它入口漏掉）
        if self._checked_path and cur != self._checked_path:
            self._invalidate_results()
        chosen = bool(cur) and os.path.isfile(cur)
        if chosen and self._phase in ("idle", "error"):
            self._phase = "ready"
        # v2.3.19：各输入行的「✕」清除按钮——选中才显示，未选隐藏
        self._set_clear_visible(self._paper_clear, bool(self.docx_path.get()))
        self._set_clear_visible(self._tpl_clear, bool(self.school_template.get()))
        self._set_clear_visible(self._ai_clear, bool(self.ai_json.get()))
        # 高级区可用态（选内置规范时置灰）随来源刷新
        self._paint_advanced_state()
        self._sync_ui()

    # -------------------------------------------------- 最近使用（v2.3.78）
    def _restore_last_session(self):
        """启动自动恢复上一次：把上次的手填问卷模板 / AI 导入文件读回界面。

        只恢复 manual / ai 两种来源（其余来源无本地记忆数据，沿用默认）；
        AI 文件路径若已不存在则跳过恢复（不恢复一个打不开的链接）。
        任何异常都静默忽略 —— 恢复失败绝不阻塞启动。
        """
        try:
            rec = recents.load_recents()
            last = rec.get("last_source")
            if last == "src_ai":
                p = recents.first_existing_ai(rec)
                if p:
                    self.ai_json.set(p)
                    self.source_mode.set("src_ai")
            elif last == "src_manual":
                data = recents.first_manual(rec)
                if data is not None:
                    self.questionnaire_data = dict(data)
                    self.source_mode.set("src_manual")
        except Exception:
            pass

    def _remember_manual(self):
        """把当前手填问卷写入「最近模板」并落盘（不阻塞、不抛）。"""
        try:
            rec = recents.load_recents()
            rec = recents.add_manual(rec, self.questionnaire_data)
            rec = recents.set_last_source(rec, "src_manual")
            recents.save_recents(rec)
            self._refresh_recent_combos()
        except Exception:
            pass

    def _remember_ai(self, path: str):
        """把刚导入的 AI 配置文件路径写入「最近导入」并落盘。"""
        try:
            rec = recents.load_recents()
            rec = recents.add_ai(rec, path)
            rec = recents.set_last_source(rec, "src_ai")
            recents.save_recents(rec)
            self._refresh_recent_combos()
        except Exception:
            pass

    def _refresh_recent_combos(self):
        """用磁盘上的 recents 重填两个「最近使用」下拉（语言切换/记忆变更后都走这）。

        下拉项是显示文案，真正的 dict / 路径存在 ``_manual_recent_items`` /
        ``_ai_recent_items`` 里，按选中下标回查，避免把摘要反解析回数据。
        控件尚未建好（极少路径）时静默跳过，不抛。
        """
        cb_m = getattr(self, "_manual_recent", None)
        cb_a = getattr(self, "_ai_recent", None)
        if cb_m is None and cb_a is None:
            return
        try:
            rec = recents.load_recents()
        except Exception:
            rec = {"manual": [], "ai": []}
        if cb_m is not None:
            items = rec.get("manual") or []
            self._manual_recent_items = [it.get("data") for it in items]
            if items:
                vals = [recents.manual_summary(it.get("data"), self.lang) for it in items]
            else:
                vals = [self.tr("recents_none")]
            try:
                cb_m.config(values=vals)
                cb_m.set("")
            except Exception:
                pass
        if cb_a is not None:
            ai_items = [it for it in (rec.get("ai") or [])
                        if os.path.isfile(it.get("path", ""))]
            self._ai_recent_items = [it.get("path") for it in ai_items]
            if ai_items:
                vals = [os.path.basename(it.get("path")) for it in ai_items]
            else:
                vals = [self.tr("recents_none")]
            try:
                cb_a.config(values=vals)
                cb_a.set("")
            except Exception:
                pass

    def _on_manual_recent_pick(self, _e=None):
        """点「最近模板」下拉：把选中的问卷 dict 回填到 questionnaire_data 并切到手填来源。"""
        cb = getattr(self, "_manual_recent", None)
        if cb is None:
            return
        idx = cb.current()
        if idx is None or idx < 0 or idx >= len(self._manual_recent_items):
            return
        data = self._manual_recent_items[idx]
        if data is None:
            return
        self.questionnaire_data = dict(data)
        if self.source_mode.get() != "src_manual":
            self._apply_source_mode("src_manual")
        else:
            self._on_input_changed()
        try:
            cb.set("")
        except Exception:
            pass

    def _on_ai_recent_pick(self, _e=None):
        """点「最近导入」下拉：把选中的配置文件重新导入（路径失效则移出历史）。"""
        cb = getattr(self, "_ai_recent", None)
        if cb is None:
            return
        idx = cb.current()
        if idx is None or idx < 0 or idx >= len(self._ai_recent_items):
            return
        path = self._ai_recent_items[idx]
        if not path or not os.path.isfile(path):
            # 文件已消失：从历史里清掉，刷新下拉
            try:
                rec = recents.load_recents()
                rec = recents.remove_ai(rec, path)
                recents.save_recents(rec)
            except Exception:
                pass
            self._refresh_recent_combos()
            return
        self.ai_json.set(path)
        if self.source_mode.get() != "src_ai":
            self._apply_source_mode("src_ai")
        else:
            self._on_input_changed()
        try:
            cb.set("")
        except Exception:
            pass

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

    # ------------------------------------------------------------ 重绘闸门
    # v2.3.32 "花屏"根治（2026-10-01 真机像素取证）：
    #
    # 一次布局变动要分好几步才走完：pack/unpack → 量卡片内容 → 改画布请求尺寸 →
    # 重排卡片内部控件。`update_idletasks()` **会**把当时的画面推上屏幕，但它推的是
    # 一个**还没走完**的中间状态，而且推得不完整（旧像素残留在被让出来的区域里）——
    # 于是屏幕上同时出现"新旧两套内容"：标题被截断、进度条压在步骤行上、同一行字
    # 出现两遍。这就是老板报的"花屏"（截图实证：点 Check 后右卡糊成一团）。
    #
    # 更糟的是 `run_check()` / `fix_docx()` 会**阻塞主线程好几秒**，事件循环停摆，
    # 坏帧就一直挂在屏幕上"花"着不消失，直到跑完才自己恢复。
    #
    # 解法：给重绘装一道**闸门**。布局期间用 Win32 的 WM_SETREDRAW 把顶层窗的重绘
    # 关掉（Tk 的控件全部画在顶层窗里，关一个就够），布局走完再打开并强制整窗重绘。
    # 用户看到的只有"变动前"和"变动后"两个正常画面，中间态一帧都不会上屏。
    # 非 Windows 平台直接 no-op，退回原有行为（X11/macOS 的 Tk 本来就没这毛病）。
    _WM_SETREDRAW = 0x000B
    _RDW_FLAGS = 0x0001 | 0x0004 | 0x0080 | 0x0100   # INVALIDATE|ERASE|ALLCHILDREN|UPDATENOW

    def _paint_hwnds(self, win=None):
        """窗口句柄（Tk 在 Windows 上会给顶层窗套一层 wrapper，两个都发一遍最稳）。

        ``win`` 传了就算那个窗口的（弹窗是独立 Toplevel，句柄与主窗无关）；不传默认主窗。
        """
        out = []
        try:
            h = (win or self.root).winfo_id()
            if h:
                out.append(h)
            p = ctypes.windll.user32.GetParent(h)
            if p:
                out.append(p)
        except Exception:
            pass
        return out

    def _freeze_win_paint(self, win, freeze: bool, redraw: bool = True):
        """冻结 / 解冻**指定窗口**的重绘，用完立刻解冻（不进深度计数）。

        与 ``_freeze_paint`` 的分工：那个是主窗的**可嵌套**闸门（深度计数在
        ``_paint_depth``）；弹窗（Toplevel）自带独立 HWND，主窗的闸门管不到它，
        所以弹窗内的一次性显隐（如折叠面板展开）要单独冻一次。非 Windows 平台 no-op。

        ``redraw=False``（v2.3.89）：解冻但**先别画**。给「解冻后还要改窗口尺寸」的
        场景用 —— 冻结期间 ``geometry()`` 会被整条丢掉（实测：``WM_SETREDRAW(0)``
        期间请求 600x560，``wm geometry`` 仍报 600x400，解冻后也不补上），
        所以尺寸变更必须排到解冻之后；先解冻不重绘、改完尺寸再统一画，才不会出现
        「旧尺寸画一帧 → 再改尺寸重画」的抖动。
        """
        if not sys.platform.startswith("win"):
            return
        try:
            user32 = ctypes.windll.user32
        except Exception:
            return
        for h in self._paint_hwnds(win):
            try:
                if freeze:
                    user32.SendMessageW(h, self._WM_SETREDRAW, 0, 0)
                else:
                    user32.SendMessageW(h, self._WM_SETREDRAW, 1, 0)
                    if redraw:
                        user32.RedrawWindow(h, None, None, self._RDW_FLAGS)
            except Exception:
                pass

    def _freeze_paint(self, freeze: bool):
        """关 / 开整窗重绘，**可嵌套**（闸门会一层套一层，见 `_toggle_advanced`）。

        用深度计数：只有最外层关闸时才真的发 WM_SETREDRAW(0)，只有回到最外层时才
        真的打开并重绘。否则内层一退出就提前解冻，外层剩下的动作又会花屏。
        """
        depth = getattr(self, "_paint_depth", 0)
        if freeze:
            # 深度计数在**所有平台**都维护：非 Windows 虽然没得可冻结，但"闸门有没有
            # 逐层关好、有没有漏开"这个不变量必须一致，测试才钉得住。
            self._paint_depth = depth + 1
            if depth > 0:
                return                      # 已在闸门内，什么都不用做
            self._paint_frozen = True
            if not sys.platform.startswith("win"):
                return
            for h in self._paint_hwnds():
                try:
                    ctypes.windll.user32.SendMessageW(h, self._WM_SETREDRAW, 0, 0)
                except Exception:
                    pass
            return
        # 解冻
        self._paint_depth = max(0, depth - 1)
        if self._paint_depth > 0:
            return
        self._thaw_paint()

    def _thaw_paint(self):
        """强制开闸 + 整窗重绘。是"保险丝"：异常路径也要保证窗口能重新画出来。"""
        self._paint_depth = 0
        self._paint_frozen = False
        if not sys.platform.startswith("win"):
            return
        user32 = ctypes.windll.user32
        for h in self._paint_hwnds():
            try:
                user32.SendMessageW(h, self._WM_SETREDRAW, 1, 0)
                user32.RedrawWindow(h, None, None, self._RDW_FLAGS)
            except Exception:
                pass

    def _paused_paint(self):
        """上下文管理器：块内一切布局变动都不上屏，出块时一次性画成最终画面。"""
        return _PausedPaint(self)

    def _flush_frame(self):
        """把当前界面**完整画一遍**再返回（几何 + 重绘都处理）。

        ``update_idletasks()`` 只跑几何计算、且画得**不完整**（见上面重绘闸门的说明）；
        ``update()`` 才会把重绘一并做完。紧接着要进"阻塞好几秒"的引擎调用时必须用它，
        否则屏幕上会一直挂着坏帧（实测：冻结帧 vs 正确帧差异 4.64% → 0.00%）。

        安全说明：``update()`` 会派发待处理的输入事件，但调用点都在按钮回调里，
        且 ``_busy`` 会把 ``_run`` / ``_run_fix`` 的重入挡在门口。

        ⚠️ **禁止在 ``with self._paused_paint():`` 块内调用本方法** —— 它会无条件
        ``_thaw_paint()``（把深度计数清零并强制开闸重绘），等于把外层闸门撕开、
        把中间态画上屏。这是**故意的**：它是异常路径的保险丝，万一某处抛异常把闸门
        留在关闭态（窗口永远空白，比花屏严重得多），只有它能无条件救回来。
        闸门内若需要推一帧，用 ``update_idletasks()``（只算几何、不上屏）。
        """
        self._thaw_paint()          # 万一是异常路径逃出来的，先解冻
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        if getattr(self, "_closing", False):
            return
        try:
            self.root.update()
        except Exception:
            pass

    def _set_busy(self, busy: bool):
        """忙态统一交给状态机（RoundButton 自绘，没有 ttk 的 state 选项）。"""
        self._busy = bool(busy)
        with self._paused_paint():
            self._sync_ui()
        # 必须用 _flush_frame（而非 update_idletasks）：见 v2.3.32 花屏说明。
        self._flush_frame()

    # ------------------------------------------------------------ 语言切换
    def _set_lang(self, lang: str):
        """切换语言：重建整套界面（低频操作，重建比逐条改文案可靠得多）。"""
        if lang == self.lang or lang not in i18n.LANGS:
            return
        i18n.save_lang(lang)
        self.lang = lang
        self.F.rebuild(lang)
        self._close_popups()
        # 冻结窗口几何：切语言前先抓真实屏幕坐标与尺寸，重建后钉死还原。
        # 坐标约定（与 _fit_modal 一致）：用 winfo_x/winfo_y（窗口外框屏幕坐标），
        # 不用 winfo_rootx/rooty —— 后者含标题栏/边框偏移（dy = rooty - y），喂回
        # geometry 会越跑越偏。主窗是 Tk 根窗两者本应相等，这里统一走可靠约定。
        try:
            _gx, _gy = self.root.winfo_x(), self.root.winfo_y()
            _gw, _gh = self.root.winfo_width(), self.root.winfo_height()
            _geo = "%dx%d+%d+%d" % (_gw, _gh, _gx, _gy)
        except Exception:
            _geo = None
        for w in self.root.winfo_children():
            w.destroy()
        self._last_scale = None
        self._build()
        self._on_resize()
        # 文案长度随语言变化会改变卡片/滚动区内容高度，必须主动重新测量，否则新文案会
        # 被旧的圆角矩形裁掉或顶不到位（界面"飘"/错位）。这一步也会触发一批 <Configure>。
        self._relayout_all()
        # v2.3.16：重建后重新驱动状态机，确保「下一步」引导条与按钮态跟随新语言/当前阶段
        # （否则切语言后引导条会空着，直到下一次交互才更新）。
        try:
            self._sync_ui()
        except Exception:
            pass
        # 先把重建带来的所有待处理布局 flush 掉，再把几何钉死成"最后一句"：
        # 重建触发的 WM 重排可能在上面的同步还原之后才被处理，把窗口又挪了位。
        # update_idletasks 让内容尺寸落定 → geometry(_geo) 钉死 → after_idle 再补钉一刀，
        # 覆盖任何真正滞后的 WM 重排，确保切语言后窗口纹丝不动。
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        if _geo:
            try:
                self.root.geometry(_geo)
                self._on_resize()
            except Exception:
                pass

            def _restate():
                try:
                    self.root.geometry(_geo)
                    self._on_resize()
                except Exception:
                    pass
            try:
                self.root.after_idle(_restate)
            except Exception:
                pass
        else:
            try:
                self._on_resize()
            except Exception:
                pass

    # ------------------------------------------------------------ 选择文件
    def _on_input_changed(self):
        """任何会影响「体检结论」的输入变了 → 作废旧结论 + 刷新界面。

        统一入口，避免逐个选择器漏掉（Codex 2026-09-12 P1-2）。
        """
        if self._checked_path:
            self._invalidate_results()
        self._refresh_rows()

    # ------------------------------------------------------------ ✕ 清除按钮（v2.3.19）
    def _clear_input(self, which):
        """论文 / 模板 / AI 配置行的小「✕」清除——导入错了点叉号即清除、可重新导入。

        · 只清掉被点的那一项，保留其它输入；
        · 但已做的检查/修正结论随对应源消失而失效，故一并作废、回到 ready/初始态
          （避免「文件已删、结果还显示」的脏状态）；
        · 全量重来仍走论文空态引导（选论文即回 ready）。
        """
        if getattr(self, "_busy", False):
            return
        if which == "paper":
            self.docx_path.set("")
        elif which == "template":
            self.school_template.set("")
        elif which == "ai":
            self.ai_json.set("")
        else:
            return
        if self._phase in ("results", "fixed"):
            self._invalidate_results()
        self._refresh_rows()

    # ------------------------------------------------------------ 文件名动态截断（v2.3.33）
    def _fit_field_label(self, lbl):
        """按 label 实际像素宽把完整文件名截断成「…」，保证不盖住右侧「✕ / ⌄」。

        取代旧版固定 40 字符截断：窄面板（Advanced 里的输入行）或更长文件名仍会
        像素溢出、盖住清除按钮。此处量真实宽度、二分截到能装下的最长前缀 + 省略号；
        label 未布局好（宽 ≤ 1）时延到 idle 再量，避免量到 0 误判。
        """
        full = getattr(lbl, "_full_text", None)
        if not full:
            return
        # 量「左框」的宽（label 被框住，文字超出即裁切、绝不溢出到右侧 ✕/⌄ 框）。
        holder = lbl.master
        total = holder.winfo_width()
        # 宽度还没布局好（比如 Advanced 收起、内部行不可见）就直接返回，不要
        # after_idle 无限重试——折叠行永远拿不到正宽度，重试会死循环。
        # 等展开 / 缩放时，上面的 <Configure> 绑定会自动再算。
        if total <= 1:
            return
        avail = total - 8  # 左右内边距留白
        if avail <= 0:
            return
        f = lbl["font"]
        # lbl["font"] 拿到的是 Tk 字体名字符串（如 "font30"），用 lbl.tk.call 量像素宽。
        # 注意 App 没有 .tk 属性，必须用控件自己的 lbl.tk，否则会抛 AttributeError
        # 被 except 吞掉、误把完整名设回去 → 截断失效、长名盖住 ✕。
        _measure = lambda t: lbl.tk.call("font", "measure", f, t)
        try:
            full_w = _measure(full)
        except Exception:
            full_w = None
        if full_w is None or full_w <= avail:
            if lbl["text"] != full:
                lbl["text"] = full
            return
        ell = "…"
        try:
            ell_w = _measure(ell)
        except Exception:
            ell_w = 0
        max_body = avail - ell_w
        if max_body <= 0:
            lbl["text"] = ell
            return
        lo, hi = 0, len(full)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            try:
                w = _measure(full[:mid])
            except Exception:
                w = mid * 8
            if w <= max_body:
                lo = mid
            else:
                hi = mid - 1
        new = full[:lo] + ell
        if lbl["text"] != new:
            lbl["text"] = new

    def _set_clear_visible(self, btn, show):
        """显隐某个「✕」清除按钮（论文用 place 固定右上角，其余用 pack 行尾）。"""
        if btn is None:
            return
        if btn is self._paper_clear:
            if show:
                btn.place(relx=1.0, rely=0.0, x=-10, y=8, anchor="ne")
            else:
                btn.place_forget()
        else:
            # v2.3.29：把 ✕ 钉在下拉箭头「⌄」左侧（before=chevron），避免挤在箭头右边
            # 被当成装饰、看不见。chevron 在 _select_field 里记到了 clear_btn._chevron。
            before = getattr(btn, "_chevron", None)
            # v2.3.64：AI 行的 ✕ 在「Import」按钮左侧（不是下拉框），间距要留在右边，
            # 否则 ✕ 会贴着 Import 的圆角边。
            padx = (0, 6) if btn is getattr(self, "_ai_clear", None) else (4, 0)
            if show and btn.winfo_manager() != "pack":
                btn.pack(side="right", padx=padx, before=before)
            elif not show and btn.winfo_manager() == "pack":
                btn.pack_forget()

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

    def _pick_ai_config(self):
        p = filedialog.askopenfilename(
            title=self.tr("btn_ai_import"),
            filetypes=[("YAML", "*.yaml *.yml"), ("JSON", "*.json"),
                       ("All files", "*.*")])
        if p:
            self.ai_json.set(p)
            self._remember_ai(p)
            self._on_input_changed()

    # ------------------------------------------------------------ 问卷
    def _open_questionnaire(self):
        """手填问卷弹窗（v2.3.64 整窗按新设计重做）。

        旧版是重设计时**唯一漏掉的弹窗**：``_scroll_frame`` + 一列 ttk.Entry +
        16 行平铺 + 满行 ``#888888`` 灰字，跟复核 / 购买 / 反馈三个 ModalShell
        弹窗完全不像同一个产品。这次统一到同一套骨架：

        ① 骨架换 ``ModalShell``（14px 圆角、单一滚动条权威、``refresh_layout`` 收口）；
        ② 16 个字段按 schema 的 ``group`` 分三块（页面排版 / 标题层级 / 参考文献），
           不再一字排开；
        ③ 栏数与会话复核弹窗用同一个词（单栏 / 双栏）下拉，客户不用猜 "1" 是什么；
        ④ 帮助文字降到字段下方的小字，不再占第三列；
        ⑤ 顺带修一个真 bug：旧版每次打开都按 schema **默认值**重置，
           ``self.questionnaire_data`` 里已保存的值**不回显** —— 客户填完再打开
           像是白填了，这也是它看着"根本没改过"的原因之一。现在先取已存值，
           没有才用默认值。
        """
        tr = self.tr
        F = self.F
        # v2.3.67 实测（probe_table_width + probe_dpi）：明细表 11 列自然宽 756
        # （96 DPI 口径）+ 内容左右各 24 padding = 804，780 连滚动条都不够 → 末三列
        # （Num/UPPER/Run-in）被裁。且 point 制字体随系统 DPI 渲染，老板机器 150%
        # 缩放时实际要 1.5 倍宽 —— 定宽必须乘缩放系数，否则高分屏照样裁。
        _Q_WIDTH = min(int(840 * self._ui_scale()),
                       self.root.winfo_screenwidth() - 16)
        win = ModalShell(self.root, tr("q_title"), width=_Q_WIDTH, height=720,
                         sticky_actions=True)
        self._track(win)
        # 居中是按父窗算的，弹窗比父窗宽时会推出屏幕左缘；夹一次保证整窗在屏内
        try:
            win.update_idletasks()
            _sx = max(8, min(win.winfo_x(),
                             self.root.winfo_screenwidth() - _Q_WIDTH - 8))
            if _sx != win.winfo_x():
                win.geometry("+%d+%d" % (_sx, win.winfo_y()))
        except Exception:
            pass
        body = win.content
        saved = dict(self.questionnaire_data or {})

        tk.Label(body, text=tr("q_hint"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_HELP"], anchor="w", justify="left",
                 wraplength=int(620 * self._ui_scale())).pack(
                     fill="x", pady=(0, 4))

        one_txt = tr("review_one_column")
        two_txt = tr("review_two_column")
        # 栏数用「显示名 ↔ 数值」显式映射，不把展示文案反解析成数字。
        # ⚠️ StringVar 一律显式给 master：不给就走 tkinter 的"默认 root"，一旦进程里
        # 先后建过/销毁过别的 Tk，变量会挂到已经死掉的解释器上（表现是下拉显示空白）。
        col_var = tk.StringVar(master=win,
                               value=two_txt if str(saved.get("columns", "")) == "2"
                               else one_txt)

        vars_: dict = {}
        first_wrap = None      # 第一个可折叠区块的外壳，预填提示行要插在它前面
        for gi, (gkey, gtitle) in enumerate((("page", "q_sec_page"),
                                             ("heading", "q_sec_heading"),
                                             ("reference", "q_sec_refs"))):
            fields = [f for f in QUESTIONNAIRE_SCHEMA if f.get("group") == gkey]
            if not fields:
                continue
            # v2.3.67：三大区块可折叠（老板：要能折叠，填完后可以折叠）。
            # 收起只是 pack_forget，变量与控件都活着——收起后再选参考文献规范，
            # 每级明细的预填照样生效，重新展开就能看到。
            sec, wrap = self._collapsible_section(
                body, tr(gtitle), opened=True,
                pad=0 if gi == 0 else theme.SECTION_GAP,
                on_toggle=lambda: self._fit_modal(win, _Q_WIDTH))
            if first_wrap is None:
                first_wrap = wrap
            grid = tk.Frame(sec, bg=theme.SURFACE)
            grid.pack(fill="x", pady=(0, 4))
            grid.columnconfigure(0, weight=1, uniform="qcol")
            grid.columnconfigure(1, weight=1, uniform="qcol")
            for i, field in enumerate(fields):
                r, c = divmod(i, 2)
                cell = tk.Frame(grid, bg=theme.SURFACE)
                cell.grid(row=r, column=c, sticky="new",
                          padx=(0, 14) if c == 0 else (0, 0), pady=(0, 10))
                self._q_field(cell, field, vars_, saved, col_var, one_txt, two_txt)
            if gkey == "heading":
                # 层级数这个数字只是"有几级"，真正干活的参数在下面这张明细表里
                lvl_var = vars_.get("heading_levels", (None, None))[1]
                self._q_heading_table(sec, saved, vars_, lvl_var)

        # —— 参考文献规范 → 默认值联动（老板 2026-10-19：客户选中哪个规范，
        #    默认值要调整为哪个规范的）——
        # 规则：只预填客户「没填过」的字段；客户手填过/上次保存过的值一律不动。
        # 预填本身也不算"客户填过"——切到另一个规范时，预填值跟着新规范走。
        state = {"prefilling": False}
        touched: set = set()

        prefill_lbl = tk.Label(body, text="", bg=theme.SURFACE, fg=theme.PRIMARY,
                               font=F["F_FOOT"], anchor="w", justify="left",
                               wraplength=int(620 * self._ui_scale()))

        def _show_prefill_hint(style_key):
            prefill_lbl.configure(text=tr("q_prefilled", style_key))
            try:
                prefill_lbl.pack(fill="x", pady=(4, 0), before=first_wrap)
            except Exception:
                prefill_lbl.pack(fill="x", pady=(4, 0))

        def _note_touch(key):
            def _on(*_a):
                if not state["prefilling"]:
                    touched.add(key)
            return _on

        for _key, _entry in vars_.items():
            if _key in ("_hf_opts", "heading_formats"):
                continue
            try:
                _entry[1].trace_add("write", _note_touch(_key))
            except Exception:
                pass

        _HF_ROW_KEYS = ("font", "size", "bold", "italic", "page_break_before",
                        "align", "indent", "num", "upper", "runin")
        for _row in (vars_.get("heading_formats") or []):
            for _vk in _HF_ROW_KEYS:
                _row[_vk].trace_add(
                    "write", _note_touch(("hf", _row["lvl"], _vk)))

        # 上次保存过的值 = 客户此前的明确选择，规范联动绝不覆盖
        touched.update(k for k in saved.keys() if k != "heading_formats")
        _HF_FIELD_OF = {"font": "font_family", "size": "size_pt",
                        "align": "alignment", "indent": "indent_in",
                        "num": "number_prefix", "upper": "uppercase",
                        "runin": "run_in", "bold": "bold", "italic": "italic",
                        "page_break_before": "page_break_before"}
        for _k, _v in (saved.get("heading_formats") or {}).items():
            try:
                _lvl = int(_k)
            except (TypeError, ValueError):
                continue
            for _vk in _HF_ROW_KEYS:
                if _HF_FIELD_OF[_vk] in (_v or {}):
                    touched.add(("hf", _lvl, _vk))

        def _apply_spec_defaults(style_key):
            """把**参考文献相关**的空白字段按所选引用规范的典型值预填。

            🔴 v2.3.73（老板 2026-10-20 定的口径）：这个下拉叫「参考文献格式」，
            客户在这里表达的只是「我的参考文献按 APA 走」，**不是**「全篇按 APA 走」。
            所以：
              · 只预填 group == "reference" 的字段；
              · 页面 / 标题明细 / 标题页**一律不预填** —— 客户走手动输入这条路，
                正是因为有学校自己的要求或自己的个性化值；把 APA 的五级标题格式自动
                灌进去等于替他拍板，还会让检查/修正按 APA 改他的手排标题。
              · 想全篇按规范走的客户，直接在顶上的「规范」下拉里选 APA/IEEE 就行。

            值的单一真相源仍是引擎的 ``target_from_spec``——UI 不抄一份规范默认值。
            """
            try:
                t = target_from_spec(style_key)
            except Exception:
                return
            state["prefilling"] = True
            try:
                for field in QUESTIONNAIRE_SCHEMA:
                    key = field["key"]
                    if field.get("group") != "reference" or key == "reference_style":
                        continue
                    entry = vars_.get(key)
                    if not entry:
                        continue
                    val = getattr(t, key, None)
                    kind = entry[0]
                    if kind == "tri":
                        if key in touched or val is None \
                                or entry[1].get() != tr("q_none"):
                            continue
                        entry[1].set(tr("bool_yes") if val else tr("bool_no"))
                    elif kind == "bool":
                        if key in touched or not isinstance(val, bool):
                            continue
                        entry[1].set(tr("bool_yes") if val else tr("bool_no"))
                    elif kind == "cols":
                        if key in touched or val not in (1, 2):
                            continue
                        entry[1].set(two_txt if val == 2 else one_txt)
                    else:
                        if key in touched or val is None:
                            continue
                        sval = str(val)
                        opts = field.get("options") or []
                        # choice 字段即便显示着 schema 默认值，只要客户没手改
                        # （touched 未记录）就跟着所选的引用规范走。
                        # 客户手改过/保存过的由 touched 挡住。
                        if opts and sval not in opts:
                            continue
                        entry[1].set(sval)
            finally:
                state["prefilling"] = False
            _show_prefill_hint(style_key)

        _rs = vars_.get("reference_style")

        def _ref_style_value():
            """下拉显示名 → 引擎值（"custom"/"APA"/"MLA"/"IEEE"；未匹配 None）。

            选项展示名可能是中文（「按我校要求（自定义）」），绝不能把显示文字
            当引擎值用 —— 统一在这里反查。
            """
            pairs = [(o, i18n.questionnaire_option(self.lang, "reference_style", o))
                     for o in REFERENCE_STYLE_CHOICES]
            return _opt_val(pairs, _rs[1].get().strip())

        def _on_style_change(*_a):
            # 只有真的选了某个内置引用规范才预填参考文献细则；
            # 「按我校要求（自定义）」= 什么都不套，客户填什么就是什么（v2.3.73）。
            if not state["prefilling"]:
                style = _ref_style_value() if _rs else None
                if style in REFERENCE_STYLE_KEYS:
                    _apply_spec_defaults(style)

        if _rs:
            _rs[1].trace_add("write", _on_style_change)

        def _save():
            data: dict = {}
            warnings: list = []
            for field in QUESTIONNAIRE_SCHEMA:
                key = field["key"]
                if key == "columns":
                    # 下拉的显示名 → 数值字符串（_coerce 会转成 int）
                    data[key] = "2" if col_var.get() == two_txt else "1"
                    continue
                if field.get("type") == "tri":
                    _txt = vars_[key][1].get()
                    if _txt == tr("bool_yes"):
                        data[key] = True
                    elif _txt == tr("bool_no"):
                        data[key] = False
                    # 「—」= 不规定：不写进数据，引擎那边就是 None
                    continue
                if field.get("type") == "bool":
                    _kind, cb, _v = vars_[key]
                    data[key] = (cb.get() == tr("bool_yes"))
                    continue
                val = vars_[key][1].get().strip()
                if not val or val == tr("q_none"):
                    continue
                if field.get("type") == "choice":
                    # 显示名 → 引擎值（「按我校要求（自定义）」这类双语展示名
                    # 绝不能直接当成 reference_style 的值写进画像）。v2.3.73
                    pairs = [(o, i18n.questionnaire_option(self.lang, key, o))
                             for o in (field.get("options") or [])]
                    hit = _opt_val(pairs, val)
                    if hit is None:
                        continue        # 未登记的显示名 → 视为「不规定」，不臆造值
                    data[key] = hit
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

            # —— 每级标题明细（与 AI 导出模板的 heading_formats 同一套字段）——
            hf_opts = vars_.get("_hf_opts") or {}
            hf_out: dict = {}
            for row in (vars_.get("heading_formats") or []):
                if not row.get("state", {}).get("show"):
                    continue      # 超出「层级数」的行不参与，避免把没打算填的算进来
                item: dict = {}
                font = row["font"].get().strip()
                if font:
                    item["font_family"] = font
                for key2, varkey, label in (("size_pt", "size", tr("q_h_size")),
                                            ("indent_in", "indent", tr("q_h_indent")),
                                            # v2.3.75：段前 / 段后（磅）与行距倍数
                                            ("space_before_pt", "space_before", tr("q_h_before")),
                                            ("space_after_pt", "space_after", tr("q_h_after")),
                                            ("line_spacing", "line_spacing", tr("q_h_lsp"))):
                    txt = row[varkey].get().strip()
                    if not txt:
                        continue
                    try:
                        item[key2] = float(txt)
                    except ValueError:
                        messagebox.showwarning(
                            tr("q_bad_number_title"),
                            tr("q_bad_number",
                               "%s · %s" % (tr("q_h_lvl", row["lvl"]), label), txt),
                            parent=win)
                        return
                    if key2 == "size_pt" and not (6.0 <= item[key2] <= 72.0):
                        warnings.append("%s · %s = %s  (6–72)" % (
                            tr("q_h_lvl", row["lvl"]), label, txt))
                    if key2 == "indent_in" and not (0.0 <= item[key2] <= 3.0):
                        warnings.append("%s · %s = %s  (0–3)" % (
                            tr("q_h_lvl", row["lvl"]), label, txt))
                    if key2 in ("space_before_pt", "space_after_pt") \
                            and not (0.0 <= item[key2] <= 120.0):
                        warnings.append("%s · %s = %s  (0–120)" % (
                            tr("q_h_lvl", row["lvl"]), label, txt))
                    if key2 == "line_spacing" and not (0.5 <= item[key2] <= 5.0):
                        warnings.append("%s · %s = %s  (0.5–5)" % (
                            tr("q_h_lvl", row["lvl"]), label, txt))
                al = _opt_val(hf_opts.get("aligns"), row["align"].get())
                if al:
                    item["alignment"] = al
                npre = _opt_val(hf_opts.get("prefixes"), row["num"].get())
                if npre:
                    item["number_prefix"] = npre
                for key2, varkey in (("bold", "bold"), ("italic", "italic"),
                                     ("uppercase", "upper"), ("run_in", "runin"),
                                     ("page_break_before", "page_break_before")):
                    tv = _opt_val(hf_opts.get("tris"), row[varkey].get())
                    if tv is not None:
                        item[key2] = tv
                if item:
                    hf_out[str(row["lvl"])] = item
            if hf_out:
                data["heading_formats"] = hf_out

            if warnings and not messagebox.askyesno(
                    tr("q_range_title"), tr("q_range_body", "\n".join(warnings)),
                    parent=win):
                return
            self.questionnaire_data = data
            self._remember_manual()
            messagebox.showinfo(tr("q_saved_title"), tr("q_saved"), parent=win)
            self._close_one(win)
            self._on_input_changed()

        win.add_primary_action(tr("q_save"), command=_save)
        win.add_cancel_action(tr("q_cancel"))
        self._fit_modal(win, _Q_WIDTH)

    def _q_field(self, cell, field, store, saved, col_var, one_txt, two_txt):
        """问卷弹窗里的单个字段：标签 / 输入控件 / 下方小字帮助。

        控件按 schema 的 ``type`` 选：bool → 是/否下拉，choice → 只读下拉，
        ``columns`` → 单栏/双栏下拉（与复核弹窗同一个词），其余 → 文本框。
        """
        tr = self.tr
        F = self.F
        key = field["key"]
        tk.Label(cell, text=i18n.questionnaire_label(self.lang, field),
                 bg=theme.SURFACE, fg=theme.TEXT_2, font=F["F_FOOT"],
                 anchor="w").pack(fill="x")
        cur = saved.get(key, field.get("default", ""))
        if key == "columns":
            ttk.Combobox(cell, textvariable=col_var, state="readonly", width=16,
                         font=F["F_SMALL"], values=[one_txt, two_txt]).pack(fill="x")
            store[key] = ("cols", col_var)
        elif field.get("type") == "tri":
            # 三态开关（v2.3.65）：「—」= 不规定，是/否 = 明确要求。
            # 参考文献的「另起一页 / 标题加粗 / 标题居中」必须能表达"不规定"，
            # 否则客户没意见的项会被当成"明确不要"，反倒去改文档。
            cur_b = cur if isinstance(cur, bool) else None
            v = tk.StringVar(
                master=cell,
                value=(tr("q_none") if cur_b is None
                       else (tr("bool_yes") if cur_b else tr("bool_no"))))
            ttk.Combobox(cell, textvariable=v, state="readonly", width=16,
                         font=F["F_SMALL"],
                         values=[tr("q_none"), tr("bool_yes"), tr("bool_no")]).pack(fill="x")
            store[key] = ("tri", v)
        elif field.get("type") == "bool":
            v = tk.BooleanVar(master=cell, value=bool(cur))
            cb = ttk.Combobox(cell, state="readonly", width=16,
                              font=F["F_SMALL"],
                              values=[tr("bool_yes"), tr("bool_no")])
            cb.set(tr("bool_yes") if v.get() else tr("bool_no"))
            cb.pack(fill="x")
            store[key] = ("bool", cb, v)
        elif field.get("type") == "choice":
            # 「可留空」的 choice（schema default 为空，如纸张大小）：补一个「—」，
            # 否则客户一旦选了 A4 就再也回不到「不规定」，只能去改文件。
            opts = list(field.get("options", []))
            allow_none = field.get("default", None) in ("", None)
            # 选项展示名统一过 i18n.questionnaire_option（值→显示名，可双语），
            # 保存时用 _opt_val 反查回原值：绝不让中文/英文显示文字变成引擎值。
            pairs = [(o, i18n.questionnaire_option(self.lang, key, o)) for o in opts]
            shown = [d for _v, d in pairs]
            if allow_none:
                shown = [tr("q_none")] + shown
            init = next((d for v, d in pairs if str(cur) == v),
                        tr("q_none") if allow_none else "")
            v = tk.StringVar(master=cell, value=init)
            ttk.Combobox(cell, textvariable=v, values=shown,
                         state="readonly", width=16,
                         font=F["F_SMALL"]).pack(fill="x")
            store[key] = ("str", v)
        else:
            v = tk.StringVar(master=cell, value=("" if cur is None else str(cur)))
            tk.Entry(cell, textvariable=v, font=F["F_BODY"], width=18,
                     bg=theme.SURFACE, fg=theme.TEXT,
                     highlightbackground=theme.BORDER,
                     highlightthickness=1).pack(fill="x")
            store[key] = ("str", v)
        tk.Label(cell, text=i18n.questionnaire_help(self.lang, field),
                 bg=theme.SURFACE, fg=theme.TEXT_3, font=F["F_FOOT"],
                 anchor="w", justify="left", wraplength=262).pack(
            fill="x", pady=(2, 0))

    # ------------------------------------------------------------ 每级标题明细
    def _q_heading_table(self, parent, saved, store, lvl_var):
        """问卷弹窗里的「每级标题明细」表（v2.3.65）。

        为什么要有这张表：老板的原话是「各层级的字体字号缩进斜体等等参数你做了吗？
        AI 导入与手动输入一样，只是让 AI 帮客户写」。此前的问卷只问一句「标题层级数」，
        学校要的「一级黑体三号居中、二级楷体四号左对齐、三级首行缩进 0.5 英寸」全无处
        可填；更糟的是读进来的学校模板参数修正器也不用（被写死的通用规则覆盖）。

        这张表与 AI 模板里的 ``heading_formats`` 是**同一套字段**（见
        ai_questionnaire_template.yaml），两个入口最终都落进 TargetProfile.heading_formats，
        检查与修正走同一条链路 —— 手填能填的，AI 也能填；AI 能填的，手填也能填。

        取值一律用「显示名 ↔ 引擎值」显式映射，绝不把界面文字反解析成值（改文案就不会坏）。
        """
        tr = self.tr
        F = self.F
        aligns = [(None, tr("q_none")), ("left", tr("q_align_left")),
                  ("center", tr("q_align_center")), ("right", tr("q_align_right")),
                  ("justify", tr("q_align_justify"))]
        prefixes = [(None, tr("q_none")), ("number", tr("q_prefix_number")),
                    ("roman", tr("q_prefix_roman")),
                    ("upper_letter", tr("q_prefix_upper_letter")),
                    ("lower_letter", tr("q_prefix_lower_letter")),
                    ("lower_roman", tr("q_prefix_lower_roman")),
                    ("number_paren", tr("q_prefix_number_paren"))]
        # 加粗 / 斜体 / 全大写 / 接排是**三态**：—=不规定，是/否=明确要求（或明确不要）
        tris = [(None, tr("q_none")), (True, tr("bool_yes")), (False, tr("bool_no"))]
        store["_hf_opts"] = {"aligns": aligns, "prefixes": prefixes, "tris": tris}

        tk.Label(parent, text=tr("q_h_table_hint"), bg=theme.SURFACE, fg=theme.TEXT_3,
                 font=F["F_FOOT"], anchor="w", justify="left",
                 wraplength=int(620 * self._ui_scale())).pack(fill="x", pady=(2, 6))

        table = tk.Frame(parent, bg=theme.SURFACE)
        table.pack(fill="x", pady=(0, 6))

        # v2.3.67 重排（老板实测截图「非常乱」）：旧版每级占两行、两组表头夹在
        # 控件行中间，列位对不上。现在**每级一行**、表头只有一行，11 列与控件
        # 严格同列对齐；弹窗同步加宽到 700 放得下。
        hdr = {1: tr("q_h_font"), 2: tr("q_h_size"), 3: tr("q_h_bold"),
               4: tr("q_h_italic"), 5: tr("q_h_pagebreak"),
               6: tr("q_h_align"), 7: tr("q_h_indent"), 8: tr("q_h_num"),
               9: tr("q_h_upper"), 10: tr("q_h_runin"),
               # v2.3.75：段前 / 段后 / 行距（老板 2026-10-20：段前段后是调整格式的核心）
               11: tr("q_h_before"), 12: tr("q_h_after"), 13: tr("q_h_lsp")}
        for col, text in hdr.items():
            tk.Label(table, text=text, bg=theme.SURFACE, fg=theme.TEXT_3,
                     font=F["F_FOOT"], anchor="w").grid(
                row=0, column=col, sticky="w", padx=(0, 8), pady=(0, 2))

        saved_hf: dict = {}
        for k, v in (saved.get("heading_formats") or {}).items():
            try:
                saved_hf[int(k)] = dict(v or {})
            except (TypeError, ValueError):
                continue

        def _disp(opts, value) -> str:
            for v, d in opts:
                if v == value:
                    return d
            return opts[0][1]

        rows: list = []
        for lvl in range(1, 10):          # schema 里 heading_levels 上限是 9
            cur = saved_hf.get(lvl, {})
            ws: list = []
            r = lvl                       # 每级一行（表头占第 0 行）

            def _lab(text, col, row):
                w = tk.Label(table, text=text, bg=theme.SURFACE, fg=theme.TEXT_2,
                             font=F["F_SMALL"], anchor="w")
                w.grid(row=row, column=col, sticky="w", padx=(0, 8), pady=(0, 3))
                ws.append(w)
                return w

            def _combo(col, row, opts, value, width):
                var = tk.StringVar(master=table, value=_disp(opts, value))
                w = ttk.Combobox(table, textvariable=var, state="readonly", width=width,
                                 font=F["F_SMALL"], values=[d for _v, d in opts])
                w.grid(row=row, column=col, sticky="w", padx=(0, 8), pady=(0, 3))
                ws.append(w)
                return var

            def _entry(col, row, value, width):
                var = tk.StringVar(master=table,
                                   value="" if value is None else str(value))
                w = tk.Entry(table, textvariable=var, font=F["F_SMALL"], width=width,
                             bg=theme.SURFACE, fg=theme.TEXT,
                             highlightbackground=theme.BORDER, highlightthickness=1)
                w.grid(row=row, column=col, sticky="w", padx=(0, 8), pady=(0, 3))
                ws.append(w)
                return var

            _lab(tr("q_h_lvl", lvl), 0, r)
            font_var = tk.StringVar(master=table, value=str(cur.get("font_family") or ""))
            cb_font = ttk.Combobox(table, textvariable=font_var, width=10,
                                   font=F["F_SMALL"],
                                   values=["Times New Roman", "Arial", "Calibri",
                                           "Cambria", "Georgia", "SimSun", "SimHei",
                                           "KaiTi", "FangSong"])
            cb_font.grid(row=r, column=1, sticky="w", padx=(0, 8), pady=(0, 3))
            ws.append(cb_font)
            size_var = _entry(2, r, cur.get("size_pt"), 4)
            bold_var = _combo(3, r, tris, cur.get("bold"), 5)
            ital_var = _combo(4, r, tris, cur.get("italic"), 5)
            # 标题分页（page_break_before）三态：—=不规定，是/否=明确要求。
            # 与加粗/斜体同口径走 tris，AI 导入可设、手填也能设，二者一致。
            pbb_var = _combo(5, r, tris, cur.get("page_break_before"), 6)
            al_var = _combo(6, r, aligns, cur.get("alignment"), 7)
            ind_var = _entry(7, r, cur.get("indent_in"), 4)
            num_var = _combo(8, r, prefixes, cur.get("number_prefix"), 7)
            up_var = _combo(9, r, tris, cur.get("uppercase"), 5)
            runin_var = _combo(10, r, tris, cur.get("run_in"), 5)
            # v2.3.75：段前 / 段后（磅）与行距倍数。留空 = 不规定（一个字节都不写）。
            before_var = _entry(11, r, cur.get("space_before_pt"), 4)
            after_var = _entry(12, r, cur.get("space_after_pt"), 4)
            lsp_var = _entry(13, r, cur.get("line_spacing"), 4)

            state = {"show": True}
            rows.append({
                "lvl": lvl, "state": state, "widgets": ws,
                "font": font_var, "size": size_var, "bold": bold_var,
                "italic": ital_var, "page_break_before": pbb_var,
                "align": al_var, "indent": ind_var,
                "num": num_var, "upper": up_var, "runin": runin_var,
                "space_before": before_var, "space_after": after_var,
                "line_spacing": lsp_var,
            })

        def _sync_rows(*_a):
            try:
                want = int(float(lvl_var.get() or 0)) if lvl_var is not None else 0
            except (TypeError, ValueError):
                want = 0
            # 至少露出 2 行：客户还没想好"几级"时，也得看得见这张表存在
            visible = max(want, 2)
            for row in rows:
                row["state"]["show"] = row["lvl"] <= visible
                for w in row["widgets"]:
                    if row["state"]["show"]:
                        w.grid()
                    else:
                        w.grid_remove()

        if lvl_var is not None:
            try:
                lvl_var.trace_add("write", _sync_rows)
            except Exception:
                pass
        store["heading_formats"] = rows
        _sync_rows()

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

        padx = 44

        # v2.3.80：标题**钉在顶部**，不随正文滚动。
        # 老板实测：一滚滑轮，上面的标题就跟着走、还露出一段空白 —— 因为标题
        # / 版本原先和正文一起塞在滚动区（_scroll_frame）里。头部独立成 frame
        # 先 pack(side="top")，正文才进滚动区，标题就永远在最上面。
        head = tk.Frame(win, bg=PAPER)
        head.pack(side="top", fill="x")
        tk.Label(head, text=tr("app_title"), bg=PAPER, fg=INK,
                 font=self.F["F_TITLE"]).pack(padx=padx, pady=(20, 2))
        tk.Label(head, text=tr("about_version", self.version), bg=PAPER, fg=MUTED,
                 font=self.F["F_SUBTITLE"]).pack(padx=padx, pady=(0, 8))
        tk.Frame(head, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 10))

        inner = _scroll_frame(win, bg=PAPER)

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

        padx = 40

        # v2.3.80：同 About —— 标题钉在顶部，不随正文滚动（原来标题/导语在
        # 滚动区里，一滚就跟着走）。
        hdr = tk.Frame(win, bg=PAPER)
        hdr.pack(side="top", fill="x")
        tk.Label(hdr, text=tr("help_title"), bg=PAPER, fg=INK,
                 font=self.F["F_TITLE"]).pack(padx=padx, pady=(18, 4))
        _intro = tk.Label(hdr, text=tr("help_intro"), bg=PAPER, fg=BODY,
                          font=self.F["F_BODY"], justify="center", wraplength=660)
        _auto_wrap(_intro, hdr, padx)
        _intro.pack(padx=padx, pady=(0, 10))
        tk.Frame(hdr, bg=LINE, height=1).pack(fill="x", padx=padx, pady=(0, 8))

        inner = _scroll_frame(win, bg=PAPER)

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

        section(tr("help_sec1"))
        for title, body in tr("help_steps"):
            item(title, body)

        section(tr("help_sec2"))
        for title, body in tr("help_faqs"):
            item(title, body)

        tk.Label(inner, text=tr("about_footer"), bg=PAPER, fg=MUTED,
                 font=self.F["F_FOOT"]).pack(padx=padx, pady=(16, 18))

    # ------------------------------------------------------------ 反馈窗口
    # 提交到 Worker /api/feedback → 邮件；断网 / 失败则回落 mailto: 开邮件客户端。
    FEEDBACK_ENDPOINT = "https://api.reedskill.com/api/feedback"

    @staticmethod
    def _is_valid_email(email: str) -> bool:
        """宽松够用的邮箱格式校验（仅反馈收集用，不做严格 RFC 校验）。"""
        if not email or "@" not in email:
            return False
        return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email) is not None

    @staticmethod
    def _feedback_post(email: str, message: str, version: str) -> bool:
        """POST 到 Worker /api/feedback；任何异常 / 非 2xx / 服务端 ok=false
        一律返回 False（fail-soft，交给 _feedback_deliver 走 mailto 兜底）。"""
        payload = {"email": email or "", "message": message,
                   "app": "paperformatpro", "version": version or ""}
        try:
            data = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(
                App.FEEDBACK_ENDPOINT, data=data, method="POST",
                headers={"Content-Type": "application/json",
                         "User-Agent": "PaperFormatPro/%s" % (version or "")})
            with urllib.request.urlopen(req, timeout=8) as resp:
                if not (200 <= getattr(resp, "status", 0) < 300):
                    return False
                try:
                    body = json.loads(resp.read().decode("utf-8"))
                except Exception:
                    return True  # 2xx 但无法解析：保守按已送达处理
                return bool(body.get("ok", True))
        except Exception:
            return False

    @staticmethod
    def _feedback_mailto(email: str, message: str) -> None:
        """兜底：开邮件客户端，内容预填好（from 行 + 正文）。"""
        addr = i18n.BRAND_EMAIL
        body = message if not email else "From: %s\n\n%s" % (email, message)
        try:
            webbrowser.open("mailto:%s?subject=%s&body=%s" % (
                addr,
                urllib.parse.quote("PaperFormat Pro feedback"),
                urllib.parse.quote(body)))
        except Exception:
            pass

    @classmethod
    def _feedback_deliver(cls, email: str, message: str, version: str) -> tuple:
        """纯函数：裁决反馈提交结果，返回 ``(status, detail)``。

        status ∈ {"invalid_email", "empty", "sent", "mailto"}

        - 邮箱非必填；但若填了且格式不对 → ``invalid_email``
        - 反馈内容必填，空 → ``empty``
        - 网络可达且 2xx → ``sent``
        - 否则回落 ``mailto`` 兜底（开邮件客户端）→ ``mailto``
        """
        email = (email or "").strip()
        message = (message or "").strip()
        if email and not cls._is_valid_email(email):
            return ("invalid_email", "")
        if not message:
            return ("empty", "")
        if cls._feedback_post(email, message, version):
            return ("sent", "")
        cls._feedback_mailto(email, message)
        return ("mailto", "")

    def _feedback_dialog(self):
        tr = self.tr
        F = self.F
        win = ModalShell(self.root, tr("fb_title"), width=560, height=520)
        self._track(win)
        body = win.content

        tk.Label(body, text=tr("fb_intro"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_SMALL"], anchor="w", justify="left",
                 wraplength=480).pack(fill="x")

        # —— 邮箱（选填）——
        tk.Label(body, text=tr("fb_email"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_LABEL"], anchor="w").pack(fill="x", pady=(14, 4))
        email_var = tk.StringVar(master=win)
        ttk.Entry(body, textvariable=email_var, width=48,
                  font=F["F_BODY"]).pack(fill="x")
        tk.Label(body, text=tr("fb_email_hint"), bg=theme.SURFACE, fg=theme.TEXT_3,
                 font=F["F_FOOT"], anchor="w").pack(fill="x", pady=(3, 0))

        # —— 反馈内容（必填）——
        tk.Label(body, text=tr("fb_message"), bg=theme.SURFACE, fg=theme.TEXT,
                 font=F["F_LABEL"], anchor="w").pack(fill="x", pady=(14, 4))
        msg = tk.Text(body, width=48, height=6, font=F["F_BODY"],
                      relief="solid", bd=1, highlightthickness=0,
                      wrap="word", bg=theme.SURFACE, fg=theme.TEXT)
        msg.pack(fill="x")

        # —— 隐私说明（一行，GDPR 友好的弱提示）——
        tk.Label(body, text=tr("fb_privacy"), bg=theme.SURFACE, fg=theme.TEXT_3,
                 font=F["F_FOOT"], anchor="w", justify="left",
                 wraplength=480).pack(fill="x", pady=(10, 0))

        # —— 行内状态（成功绿 / 错误红，不弹窗打断）——
        status_var = tk.StringVar(master=win, value="")
        status_lbl = tk.Label(body, textvariable=status_var, bg=theme.SURFACE,
                              fg=theme.ERROR, font=F["F_FOOT"], anchor="w",
                              wraplength=480)
        status_lbl.pack(fill="x", pady=(6, 0))

        def _submit():
            status_var.set("")
            email = email_var.get().strip()
            message = msg.get("1.0", "end").strip()
            result, _ = App._feedback_deliver(email, message, self.version)
            if result == "invalid_email":
                status_var.set(tr("fb_invalid_email"))
                status_lbl.config(fg=theme.ERROR)
                return
            if result == "empty":
                status_var.set(tr("fb_empty"))
                status_lbl.config(fg=theme.ERROR)
                return
            if result == "sent":
                status_var.set(tr("fb_sent"))
                status_lbl.config(fg=theme.SUCCESS)
                win.after(1000, lambda: self._close_one(win))
                return
            # mailto 兜底
            status_var.set(tr("fb_mailto_fallback"))
            status_lbl.config(fg=theme.TEXT_2)

        win.add_primary_action(tr("fb_submit"), command=_submit)
        win.add_cancel_action(tr("fb_cancel"))
        self._fit_modal(win, 560)

    # ------------------------------------------------------------ 升级 / 激活（合并弹窗）
    def _upgrade_dialog(self):
        """升级 / 购买 / 激活 合并弹窗（规格 COMMERCIAL UI）。

        设计（老板 2026-10-01 定）：
        - 顶部一行 SKU 信息（次卡/季卡/年卡/终身 四档现价，与官网 pricing.html 对齐），让客户买前就看清价格；
        - 一个「购买」按钮直接开官网 pricing 页（价格单一来源在 store.js，app 不写死）；
        - 在线激活码输入框**直接放在弹窗主区域**（默认可见），客户买前就知道码填哪儿、
          买后粘贴即激活，无需再开第二个弹窗；
        - 离线激活（机器码 + 离线码）默认折叠，断网灾备用。

        授权逻辑仍走 ``lic.activate`` / ``lic.activate_offline``，本弹窗只换外壳。
        """
        tr = self.tr
        F = self.F
        # v2.3.89：吸底动作栏（老板 2026-10-22「升级/激活那个弹窗这个也做了吧」）。
        # 内容高度随「离线激活」折叠区展开而变：收起时约 640px 勉强装得下，
        # 一展开就冲到 1100px 以上 —— Close 按钮排在 _inner 末尾，展开后
        # 客户得滚到最底才能关（右上角 ✕ 虽在，但不该逼客户去找）。
        # 置 sticky 后 Close 固定在窗口右下角，折叠展开都纹丝不动。
        win = ModalShell(self.root, tr("up_title"), width=600, height=720,
                         sticky_actions=True)
        self._track(win)
        body = win.content

        tk.Label(body, text=tr("up_subtitle"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_SMALL"], anchor="w", justify="left",
                 wraplength=520).pack(fill="x")

        # SKU 一行（四档现价；须与官网 pricing.html / store.js 单一来源保持一致，
        # 官网改价时这里同步改，否则客户在 app 里看到的就是旧价）。
        # 特惠价用鲜艳珊瑚红高亮、原价划线，突出「现在买最划算」。
        sku = RoundCard(body, radius=9, fill=theme.SURFACE_SOFT,
                        border=theme.INFO_BORDER, shadow=False, padx=14, pady=8)
        sku.pack(fill="x", pady=(12, 0))
        sku_row = tk.Frame(sku.inner, bg=theme.SURFACE_SOFT)
        sku_row.pack(fill="x", pady=2)

        def _seg(text, fg, overstrike=False, padx=0):
            fnt = F["F_LABEL"]
            if overstrike:
                fnt = fnt.copy()
                fnt.config(overstrike=1)
            tk.Label(sku_row, text=text, bg=theme.SURFACE_SOFT, fg=fg,
                     font=fnt, anchor="w").pack(side="left", padx=padx)

        SKU_HOT = "#E8453C"  # 特惠价高亮（珊瑚红，与青绿主色形成对比）
        _seg(tr("up_tier_two"), theme.PRIMARY_HOVER)
        _seg(" ($5.99)", theme.TEXT_2)
        _seg("  ·  ", theme.TEXT_3)
        _seg(tr("up_tier_quarter"), theme.PRIMARY_HOVER)
        _seg(" ($39)", theme.TEXT_2)
        _seg("  ·  ", theme.TEXT_3)
        _seg(tr("up_tier_annual"), theme.PRIMARY_HOVER)
        _seg(" ($79)", theme.TEXT_2)
        _seg("  ·  ", theme.TEXT_3)
        _seg(tr("up_tier_lifetime") + " (", theme.PRIMARY_HOVER)
        _seg("$149", SKU_HOT)
        _seg(" $199", theme.TEXT_3, overstrike=True)
        _seg(")", theme.PRIMARY_HOVER)

        # 购买按钮（开官网 pricing 页）
        RoundButton(body, text=tr("up_buy"), command=self._open_buy,
                    style="primary", font=F["F_BTN"], height=38,
                    padx=18).pack(fill="x", pady=(12, 0))

        tk.Label(body, text=tr("up_note"), bg=theme.SURFACE, fg=theme.TEXT_3,
                 font=F["F_FOOT"], anchor="w", justify="left",
                 wraplength=520).pack(fill="x", pady=(8, 0))

        # 在线激活（主区域，默认可见）—— 买前就知道码填哪儿
        act = RoundCard(body, radius=12, fill=theme.SURFACE_SOFT,
                        border=theme.INFO_BORDER, shadow=False, padx=14, pady=10)
        act.pack(fill="x", pady=(12, 0))
        tk.Label(act.inner, text=tr("up_act_label"), bg=theme.SURFACE_SOFT,
                 fg=theme.PRIMARY_HOVER, font=F["F_SMALL_B"], anchor="w").pack(fill="x")
        code_var = tk.StringVar(master=win)
        ttk.Entry(act.inner, textvariable=code_var, width=40,
                  font=F["F_BODY"]).pack(fill="x", pady=(8, 0))
        tk.Label(act.inner, text=tr("act_online_hint"), bg=theme.SURFACE_SOFT,
                 fg=theme.TEXT_3, font=F["F_FOOT"], anchor="w",
                 justify="left").pack(fill="x", pady=(3, 6))

        def _submit():
            code = code_var.get().strip()
            if not code:
                messagebox.showwarning(tr("act_title"), tr("act_need_code"), parent=win)
                return
            r = lic.activate(code)
            messagebox.showinfo(tr("act_result_title"), r["message"], parent=win)
            if r["ok"]:
                self._refresh_commercial()
                self._close_one(win)

        RoundButton(act.inner, text=tr("act_online_btn"), command=_submit,
                    style="secondary", font=F["F_BTN_S"], height=34,
                    padx=14).pack(anchor="w")

        # 离线激活（默认折叠）
        self._build_offline_fold(win, body)

        win.add_cancel_action(tr("up_close"))
        self._fit_modal(win, 600)

    def _open_buy(self):
        """「购买」按钮：开官网 pricing 页（价格单一来源在 store.js，app 不写死）。"""
        try:
            webbrowser.open(i18n.BRAND_BUY_URL)
        except Exception:
            pass

    def _build_offline_fold(self, win, parent):
        """离线激活折叠区：机器码（复制给卖家）+ 离线激活码（粘贴即用）。默认收起。"""
        tr = self.tr
        F = self.F

        head = tk.Frame(parent, bg=theme.SURFACE, cursor="hand2")
        head.pack(fill="x", pady=(12, 0))
        marker = tk.StringVar(master=win, value="[+]")
        mark_lbl = tk.Label(head, textvariable=marker, bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_LABEL"], width=3, anchor="w")
        mark_lbl.pack(side="left")
        title_lbl = tk.Label(head, text=tr("up_offline_expand"), bg=theme.SURFACE,
                 fg=theme.TEXT_2, font=F["F_LABEL"], anchor="w")
        title_lbl.pack(side="left", padx=(2, 0))

        pane = tk.Frame(parent, bg=theme.SURFACE)

        # 机器码
        mc_var = tk.StringVar(master=win, value=lic._machine_fingerprint())
        mc_card = RoundCard(pane, radius=10, fill=theme.SURFACE_SOFT,
                            border=theme.INFO_BORDER, shadow=False, padx=14, pady=8)
        mc_card.pack(fill="x", pady=(8, 0))
        tk.Label(mc_card.inner, text=tr("act_machine_label"), bg=theme.SURFACE_SOFT,
                 fg=theme.PRIMARY_HOVER, font=F["F_SMALL_B"], anchor="w").pack(fill="x")
        ttk.Entry(mc_card.inner, textvariable=mc_var, width=40, state="readonly",
                  font=F["F_MONO"]).pack(fill="x", pady=(6, 0))
        tk.Label(mc_card.inner, text=tr("act_machine_hint"), bg=theme.SURFACE_SOFT,
                 fg=theme.TEXT_3, font=F["F_FOOT"], anchor="w",
                 justify="left").pack(fill="x", pady=(3, 6))

        def _copy_mc():
            try:
                win.clipboard_clear()
                win.clipboard_append(mc_var.get())
                messagebox.showinfo(tr("act_copy_ok_title"), tr("act_copy_ok"), parent=win)
            except Exception:
                pass

        RoundButton(mc_card.inner, text=tr("act_copy_btn"), command=_copy_mc,
                    style="secondary", font=F["F_BTN_S"], height=32, padx=12).pack(anchor="w")

        # 离线激活码
        off_card = RoundCard(pane, radius=10, fill=theme.SURFACE_SOFT,
                             border=theme.INFO_BORDER, shadow=False, padx=14, pady=8)
        off_card.pack(fill="x", pady=(8, 0))
        tk.Label(off_card.inner, text=tr("act_offline_label"), bg=theme.SURFACE_SOFT,
                 fg=theme.PRIMARY_HOVER, font=F["F_SMALL_B"], anchor="w").pack(fill="x")
        off_var = tk.StringVar(master=win)
        ttk.Entry(off_card.inner, textvariable=off_var, width=40,
                  font=F["F_MONO"]).pack(fill="x", pady=(6, 0))
        tk.Label(off_card.inner, text=tr("act_offline_hint"), bg=theme.SURFACE_SOFT,
                 fg=theme.TEXT_3, font=F["F_FOOT"], anchor="w",
                 justify="left").pack(fill="x", pady=(3, 6))

        def _submit_offline():
            code = off_var.get().strip()
            if not code:
                messagebox.showwarning(tr("act_title"), tr("act_need_offline"), parent=win)
                return
            r = lic.activate_offline(code)
            messagebox.showinfo(tr("act_result_title"), r["message"], parent=win)
            if r["ok"]:
                self._refresh_commercial()
                self._close_one(win)

        RoundButton(off_card.inner, text=tr("act_offline_btn"), command=_submit_offline,
                    style="secondary", font=F["F_BTN_S"], height=32, padx=12).pack(anchor="w")

        def _toggle():
            # 折叠/展开会同时改「内容自然高度」和「窗口高度」，中间态一旦上屏就是
            # 「内容先跳到新位置、窗口还是旧高度」的那一帧 —— 肉眼看到的"弹窗抖一下/漂移"。
            # 弹窗是独立 Toplevel，主窗的重绘闸门管不到它，这里单独冻一次，
            # 显隐 + 重新定高 + 重算滚动区全部做完再整体画出来。
            #
            # ⚠️ v2.3.89：定高必须排在**解冻之后**。冻结（WM_SETREDRAW 0）期间调
            # `win.geometry()` 会被整条丢掉（实测：冻结中请求 560，`wm geometry` 仍报
            # 400，解冻后也不补），于是「点开离线激活 → 内容长高了、窗口没长高」，
            # 新展开的机器码/离线码全落在折线以下，客户以为点了没反应。
            # 现在的顺序：冻结 → 显隐 → 解冻但不重绘 → 定高 → 补一次重绘。
            self._freeze_win_paint(win, True)
            try:
                if pane.winfo_ismapped():
                    pane.pack_forget()
                    marker.set("[+]")
                else:
                    pane.pack(fill="x")
                    marker.set("[-]")
                self._freeze_win_paint(win, False, redraw=False)   # 解冻，先别画
                self._fit_modal(win, 600)      # 内含 refresh_layout（重算滚动区）
            finally:
                self._freeze_win_paint(win, False)                 # 补一次重绘（幂等）

        head.bind("<Button-1>", lambda e: _toggle())
        mark_lbl.bind("<Button-1>", lambda e: _toggle())
        title_lbl.bind("<Button-1>", lambda e: _toggle())
        # 默认折叠（pane 尚未 pack）

    def _ui_scale(self) -> float:
        """系统 DPI 缩放系数（96 DPI = 1.0，150% 缩放 = 1.5）。

        本应用整体是 DPI-naive 的：tkinter 的坐标在 DPI-aware 进程里是物理像素，
        但 point 制字体按真实 DPI 渲染 —— 150% 屏上同一排控件实际占 1.5 倍宽。
        老板机器就是 150%（2026-10-19 实测截图：明细表末三列被物理裁掉，正是这个
        原因）。给宽度敏感的弹窗按此系数放大定宽，不动全局。
        """
        try:
            return max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        except Exception:
            return 1.0

    # ------------------------------------------------ 原生对话框的显示自愈（v2.3.83）
    #
    # 老板 2026-10-21 报：保存修正稿时窗口突然变小、字变清楚、一直保持到关软件。
    # 根因是「原生 shell 对话框把进程 DPI 感知提升 → Windows 停止位图拉伸」，
    # 完整取证见 src/ui/dpi.py 顶部。主治在 main() 里锁死感知（pin_process_dpi_awareness），
    # 这里这两个方法只是**保险**：万一某些 Windows 版本仍改了缩放/窗口框，弹完立刻还原。
    # 正常路径下 _heal_display 是空操作（restore_if_drifted 只在实测到变化时才动手）。
    def _snap_display(self):
        """原生对话框弹出**之前**抓一份显示状态；抓不到返回 None（不影响主流程）。"""
        try:
            return _dpi_mod.snapshot(self.root)
        except Exception:
            return None

    def _heal_display(self, snap):
        """原生对话框返回**之后**自愈：窗口被改了尺寸/缩放就恢复原样。"""
        if not snap:
            return
        try:
            _dpi_mod.restore_if_drifted(self.root, snap)
        except Exception:
            pass

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
            # v2.3.88：吸底弹窗的动作栏是**窗口底部**的真控件，不在 _inner 里 ——
            # 窗口高度 = 画布可视高 + 动作栏高。不把栏高算进来，窗口就正好少一条
            # 按钮栏的高度：画布比内容矮一截，明明装得下也永远挂个滚动条，最后一行
            # 还被裁在按钮栏后面。非吸底弹窗栏本来就在 _inner 里（reqheight 已含），
            # 不能重复加。
            if getattr(win, "_sticky", False):
                try:
                    need += int(win._button_bar.winfo_reqheight())
                except Exception:
                    pass
            sh = int(self.root.winfo_screenheight())
            # 桌面保留区（任务栏）：1080p 下约 40-48px，取 56 留余量。
            bottom_limit = sh - 56
            # v2.3.88：``geometry`` 给的是**客户区**高度，而屏幕装得下的是**外框**
            # 高度（客户区 + 标题栏 + 边框）。旧写法按客户区直接顶到 bottom_limit，
            # 外框就超出了工作区 —— Windows 会把窗口自己压小：
            # 实测请求 704 高，系统只给了 656（差 48 = 标题栏），
            # 而 pack 早已按 704 排好版 → 吸底动作栏被挤出窗口下沿，
            # 按钮底边落在窗口外 48px（小屏笔记本上就是「按钮看不全 / 点不到」）。
            # dy = 标题栏 + 边框高度，cap 必须先把它扣掉，再谈客户区能有多高。
            dy = max(0, win.winfo_rooty() - win.winfo_y())
            cap = max(320, bottom_limit - 8 - dy)
            height = min(max(need, 1), cap)
            y = min(win.winfo_y(), bottom_limit - height - dy)
            y = max(8, y)
            win.geometry("%dx%d+%d+%d" % (width, height, win.winfo_x(), y))
            # 尺寸定下来之后**必须**再同步一次滚动区：ModalShell 的滚动条是 pack 进
            # 窗口右侧的，内容比可视区高就得立刻挂出来，否则末段内容被裁且滚不到
            # （小屏幕上的典型"布局漂移"）。放在这里一刀收口，三个弹窗（反馈 / 购买激活 /
            # 复核）以及折叠面板展开全部自动受益，不必在调用点各写一遍。
            win.refresh_layout()
        except Exception:
            pass

    def _activate_dialog(self):
        """兼容 dev 工具（截图 / smoke）旧引用：直接打开合并后的「购买 · 激活」弹窗。"""
        self._upgrade_dialog()

    # ------------------------------------------------------------ 目标画像
    def _build_target(self) -> TargetProfile:
        """按**当前来源**组装目标画像（6 选 1 互斥 → 只取这一条路径的输入）。

        v2.3.65：旧版不管选了什么来源，都把 school_template / ai / questionnaire
        一股脑传给引擎。后果有两个:
          ① 选了 Custom，之前点过的学校模板还在背后参与合并（界面看不见、实际生效）；
          ② 选了内置规范，之前填过的问卷会通过"高优先级叠加"盖掉规范本身。
        来源既然是互斥的，就只交代当前那条：规范 → 只给 spec_key；
        学校模板 → 只给 template；自定义 → 只给问卷 / AI。
        """
        ref = self.school_ref_style.get()
        mode = self.source_mode.get()
        is_school = mode == "src_school"
        is_manual = mode == "src_manual"
        is_ai = mode == "src_ai"
        # 默认「不检查引用」对应的显示串不是合法风格键，build_target 会忽略为 None
        return build_target(
            # spec_key 是 StringVar，无规范时是 ""（不是 None）—— 统一转成 None 再交给
            # 引擎，别让空串被当成一个"空规范"（v2.3.65 引擎侧也做了兜底）
            self.spec_key.get() or None,
            template=(self.school_template.get() or None) if is_school else None,
            ai=(self.ai_json.get() or None) if is_ai else None,
            questionnaire=(self.questionnaire_data or None) if is_manual else None,
            reference_style=ref or None,
        )

    # ------------------------------------------------------------ 体检
    def _run(self):
        """运行格式体检（免费不限次）：先弹「复核弹窗」让客户确认/改错，再跑检查。"""
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
        target = self._build_target()
        self._confirm_dialog(target)

    def _confirm_dialog(self, target):
        """共享复核弹窗：把各来源读到的格式值摊开给客户二次确认 / 改错 / 补漏。

        四个类别（Page Layout / Headings / References / Style Rules）分组展示；
        Page Layout 的标量字段可编辑（改错 / 补漏），其余只读展示。点「Run check」
        才真正跑体检（_execute_check）。学校模板 / 手动问卷 / AI 三来源共用此弹窗。
        """
        tr = self.tr
        F = self.F
        # v2.3.88：吸底动作栏。这个弹窗是四个里内容最高的（4 个分区 + 页面网格 +
        # 栏数/纸型/页码三行 + 标题表 + 风格规则），内容必然超出 720 的可视区；
        # 旧版按钮栏排在 _inner 末尾，客户必须一路滚到底才看得见 Back / Run check
        # （老板实测圈出来的就是这块）。置 sticky_actions 后按钮栏是**窗口底部**
        # 的真控件，完全不参与滚动。
        win = ModalShell(self.root, tr("review_title"), width=680, height=720,
                         sticky_actions=True)
        self._track(win)
        body = win.content

        tk.Label(body, text=tr("review_helper"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_SMALL"], anchor="w", justify="left",
                 wraplength=620).pack(fill="x", pady=(0, 8))

        # ModalShell 自身已内建滚动（_cv Canvas + 自动滚动条），content 直接承接
        # 所有 section 即可，不要再嵌一层 _scroll_frame（否则 canvas-in-canvas 双滚动条）。
        inner = body
        warn_fg = getattr(theme, "GOLD", theme.TEXT_3)

        # —— ① Page Layout（可编辑）——
        self._review_section(inner, tr("review_page"), theme.SECTION_GAP)
        page_vars = {}
        label_en = {
            "margin_top_in": "Top margin (in)", "margin_bottom_in": "Bottom margin (in)",
            "margin_left_in": "Left margin (in)", "margin_right_in": "Right margin (in)",
            "font_family": "Body font", "font_size_pt": "Body size (pt)",
            "line_spacing": "Line spacing (x)", "body_alignment": "Body align",
            "first_line_indent_in": "First-line indent (in)",
            "space_after_pt": "Space after (pt)", "header_distance_in": "Header dist. (in)",
        }
        grid = tk.Frame(inner, bg=theme.SURFACE)
        grid.pack(fill="x", padx=12, pady=(0, 8))
        for i, key in enumerate(TargetProfile.PAGE_KEYS):
            row, col = divmod(i, 3)
            cell = tk.Frame(grid, bg=theme.SURFACE)
            cell.grid(row=row, column=col, sticky="ew", padx=4, pady=3)
            grid.columnconfigure(col, weight=1)
            val = target.page.get(key)
            tk.Label(cell, text=label_en.get(key, key), bg=theme.SURFACE, fg=theme.TEXT_2,
                     font=F["F_FOOT"], anchor="w").pack(fill="x")
            # ⚠️ 必须显式 master：不给就走 tkinter 的"默认 root"。全量测试 / 多次开窗时
            # 进程里先后建过又销毁过别的 Tk，变量会挂到已死掉的解释器上 →
            # 下拉/输入框显示**空白**（手工问卷那边早就踩过并注释过同一个坑）。
            var = tk.StringVar(master=win, value=("" if val is None else str(val)))
            ent = tk.Entry(cell, textvariable=var, font=F["F_BODY"], width=14,
                          bg=theme.SURFACE, fg=theme.TEXT,
                          highlightbackground=theme.BORDER, highlightthickness=1)
            ent.pack(fill="x")
            if val is None:
                # 缺失：标黄提示客户确认 / 补填
                ent.config(fg=warn_fg)
            page_vars[key] = var

        # 栏数（结构性维度，不在 PAGE_KEYS 的标量格里）：单选下拉，默认单栏。
        # 用「显示名 ↔ 数值」显式映射，不把展示文案反解析成数字（i18n 文案改动也不会解析失败）。
        one_txt = tr("review_one_column")
        two_txt = tr("review_two_column")
        _none_txt = tr("q_none")
        # 🔴 v2.3.102（老板 2026-10-24 截图）：columns=None 是「该来源不规定栏数」
        # （如 IEEE-Journal 按裁定不碰版式），绝不能默认显示成 "Single column"：
        # ① 那是在说谎；② 确认时会把 "1" 发给 apply_page_overrides，而它只把
        # **空串**当"未改动" → None 被静默升级成「强制单栏」，双栏期刊论文一检查就报错。
        _cols_disp = _columns_display_value(target.columns, _none_txt, one_txt, two_txt)
        columns_var = tk.StringVar(master=win, value=_cols_disp)
        col_row = tk.Frame(inner, bg=theme.SURFACE)
        col_row.pack(fill="x", padx=12, pady=(4, 8))
        tk.Label(col_row, text=tr("review_columns"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_FOOT"], anchor="w").pack(side="left", padx=(0, 8))
        ttk.Combobox(col_row, textvariable=columns_var, state="readonly", width=18,
                     values=[_none_txt, one_txt, two_txt]).pack(side="left")

        # 纸张大小（通用维度：北美 Letter / 欧洲与中国 A4）：与引用规范无关，单独一行。
        paper_var = tk.StringVar(master=win, value=(target.paper_size or tr("q_none")))
        tk.Label(col_row, text=tr("review_paper_size"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_FOOT"], anchor="w").pack(side="left", padx=(18, 8))
        ttk.Combobox(col_row, textvariable=paper_var, state="readonly", width=14,
                     values=[tr("q_none")] + list(PAPER_SIZES)).pack(side="left")

        # 页码位置（v2.3.66 客户可填）：top-right = **页眉**右上（APA/MLA 的语义），
        # bottom-center = 页脚居中（中文论文常规）。留「—」= 不规定、不检查。
        # 单独一行：与栏数/纸型挤同一行会把三项都挤窄（本轮之前的裁字教训）。
        pn_row = tk.Frame(inner, bg=theme.SURFACE)
        pn_row.pack(fill="x", padx=12, pady=(0, 8))
        pn_var = tk.StringVar(master=win, value=(target.page_number_position or tr("q_none")))
        tk.Label(pn_row, text=tr("review_page_number"), bg=theme.SURFACE, fg=theme.TEXT_2,
                 font=F["F_FOOT"], anchor="w").pack(side="left", padx=(0, 8))
        ttk.Combobox(pn_row, textvariable=pn_var, state="readonly", width=18,
                     values=[tr("q_none")] + list(PAGE_NUMBER_POSITIONS)).pack(side="left")

        # —— ② Headings（只读，每级一行）——
        self._review_section(inner, tr("review_headings"), theme.SECTION_GAP)
        hfmt = target.heading_formats or {}
        if hfmt:
            head = tk.Frame(inner, bg=theme.SURFACE)
            head.pack(fill="x", padx=12, pady=(0, 4))
            cols = (tr("review_lvl"), tr("q_h_font"), tr("q_h_size"), tr("q_h_bold"),
                    tr("q_h_italic"), tr("review_align"), tr("q_h_indent"),
                    tr("q_h_num"), tr("review_upper"))
            for j, col in enumerate(cols):
                tk.Label(head, text=col, bg=theme.SURFACE, fg=theme.TEXT_2,
                         font=F["F_FOOT"], anchor="w").grid(
                    row=0, column=j, sticky="w", padx=(0, 6))
            _none = tr("q_none")

            def _tri(v):
                """三态显示：None=不规定（—），True/False=是/否。"""
                if v is None:
                    return _none
                return tr("review_enabled") if v else tr("review_disabled")

            r = 1
            for lvl in sorted(int(k) for k in hfmt):
                d = hfmt.get(str(lvl), hfmt.get(lvl, {}))
                _ind = d.get("indent_in")
                if _ind is None and d.get("indent"):
                    _ind = 0.5
                vals = [str(lvl),
                        str(d.get("font_family") or _none),
                        _none if d.get("size_pt") is None else str(d.get("size_pt")),
                        _tri(d.get("bold")),
                        _tri(d.get("italic")),
                        str(d.get("alignment") or _none),
                        _none if _ind is None else str(_ind),
                        str(d.get("number_prefix") or _none),
                        _tri(d.get("uppercase"))]
                for j, v in enumerate(vals):
                    tk.Label(head, text=v, bg=theme.SURFACE, fg=theme.TEXT,
                             font=F["F_SMALL"], anchor="w").grid(
                        row=r, column=j, sticky="w", padx=(0, 6))
                r += 1
        else:
            # 🔴 v2.3.94（老板 2026-10-23 截图）：客户明明选了 MLA 9th，
            # 这里只写「Not detected — confirm or fill in」—— 读起来像工具没检测出来、
            # 还让客户去填一项**根本不存在**的东西。MLA 9th 自己就不定义分级编号标题体系
            # （``specs.py`` 里 ``heading_levels=0``，notes 原文 "Numbered headings are not
            # required (use a consistent format of your own)"）。这是「规范不要求」，
            # 不是「没检测到」，两者的处置完全相反：前者无需任何动作，后者要客户补。
            # 只有**客户真在规范下拉里选了内置规范**（spec_is_baseline）才敢这么说 ——
            # 走学校模板 / 手填 / AI 时读不到标题格式是真的缺信息，仍按原提示让客户补。
            _h_sp = None
            try:
                _h_sp = checker.resolve_spec(target)
            except Exception:
                _h_sp = None
            # 🔴 v2.3.103（老板 2026-10-24 追问三条通道）：**先读 target.heading_levels**
            # —— 那才是学校模板 / 手填 / AI 的真实输入。旧写法只读 resolve_spec 的层级数：
            # 手填 / AI 没有 spec_key → resolve_spec 落 "Other" = 0，于是客户明明填了
            # 「3 级」，弹窗照样说 "this source sets no heading levels"（自相矛盾）。
            _h_lv = int(getattr(target, "heading_levels", 0) or 0)
            if _h_lv <= 0:
                _h_lv = int(getattr(_h_sp, "heading_levels", 0) or 0)
            _h_name = (getattr(_h_sp, "name", None) or target.spec_key or "")
            if spec_is_baseline(target) and _h_lv <= 0:
                # 规范本身不要求分级标题（MLA 9th）：无需任何动作
                tk.Label(inner,
                         text=tr("review_headings_not_required").format(name=_h_name),
                         bg=theme.SURFACE, fg=theme.TEXT_2,
                         font=F["F_HELP"], anchor="w", justify="left",
                         wraplength=int(620 * self._ui_scale())).pack(
                    fill="x", padx=12, pady=(0, 6))
            elif _h_lv > 0 and spec_is_baseline(target):
                # 🔴 v2.3.102（老板 2026-10-24 截图）：规范**要求** N 级标题，但**不规定
                # 各级字形**（IEEE-Journal 按裁定「不碰版式」，格式交给目标期刊模板）。
                # 以前这里落到 review_headings_missing → 说成 "this source sets no
                # heading levels"，与事实相反（Journal 是 heading_levels=4），
                # 客户会以为要自己去 Advanced options 补标题层级。
                tk.Label(inner,
                         text=tr("review_headings_levels_only").format(
                             name=_h_name, n=_h_lv),
                         bg=theme.SURFACE, fg=theme.TEXT_2,
                         font=F["F_HELP"], anchor="w", justify="left",
                         wraplength=int(620 * self._ui_scale())).pack(
                    fill="x", padx=12, pady=(0, 6))
            elif _h_lv > 0:
                # 🔴 v2.3.103（老板 2026-10-24 追问三条通道）：手填 / AI / 模板给了
                # **层级数**、但没给**各级明细**（heading_formats 为空）。这同样不是
                # "没设层级" —— 只是各级格式待补，指路 Advanced options。
                tk.Label(inner,
                         text=tr("review_headings_levels_no_formats").format(n=_h_lv),
                         bg=theme.SURFACE, fg=warn_fg,
                         font=F["F_HELP"], anchor="w", justify="left",
                         wraplength=int(620 * self._ui_scale())).pack(
                    fill="x", padx=12, pady=(0, 6))
            else:
                # 这个弹窗里标题表是只读的 → 不能写「confirm or fill in」（它自己做不到），
                # 要说清去哪儿填（Advanced options 里的手动问卷）。
                tk.Label(inner, text=tr("review_headings_missing"), bg=theme.SURFACE,
                         fg=warn_fg, font=F["F_HELP"], anchor="w", justify="left",
                         wraplength=int(620 * self._ui_scale())).pack(
                    fill="x", padx=12, pady=(0, 6))

        # —— ③ References（只读）——
        self._review_section(inner, tr("review_refs"), theme.SECTION_GAP)
        _none = tr("q_none")

        def _lbl(key):
            """问卷字段的本地化标签（与手填弹窗同一套词，单一真相）。"""
            return i18n.questionnaire_label(self.lang, {"key": key, "label": key})

        def _tri_txt(v):
            if v is None:
                return _none
            return tr("review_enabled") if v else tr("review_disabled")

        def _num_txt(v, unit):
            return _none if v is None else ("%g%s" % (float(v), unit))

        # v2.3.65：参考文献条目本身的排版也要摊开（学校模板里常与正文不同口径），
        # 客户在「确认」这一步就该看到「条目五号宋体单倍」这类要求，而不是只有风格名。
        ref_lines = [
            "%s: %s" % (_lbl("reference_style"),
                        target.reference_style or tr("review_missing")),
            "%s: %s" % (_lbl("reference_hanging_indent_in"),
                        _num_txt(target.reference_hanging_indent_in, "″")),
            "%s: %s" % (_lbl("reference_numbered"), _tri_txt(target.reference_numbered)),
            "%s: %s" % (_lbl("reference_font_family"),
                        target.reference_font_family or _none),
            "%s: %s" % (_lbl("reference_size_pt"), _num_txt(target.reference_size_pt, "pt")),
            "%s: %s" % (_lbl("reference_line_spacing"),
                        _num_txt(target.reference_line_spacing, "x")),
            "%s: %s" % (_lbl("reference_new_page"), _tri_txt(target.reference_new_page)),
            "%s: %s" % (_lbl("reference_heading_bold"),
                        _tri_txt(target.reference_heading_bold)),
            "%s: %s" % (_lbl("reference_heading_centered"),
                        _tri_txt(target.reference_heading_centered)),
        ]
        for ln in ref_lines:
            tk.Label(inner, text=ln, bg=theme.SURFACE, fg=theme.TEXT, font=F["F_SMALL"],
                     anchor="w", justify="left").pack(fill="x", padx=12, pady=(0, 2))

        # —— 摘要 / Running head / 图表 caption（v2.3.67 新增维度，只读展示）——
        self._review_section(inner, tr("review_frontmatter"), theme.SECTION_GAP)
        _fm_lines = [
            "%s: %s" % (_lbl("abstract_first_line_indent_in"),
                        _num_txt(target.abstract_first_line_indent_in, "″")),
            "%s: %s" % (_lbl("abstract_new_page"), _tri_txt(target.abstract_new_page)),
            "%s: %s" % (_lbl("running_head_required"), _tri_txt(target.running_head_required)),
            "%s: %s" % (_lbl("page_number_include_author_name"),
                        _tri_txt(target.page_number_include_author_name)),
            "%s: %s" % (_lbl("caption_font_size_pt"),
                        _num_txt(target.caption_font_size_pt, "pt")),
            "%s: %s" % (_lbl("caption_line_spacing"),
                        _num_txt(target.caption_line_spacing, "x")),
            "%s: %s" % (_lbl("caption_position_table"),
                        target.caption_position_table or _none),
            "%s: %s" % (_lbl("caption_position_figure"),
                        target.caption_position_figure or _none),
            "%s: %s" % (_lbl("table_lines"), target.table_lines or _none),
        ]
        for ln in _fm_lines:
            tk.Label(inner, text=ln, bg=theme.SURFACE, fg=theme.TEXT, font=F["F_SMALL"],
                     anchor="w", justify="left").pack(fill="x", padx=12, pady=(0, 2))

        # —— ④ Style Rules（只读，由 resolve_spec 解析）——
        self._review_section(inner, tr("review_style"), theme.SECTION_GAP)
        try:
            spec = checker.resolve_spec(target)
            if spec_is_baseline(target):
                style_lines = [
                    "Running head: %s" % (tr("review_enabled") if spec.running_head else tr("review_disabled")),
                    "Page number: %s" % (spec.page_number_position or "—"),
                    "References on new page: %s" % (tr("review_enabled") if spec.reference_new_page else tr("review_disabled")),
                    "Reference heading bold: %s" % (tr("review_enabled") if spec.reference_heading_bold else tr("review_disabled")),
                    "Font allowlist: %s" % (", ".join(spec.font_allowlist) if spec.font_allowlist else "—"),
                ]
            else:
                # 🔴 v2.3.103（老板 2026-10-24 追问三条通道）：学校模板 / 手填 / AI
                # **没有内置规范基线**。即使手填里选了「引用按 APA」，也只有**参考文献**
                # 归 APA；页面 / 标题层级仍由客户自己的输入决定（checker 用
                # spec_is_baseline 门控）。这里若照旧摊开 APA 全套
                # （Page number: top-right / References on new page: On …），
                # 客户会以为页面也要按 APA 改 —— 与 v2.3.72 那个事故同一性质。
                style_lines = [
                    "Page / heading rules: from your own inputs (no built-in style)",
                    "Reference list style: %s (references only)" % (target.reference_style or "—"),
                ]
        except Exception:
            style_lines = [tr("review_missing")]
        for ln in style_lines:
            tk.Label(inner, text=ln, bg=theme.SURFACE, fg=theme.TEXT, font=F["F_SMALL"],
                     anchor="w", justify="left").pack(fill="x", padx=12, pady=(0, 2))

        def _on_confirm():
            # 仅把用户真正改动过的 Page Layout 字段写回 target（其余保持原 resolution，
            # 避免把规范软参考误升级为硬规则）——逻辑在 engine.apply_page_overrides，headless 可测。
            edits = {key: var.get() for key, var in page_vars.items()}
            # 栏数：显示名 → 数值；「—」= 不规定 → 回**空串**（apply_page_overrides 视为
            # "未改动"、保持 None）。绝不能让"不规定"被静默升级成"强制单栏"（v2.3.102）。
            edits["columns"] = _columns_edit_value(columns_var.get(), _none_txt, two_txt)
            # 纸张大小：「—」= 不规定（apply_page_overrides 会清空并取消检查）
            edits["paper_size"] = paper_var.get()
            # 页码位置：同上，「—」= 不规定；白名单外的值 apply_page_overrides 会丢弃
            edits["page_number_position"] = pn_var.get()
            apply_page_overrides(target, edits)
            self._execute_check(target, win)

        win.add_primary_action(tr("review_confirm"), command=_on_confirm)
        win.add_cancel_action(tr("review_back"))
        # 本弹窗内容最高（4 个分区 + 页面网格 + 栏数下拉 + 标题表 + 风格规则），
        # 小屏幕上会超出可视区；_fit_modal 末尾已统一重算滚动区（内容被裁且滚不到的根治点）。
        self._fit_modal(win, 680)

    def _review_section(self, parent, title, pad):
        """复核弹窗里的小节标题（带上方留白）。"""
        tk.Label(parent, text=title, bg=theme.SURFACE, fg=theme.PRIMARY,
                 font=self.F["F_LABEL"], anchor="w").pack(
            fill="x", padx=12, pady=(pad, 4))

    def _collapsible_section(self, parent, title, opened=True, pad=0,
                             on_toggle=None):
        """可折叠小节（老板 2026-10-19：手填问卷的大区块要能折叠，填完可收起）。

        标题行就是开关：标题在左、chevron 标记在**右**（▾ 展开 / ▸ 收起）——
        右侧折叠箭头符合多数人的视觉习惯。点整行切换。收起只是 ``pack_forget``，
        变量与控件都还活着——收起「标题层级」后选参考文献规范，每级明细的
        规范联动预填照样生效，重新展开就能看到。

        返回 ``(内容容器, 外壳)``；外壳给调用方当 pack 排序锚点用。
        """
        wrap = tk.Frame(parent, bg=theme.SURFACE)
        wrap.pack(fill="x", pady=(pad, 2))
        # 标题行：左标题 + 右 chevron。
        # ⚠️ Tk 的 bindtags 不含中间 Frame：点标题文字/箭头（子 Label）时事件
        # 不会冒泡到 head Frame，只绑 Frame 会导致「点箭头/标题没反应」。
        # 所以标题文字与箭头两个 Label 各自也要绑点击（点 Frame 间隙也行）。
        head = tk.Frame(wrap, bg=theme.SURFACE, cursor="hand2")
        head.pack(fill="x", padx=12, pady=(0, 4))
        title_lbl = tk.Label(head, text=title, bg=theme.SURFACE, fg=theme.PRIMARY,
                            font=self.F["F_LABEL"], anchor="w", cursor="hand2")
        title_lbl.pack(side="left")
        mark_lbl = tk.Label(head, text="▾" if opened else "▸",
                           bg=theme.SURFACE, fg=theme.PRIMARY,
                           font=self.F["F_LABEL"], anchor="e", cursor="hand2")
        mark_lbl.pack(side="right")
        holder = tk.Frame(wrap, bg=theme.SURFACE)
        if opened:
            holder.pack(fill="x")
        st = {"open": opened}

        def _toggle(_e=None):
            st["open"] = not st["open"]
            if st["open"]:
                holder.pack(fill="x")
                mark_lbl.configure(text="▾")
            else:
                holder.pack_forget()
                mark_lbl.configure(text="▸")
            if on_toggle:
                try:
                    on_toggle()
                except Exception:
                    pass
            return "break"

        head.bind("<Button-1>", _toggle)
        title_lbl.bind("<Button-1>", _toggle)
        mark_lbl.bind("<Button-1>", _toggle)
        return holder, wrap

    def _execute_check(self, target, win=None):
        """真正跑体检（复核确认后调用）：原 _run 的检查逻辑。"""
        if self._busy:
            return
        tr = self.tr
        docx = self.docx_path.get()
        if win is not None:
            try:
                self._close_one(win)
            except Exception:
                pass
        self._phase = "ready"
        self._set_busy(True)
        self._set_status(tr("status_checking"), "run")
        self._set_bar("run")
        # v2.3.32：把状态行/引导条的最后一次改动也画进去，再进阻塞的引擎调用。
        self._flush_frame()
        try:
            prof, findings = checker.run_check(docx, target)
            rep = report_mod.build_report(docx, target, prof, findings)
            self._last_report = rep

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
        # v2.3.32：同上，进阻塞的 compute_changes 之前先画成正确帧。
        self._flush_frame()
        # 整段「算改动 → 授权裁决 → 预览」都在 try 里，由 finally 收忙态：
        # 任何一处抛异常（例如 gate 结构异常、文案格式化失败）都不能让
        # `_busy` 卡在 True —— 那会让主按钮永久停在 "Checking…"。
        gate_no_go = ""
        preview_text = ""
        try:
            target = self._build_target()
            self._fix_findings = None
            _chk_prof, _chk_findings = checker.run_check(docx, target)
            self._fix_findings = _chk_findings
            changes = _compute_changes(_chk_prof, target, self.lang,
                                      findings=_chk_findings)

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
                gate_no_go = self._gate_message(gate)
                self._set_status(gate_no_go, "err")
                self._set_bar("err")
                # 注意：此处绝不能再 return。旧代码在这里 return 会直接
                # 退出整个函数，导致下方的「授权不足」分支（弹合并激活窗
                # 引导购买）永不执行 —— 用户只看到红状态栏、却找不到买码入口。
                # 正确做法：try 正常结束 → finally 解除忙态 → 再弹窗。
            else:
                preview_text = (
                    tr("preview_header") + "\n"
                    + "\n".join("  · " + line for line in changes)
                    + "\n\n" + tr("preview_trial", self._gate_message(gate)))
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

        # 授权不足：购买/重新激活决策 → 直接给合并激活窗（含购买+激活码），
        # 保证用户一定看见；state_error 才用普通警告（那是程序故障，不是购买决策）。
        # （必须等忙态解除后再弹，否则模态框会挡住还没刷新的按钮状态）
        if gate_no_go:
            reason = (gate or {}).get("reason")
            if reason in ("revoked", "needs_reactivation", "trial_exhausted"):
                # 同上：状态行要跟着变成"已撤销"，并露出 Upgrade 入口
                self._refresh_commercial()
                self._upgrade_dialog()
            else:
                messagebox.showwarning(tr("msg_need_activation_title"), gate_no_go)
            return

        # v2.3.83：从这里开始连弹三个**Windows 原生**对话框（确认 / 另存为 / 完成），
        # 而原生 shell 对话框会把进程 DPI 感知提升、让窗口停止位图拉伸（变小+变清晰
        # 且退不回去，老板 2026-10-21 报的就是这个）。先抓一份显示状态，返回后自愈。
        _disp = self._snap_display()
        if not messagebox.askyesno(tr("msg_fix_confirm_title"),
                                   preview_text + tr("msg_fix_confirm_body")):
            self._heal_display(_disp)
            self._set_bar("idle")
            return

        default_name = os.path.splitext(os.path.basename(docx))[0] + "_fixed.docx"
        out = filedialog.asksaveasfilename(
            title=tr("msg_fix_pick_output"), defaultextension=".docx",
            filetypes=[("Word", "*.docx")], initialfile=default_name,
            initialdir=os.path.dirname(docx) or None)
        if not out:
            self._heal_display(_disp)
            self._set_bar("idle")
            return
        # 保存对话框已自带覆盖确认，这里只拦「与原件同一文件」这一种不可逆情形
        if same_file(out, docx):
            messagebox.showerror(tr("msg_cannot_overwrite_title"),
                                 tr("msg_cannot_overwrite"))
            self._heal_display(_disp)
            self._set_bar("idle")
            return

        self._set_busy(True)
        try:
            fix_docx(docx, out, target,
                     findings=self._fix_findings, spec=checker.resolve_spec(target))
            counted = lic.record_fix_used()
            self._last_output = out
            self._phase = "fixed"
        except DocxReadError as e:
            self._heal_display(_disp)
            self._fail(tr("msg_read_error_title"), str(e))
            return
        except (ValueError, OSError) as e:
            self._heal_display(_disp)
            self._fail(tr("msg_fix_error_title"), str(e))
            return
        finally:
            self._set_busy(False)

        self._sync_ui()                      # 主按钮 → Review changes，三步全部打勾
        self._set_status(tr("state_fixed"), "ok", spec=("state_fixed", ()))
        self._set_bar("ok")
        self._refresh_commercial()           # 扣次后刷新状态行（次卡剩余次数 / 近到期提示）
        # 展示**扣减后**的状态（gate['message'] 是扣减前的，会显示成「首次免费试用」）
        messagebox.showinfo(tr("msg_fix_done_title"),
                            "%s\n%s\n\n%s" % (tr("msg_fix_done_before_file"), out,
                                              lic.post_fix_message(gate, counted)))
        # 三个原生对话框全部走完 —— 最后一次自愈，客户看到的窗口纹丝不动
        self._heal_display(_disp)

    # ------------------------------------------------------------ 报告 / 模板
    def _save_report(self):
        tr = self.tr
        content = self._report_text.strip()
        if not content:
            messagebox.showwarning(tr("save_no_report_title"), tr("save_no_report"))
            return
        _disp = self._snap_display()
        dst = filedialog.asksaveasfilename(
            title=tr("save_report_dialog"), defaultextension=".docx",
            filetypes=[("Word document", "*.docx"), ("Markdown", "*.md"),
                       ("All files", "*.*")],
            initialfile="format-check-report.docx")
        if not dst:
            self._heal_display(_disp)
            return
        try:
            if dst.lower().endswith(".docx") and getattr(self, "_last_report", None) is not None:
                # Render a structured Word report (keeps the GUI English-only
                # and mirrors the Markdown report's section order).
                report_mod.build_docx_report(self._last_report, dst)
            else:
                with open(dst, "w", encoding="utf-8") as fh:
                    fh.write(content + "\n")
        except OSError as e:
            self._heal_display(_disp)
            messagebox.showerror(tr("msg_fix_error_title"), str(e))
            return
        messagebox.showinfo(tr("save_report_done_title"),
                            "%s\n%s" % (tr("save_report_done"), dst))
        self._heal_display(_disp)

    def _export_template(self):
        tr = self.tr
        _disp = self._snap_display()
        dst = filedialog.asksaveasfilename(
            defaultextension=".yaml", filetypes=[("YAML", "*.yaml *.yml"),
                                                 ("All files", "*.*")],
            initialfile="ai_questionnaire_template.yaml")
        if not dst:
            self._heal_display(_disp)
            return
        try:
            with open(_TEMPLATE_PATH, "r", encoding="utf-8") as src, \
                    open(dst, "w", encoding="utf-8") as out:
                out.write(src.read())
        except OSError as e:
            self._heal_display(_disp)
            messagebox.showerror(tr("msg_fix_error_title"), str(e))
            return
        messagebox.showinfo(tr("export_done_title"), tr("export_done", dst))
        self._heal_display(_disp)


def main():
    # v2.3.83：DPI 感知锁定（老板 2026-10-21 报「保存修正稿时窗口变小、字变清楚、
    # 一直保持到关软件」）。根因与取证见 src/ui/dpi.py 顶部说明。一句话：
    # 进程默认 DPI 不感知 → 150% 屏上整窗被 Windows 拉伸；弹原生「另存为」框时
    # 进程感知被提升 → 停止拉伸 → 窗口按物理像素呈现（缩小 + 变清晰），且进程级
    # 设定退不回去。这里**在创建任何窗口之前**显式锁成 UNAWARE：实测锁定后任何
    # 提升尝试（V2 / GDIscaled）都返回 False，感知保持 UNAWARE —— 显示方式与现状
    # 完全一致，只是再也跳不了。必须在 tk.Tk() 之前，否则窗口已按旧感知建好。
    _dpi_mod.pin_process_dpi_awareness()
    root = tk.Tk()
    root.withdraw()          # 消掉启动一闪：先隐藏主窗，等 UI 全部建完再显示
    app = App(root)
    # v2.3.82 启动花屏（老板 2026-10-21 报）。
    #
    # 抓屏取证（_smoke/probe_startup_flicker.py）与时间轴（probe_startup_timeline.py）
    # 说明问题：deiconify 一到窗口就被系统映射上屏，但它此时**还没画过**，要等映射后
    # 的布局收敛 + 首帧 flush（实测空窗 1.2 秒）才有内容；期间屏幕上是空白纸壳，
    # 以及"右卡片压在左卡片上、文字重影"的错位叠影（截图见
    # _smoke/_startup_frames_noshroud/f028_t1521ms.png）。
    #
    # 为什么不能简单"先排好再显示"：窗口真实客户区尺寸只有**映射之后**才由系统给
    # （隐藏时 Tk 报 1x1；强行在隐藏状态排，会让 Tk 提前兑现几何 → 映射后窗口永久
    # 停在 1x1，实测过，是 P0）。也不能用 ``-alpha 0`` 当遮罩：Tk 在 alpha=0 时根本不
    # 产生重绘，调回 1.0 后窗口永久空白（实测过，也是 P0）。这两条都写进了
    # _smoke/probe_startup_paint.py 的注释里，别再试。
    #
    # 现方案：**在屏幕外完成映射与首帧绘制，画完再移入最终位置**。窗口尺寸/布局走的是
    # 完全正常的路径（不会被 1x1 卡住），只是这段时间它在屏幕外，客户看不到；画完后
    # 的移动是纯位置变化，Windows 不需要重绘客户区 —— 客户看到的第一眼就是成品。
    # 三重保险，任何一环失败都只会退回"跟旧版一样"，不会让窗口消失：
    #   ①挪不动位置就留在原位；②首帧钩子没触发时，4 秒后强制挪回；③移入幂等。
    _final_geo = getattr(app, "_start_geometry", None)
    _moved_off = False
    if _final_geo:
        try:
            root.geometry("+20000+20000")   # 只改位置，尺寸沿用 App 请求值
            _moved_off = True
        except Exception:
            _moved_off = False
    root.deiconify()         # UI 构建完毕，正式显示
    # 强制窗口可见并置于前台/激活：缩放屏 + Nuitka 打包下，deiconify 有时不会让 OS 真正
    # 映射/激活窗口，导致客户区一直不绘制，直到用户点击上边缘才激活并画出来（"需点击才出"
    # 的典型激活类根因）。lift + focus_force + 短暂 topmost 是最稳的强制可见手段。
    try:
        root.lift()
        root.update_idletasks()
        root.focus_force()
        root.attributes("-topmost", True)
        root.after(120, lambda: root.attributes("-topmost", False))
    except Exception:
        pass
    # v2.3.47：图标必须在窗口映射之后才设得住。App.__init__ 里那次发生在 withdraw()
    # 状态、真机实测无效（WM_GETICON 恒 0），所以 deiconify 之后再补一次。
    # 多排几个时点，兼顾「映射消息已到达」与「首次布局尚未把它覆盖」。
    for _d in (0, 80, 400):
        try:
            root.after(_d, app._reapply_window_icon)
        except Exception:
            pass

    if _moved_off and _final_geo:
        # 只取位置部分（"+370+8"）：移回时**连尺寸都不碰**，免得触发 resize → 重绘。
        _pos = ""
        _i = _final_geo.find("+")
        if _i > 0:
            _pos = _final_geo[_i:]
        _state = {"done": False}

        def _place_on_screen():
            """把窗口移入最终位置（幂等；只改位置，不重绘客户区）。"""
            if _state["done"]:
                return
            _state["done"] = True
            try:
                root.geometry(_pos or _final_geo)
                root.update_idletasks()
            except Exception:
                pass

        # ①首帧画完立刻移回；②兜底：万一钩子没触发（异常/几何一直没就绪），4 秒必回
        app.on_first_frame_ready(_place_on_screen)
        root.after(4000, _place_on_screen)

    # 首帧布局：映射后的 <Configure> 收敛、排版、整窗重绘都在这一句里跑完
    # （窗口在屏幕外时，客户什么都看不到）。
    app._schedule_first_layout()
    root.mainloop()


if __name__ == "__main__":
    main()
