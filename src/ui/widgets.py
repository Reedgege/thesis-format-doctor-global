"""自绘控件：圆角卡片 + 步进圆点。

**为什么要自绘**：tkinter 的 ``Frame`` 只有直角，原生没有 ``border-radius``，
也没有 ``box-shadow``。规格书要求 12px 圆角 + 轻微投影，所以这里用 ``Canvas``
自己把圆角矩形画出来，再把内容 ``Frame`` 作为 canvas window item 放进去。

**两个关键实现细节**（都是踩过才知道的）：

1. ``Canvas`` 的**默认请求尺寸是 378×278**，不显式压到 ``width=1, height=1``
   的话，卡片会凭空撑出 378 宽、把布局顶乱。

2. 卡片的**布局高度必须由内容驱动**（否则窗口变小时两栏会塌成一条缝）。
   所以 ``inner`` 一变就把自身请求高度回灌给 canvas（``_sync_request``）；
   canvas 实际尺寸一变又把宽度回灌给 ``inner``（``_relayout``）。
   两个方向都带"值不变就跳过"的判断——否则 ``<Configure>`` 会来回触发成死循环
   （文字换行会随宽度变化而改变请求高度，这是最容易滚起来的路径）。
"""

from __future__ import annotations

import math

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

from . import theme


# ---------------------------------------------------------------------------
# 圆角矩形路径
# ---------------------------------------------------------------------------
def rounded_points(x0: float, y0: float, x1: float, y1: float,
                   r: float, steps: int = 7) -> list:
    """返回圆角矩形的多边形顶点（真实圆弧采样，不是样条近似）。

    从右下角开始顺时针绕一圈：右下 → 左下 → 左上 → 右上。
    ``steps`` 是每个圆角的采样段数，7 段在 12px 圆角上肉眼已看不出棱角。
    """
    r = max(0.0, min(r, (x1 - x0) / 2.0, (y1 - y0) / 2.0))
    steps = max(1, int(steps))          # steps=0 会除零
    if r <= 0.5:
        return [x0, y0, x1, y0, x1, y1, x0, y1]
    pts: list = []
    # (圆心x, 圆心y, 起始角度)——角度用屏幕坐标（y 向下），顺时针递增
    corners = (
        (x1 - r, y1 - r, 0.0),      # 右下：右 → 下
        (x0 + r, y1 - r, 90.0),     # 左下：下 → 左
        (x0 + r, y0 + r, 180.0),    # 左上：左 → 上
        (x1 - r, y0 + r, 270.0),    # 右上：上 → 右
    )
    for cx, cy, a0 in corners:
        for i in range(steps + 1):
            a = math.radians(a0 + 90.0 * i / steps)
            pts.append(cx + r * math.cos(a))
            pts.append(cy + r * math.sin(a))
    return pts


# ---------------------------------------------------------------------------
# 圆角卡片
# ---------------------------------------------------------------------------
class RoundCard(tk.Frame):
    """圆角卡片容器。内容放进 ``card.inner``。

    参数
    ----
    radius : 圆角半径（px）
    fill   : 卡片填充色
    border : 边框色
    shadow : 是否画一层轻微投影
    padx/pady : 内容与卡片边缘的距离
    """

    def __init__(self, parent, *, radius: int = theme.CARD_RADIUS,
                 fill: str = theme.SURFACE, border: str = theme.BORDER,
                 shadow: bool = True, padx: int = theme.CARD_PAD_X,
                 pady: int = theme.CARD_PAD_Y, bg: str = None,
                 min_height: int = 0, dashed: bool = False):
        bg = bg or parent.cget("bg")
        super().__init__(parent, bg=bg, highlightthickness=0, bd=0)
        self._radius = int(radius)
        self._fill = fill
        self._border = border
        self._shadow = bool(shadow)
        self._dashed = bool(dashed)
        self._bg = bg
        self._padx = int(padx)
        self._pady = int(pady)
        self._min_h = int(min_height)
        self._req = (0, 0)

        # width/height=1：压掉 Canvas 的 378×278 默认请求尺寸（见模块 docstring）
        self._cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                             width=1, height=1)
        self._cv.pack(fill="both", expand=True)

        self.inner = tk.Frame(self._cv, bg=fill)
        self._win = self._cv.create_window(self._padx, self._pady,
                                           window=self.inner, anchor="nw")

        self._cv.bind("<Configure>", self._relayout)
        self.inner.bind("<Configure>", self._sync_request)
        # 首帧自愈：canvas 的 <Configure> 有可能早于"内容装进 inner"发生（量到 1px），
        # 之后再没有任何事件能把它唤醒 —— 虚线投放区就被压成一条线。idle 时补量一次；
        # 窗口在 idle 前销毁则取消，避免触发已删除的 Tcl 命令。
        self._idle_id = self.after_idle(self._sync_request)
        self.bind("<Destroy>", self._cancel_idle, add="+")

    def _cancel_idle(self, _e=None):
        aid = getattr(self, "_idle_id", None)
        if aid is not None:
            try:
                self.after_cancel(aid)
            except Exception:
                pass
            self._idle_id = None

    # ---------------------------------------------------------------- 内部
    def _sync_request(self, _e=None):
        """内容尺寸变化 → 回灌 canvas 的请求尺寸（决定布局占多高/多宽）。"""
        try:
            w = self.inner.winfo_reqwidth() + 2 * self._padx
            h = max(self.inner.winfo_reqheight() + 2 * self._pady, self._min_h)
        except Exception:
            return
        if (w, h) == self._req:
            return                      # 值没变就停，否则会来回触发成死循环
        # 宽度请求封顶：卡片都长在 fill="x" / sticky="nsew" 的格子里，真实宽度由窗口
        # 决定；而自动换行标签的 wraplength 又反过来按格宽算 —— 把内容宽度原样发布
        # 出去会和 grid 列宽形成正反馈（两列 506↔669 无限互推、update_idletasks 卡死）。
        # 封顶后列宽只由窗口决定，换行再也不会反向撑宽布局。
        w = min(w, theme.CARD_REQ_W)
        if (w, h) == self._req:
            return
        self._req = (w, h)
        try:
            self._cv.configure(width=w, height=h)
        except Exception:
            pass

    def _relayout(self, e=None):
        """canvas 实际尺寸变化 → 重画圆角 + 让内容宽度（并尽量撑满高度）跟上。

        高度给 ``max(内容自然高度, 卡片高度 - 上下留白)``：卡片被 grid 拉得比内容高时，
        内容区跟着撑满，这样内部 ``side="bottom"`` 的部件（信任卡）才能贴到卡片底部，
        而不是紧贴在上一行后面。内容比卡片高时仍用自然高度（宁可溢出也不裁切）。
        """
        try:
            w = e.width if e is not None else self._cv.winfo_width()
            h = e.height if e is not None else self._cv.winfo_height()
        except Exception:
            return
        # 高度可能还是首帧的 1px：pack 只按"请求高度"给空间，而请求高度要等内容量完
        # 才有值。以前这里连 h<=2 一起早退 → item 永远停在 1×1、inner 的 <Configure>
        # 再也不触发，圆形/虚线卡片（上传投放区）就这么被压成一条线。
        if w <= 2:
            return
        if h > 2:
            self._draw(w, h)
        try:
            natural = self.inner.winfo_reqheight()
            self._cv.itemconfigure(
                self._win,
                width=max(1, w - 2 * self._padx),
                height=max(natural, h - 2 * self._pady))
        except Exception:
            pass

    def _draw(self, w: int, h: int):
        cv = self._cv
        try:
            cv.delete("card")
        except Exception:
            return
        r = self._radius
        if self._shadow:
            # 投影：同一圆角矩形下移 2px，用比背景略深的色画在卡片下面一层。
            # 只露底部 2px + 两侧极窄，观感上就是一层轻投影。
            sh = rounded_points(1, 3, w - 1, h - 1 + 2, r)
            cv.create_polygon(sh, fill=theme.SHADOW, outline=theme.SHADOW,
                              tags="card")
        pts = rounded_points(0.5, 0.5, w - 0.5, h - 0.5, r)
        if self._dashed:
            # 虚线圈：填充走多边形，描边走一条**首尾闭合**的虚线折线
            # （tkinter 的 polygon 不支持 dash，所以拆成两笔画）
            cv.create_polygon(pts, fill=self._fill, outline="", tags="card")
            cv.create_line(*(list(pts) + pts[:2]), fill=self._border,
                           dash=(5, 4), width=1, tags="card")
        else:
            cv.create_polygon(pts, fill=self._fill, outline=self._border,
                              width=1, tags="card")
        cv.tag_lower("card")

    # ---------------------------------------------------------------- 对外
    def refresh(self):
        """内容尺寸可能变了 → **重新测量**并重画。

        为什么必须显式调：``inner`` 是画布里的 window item，它的宽高由我们钉死
        （见 ``_relayout``），所以"只往里面加内容"时**不会**触发 ``inner.<Configure>``，
        ``_sync_request`` 也就不跑 —— 结果是卡片不涨高、新内容被圆角矩形裁掉，
        而且外层滚动区也拿不到新的 ``scrollregion``，用户既看不到也滚不到。
        （Codex 2026-09-12 P0-2 指出的正是这条路径：展开 Advanced options 后
        最下面一行不可达。）

        先 ``update_idletasks`` 让刚 pack 的子控件把请求尺寸算出来，再测量。
        """
        try:
            self.inner.update_idletasks()
        except Exception:
            pass
        try:
            self._sync_request()
        except Exception:
            pass
        try:
            self._relayout()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 竖向滚动区
# ---------------------------------------------------------------------------
class ScrollArea(tk.Frame):
    """竖向滚动区：内容放进 ``.inner``；**装得下时滚动条自动隐藏**。

    为什么主区域需要它：左卡（4 个字段 + 高级折叠区）在 1100 宽下自然高度约 700px，
    而 1366×768 / 1600×900 这类笔记本的可视高度只有 500~600px。没有滚动区的话，
    卡片底部会被圆角矩形**直接裁掉**——用户既看不到字段也看不到滚动条，是硬 bug。

    另一个作用：内容比视口矮时把 ``inner`` 拉到视口高度，外层 grid 的行权重才能把
    左右两栏拉成**等高**（对应规格 "Keep left/right cards visually balanced"）。

    滚轮处理刻意**不用独占绑定**：只暴露 ``handle_wheel``，由 ``App`` 统一挂一次
    ``bind_all(..., add="+")`` 再分发过来，回调里先判断指针是否真的在本区域内。
    自己绑会有两个坑：切语言重建界面时旧绑定残留，以及跟弹窗的滚轮绑定互相顶掉。
    """

    def __init__(self, parent, bg: str = theme.BG):
        super().__init__(parent, bg=bg)
        self._bar_shown = False
        self._cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                             width=1, height=1)
        self._vsb = ttk.Scrollbar(self, orient="vertical", command=self._cv.yview)
        self._cv.configure(yscrollcommand=self._on_scroll)
        self._cv.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self._cv, bg=bg)
        self._win = self._cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._sync_region)
        self._cv.bind("<Configure>", self._relayout)

    # ---------------------------------------------------------------- 滚动条
    def _on_scroll(self, first, last):
        try:
            need = not (float(first) <= 0.0 and float(last) >= 1.0)
        except Exception:
            need = True
        self._set_bar(need)
        try:
            self._vsb.set(first, last)
        except Exception:
            pass

    def _set_bar(self, show: bool):
        if show == self._bar_shown:
            return
        self._bar_shown = show
        try:
            if show:
                # before=self._cv：保证滚动条在 canvas 右侧，而不是被挤到下面
                self._vsb.pack(side="right", fill="y", before=self._cv)
            else:
                self._vsb.pack_forget()
        except Exception:
            pass

    # ---------------------------------------------------------------- 布局
    def _sync_region(self, _e=None):
        try:
            self._cv.configure(scrollregion=self._cv.bbox("all"))
        except Exception:
            pass

    def _relayout(self, e=None):
        try:
            w = e.width if e is not None else self._cv.winfo_width()
            h = e.height if e is not None else self._cv.winfo_height()
        except Exception:
            return
        if w <= 2 or h <= 2:
            return
        try:
            natural = self.inner.winfo_reqheight()
            # 宽度跟随视口；高度至少撑满视口，让外层 grid 能把两栏拉成等高
            self._cv.itemconfigure(self._win, width=w, height=max(natural, h))
        except Exception:
            pass
        self._sync_region()

    # ---------------------------------------------------------------- 滚轮
    def _in_area(self, e) -> bool:
        try:
            w = self._cv.winfo_containing(e.x_root, e.y_root)
        except Exception:
            return False
        while w is not None:
            if w is self._cv:
                return True
            try:
                w = w.master
            except Exception:
                return False
        return False

    def handle_wheel(self, e):
        """由 ``App`` 的全局滚轮回调分发进来；不在本区域内则什么都不做。"""
        if not self._bar_shown or not self._in_area(e):
            return
        try:
            # 精密触控板的 delta 常见 ±30/±60，用 int(delta/120) 会被截成 0
            # 而"滚不动"，所以只取方向。
            self._cv.yview_scroll(-1 if e.delta > 0 else 1, "units")
        except Exception:
            pass

    def scroll_to_top(self):
        try:
            self._cv.yview_moveto(0)
        except Exception:
            pass

    def refresh_layout(self):
        """内容尺寸可能变了 → 重算 item 尺寸与 ``scrollregion``。

        与 ``RoundCard.refresh`` 同理：``inner`` 的尺寸被钉死，往里加内容不会触发
        ``<Configure>``，所以必须由外部在 pack/unpack/换文案之后主动喊一声。
        """
        try:
            self.inner.update_idletasks()
        except Exception:
            pass
        try:
            self._relayout()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 圆角按钮
# ---------------------------------------------------------------------------
class RoundButton(tk.Frame):
    """圆角按钮（Canvas 自绘，9px 圆角）。

    为什么不用 ``ttk.Button``：规格第 5 节第 8 条要求「主按钮必须视觉上绝对主导」，
    而 ttk 的按钮在 Windows 上要么是系统灰方块、要么圆角/配色都调不动。
    自己做还能顺手拿到 hover 态与置灰态。

    对外接口刻意做成"傻瓜式"，调用方不碰 canvas：
    ``set_text`` / ``set_enabled`` / ``set_action`` / ``set_visible``。

    实现上本类是 ``tk.Frame``，真正画图的是内部 canvas —— 这样 ``set_visible``
    可以直接对这个 Frame 做 pack/pack_forget，不用去猜调用方当初用了什么 pack 参数
    （所以这里覆写了 ``pack`` 把参数记下来）。
    """

    STYLES = {
        # style: (底色, 文字色, 边框色, hover底色)
        "primary": (theme.PRIMARY, "#FFFFFF", theme.PRIMARY, theme.PRIMARY_HOVER),
        "secondary": (theme.SURFACE, theme.PRIMARY, theme.SECONDARY, theme.PRIMARY_SOFT),
        "ghost": (None, theme.PRIMARY, None, theme.PRIMARY_SOFT),
    }

    def __init__(self, parent, text: str, command=None, *, style: str = "primary",
                 font=None, radius: int = 9, height: int = 52, padx: int = 20,
                 bg: str = None, icon: str = None):
        bg = bg or parent.cget("bg")
        super().__init__(parent, bg=bg, highlightthickness=0, bd=0)
        self._style = style if style in self.STYLES else "primary"
        self._radius = int(radius)
        self._h = int(height)
        self._padx = int(padx)
        self._font = font
        self._text = text
        self._icon = icon
        self._command = command
        self._enabled = True
        self._hover = False
        self._visible = True
        self._pack_opts: dict = {"fill": "x"}

        self._cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                             height=self._h, width=self._fit_width(),
                             takefocus=1)
        self._cv.pack(fill="both", expand=True)
        self._cv.bind("<Configure>", lambda e: self.draw())
        self._cv.bind("<Enter>", self._on_enter)
        self._cv.bind("<Leave>", self._on_leave)
        self._cv.bind("<Button-1>", self._on_click)
        self._cv.bind("<Return>", self._on_click)
        self._cv.bind("<space>", self._on_click)
        self.draw()

    # ---------------------------------------------------------------- pack
    def pack(self, **kw):                      # noqa: D102  （见类 docstring）
        self._pack_opts = dict(kw)
        super().pack(**kw)

    def pack_anchor(self, widget):
        """记下 ``set_visible(True)`` 重新 pack 时的锚点（插到 ``widget`` 之前）。

        主/次 CTA 的摆法在建卡时定不下来（锚点控件可能后建），显隐又要靠
        ``set_visible`` 重新入队，所以单独给一个"事后补锚点"的入口。
        """
        if widget is not None:
            self._pack_opts = dict(self._pack_opts)
            self._pack_opts["before"] = widget

    # ---------------------------------------------------------------- 绘制
    def draw(self):
        cv = self._cv
        try:
            cv.delete("all")
            w = cv.winfo_width()
            h = cv.winfo_height()
        except Exception:
            return
        if w <= 2 or h <= 2:
            return
        fill, fg, border, hover_fill = self.STYLES[self._style]
        if not self._enabled:
            fill, fg, border = theme.BTN_DISABLED_BG, theme.BTN_DISABLED_FG, None
        elif self._hover and hover_fill:
            fill = hover_fill
            if self._style == "secondary":
                fg = theme.PRIMARY
        pts = rounded_points(0.5, 0.5, w - 0.5, h - 0.5, self._radius)
        try:
            if fill:
                cv.create_polygon(pts, fill=fill,
                                  outline=border if border else fill, width=1)
            elif border:                       # ghost：只描边不填充
                cv.create_polygon(pts, fill="", outline=border, width=1)
            cx, cy = w / 2.0, h / 2.0
            if self._icon:
                # 规格稿的主 CTA 是「图标 + 文字」整体居中：先量文字宽度，把图标放到
                # 文字左侧，再把文字整体右移半个图标位。
                try:
                    tw = tkfont.Font(font=self._font).measure(self._text)
                except Exception:
                    tw = len(self._text) * 8
                draw_icon(cv, self._icon, cx - tw / 2.0 - 15, cy, 16, fg, 2)
                cx += 8
            cv.create_text(cx, cy, text=self._text, fill=fg,
                           font=self._font, width=max(40, w - 2 * self._padx))
        except Exception:
            pass

    # ---------------------------------------------------------------- 交互
    def _on_enter(self, _e=None):
        self._hover = True
        try:
            self._cv.configure(cursor="hand2" if self._enabled else "arrow")
        except Exception:
            pass
        self.draw()

    def _on_leave(self, _e=None):
        self._hover = False
        self.draw()

    def _on_click(self, _e=None):
        if not self._enabled or self._command is None:
            return
        try:
            self._cv.focus_set()
        except Exception:
            pass
        try:
            self._command()
        except Exception:
            pass

    # ---------------------------------------------------------------- 对外
    def set_text(self, text: str):
        text = str(text)
        if text != self._text:
            self._text = text
            # 不给 fill="x" 的摆法（如商业区的 Upgrade）靠请求宽度定宽，改字要跟着改
            try:
                self._cv.configure(width=self._fit_width())
            except Exception:
                pass
            self.draw()

    def _fit_width(self) -> int:
        """按文字实测宽度算请求宽度（+ 图标位 + 左右内边距）。

        ttk 按钮自己会按文字算宽；自绘按钮得自己算 —— 否则 ``pack(side="right")``
        这类不给 fill 的摆法会把按钮压成 1px（商业区 Upgrade 按钮踩过这个坑）。
        """
        try:
            tw = tkfont.Font(font=self._font).measure(self._text)
        except Exception:
            tw = len(self._text) * 8
        extra = 26 if self._icon else 0
        return max(24, int(tw) + extra + 2 * self._padx)

    @property
    def text(self) -> str:
        return self._text

    def set_action(self, command):
        self._command = command

    def set_enabled(self, enabled: bool):
        enabled = bool(enabled)
        if enabled == self._enabled:
            return
        self._enabled = enabled
        self._hover = False
        try:
            self._cv.configure(cursor="hand2" if enabled else "arrow")
        except Exception:
            pass
        self.draw()

    @property
    def enabled(self) -> bool:
        return self._enabled

    def set_visible(self, visible: bool):
        visible = bool(visible)
        if visible == self._visible:
            return
        self._visible = visible
        try:
            if visible:
                super().pack(**self._pack_opts)
            else:
                self.pack_forget()
        except Exception:
            pass
        # 卡片高度由内容驱动，显隐后要让外层 RoundCard 重新量一次
        try:
            self.update_idletasks()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 极简进度条
# ---------------------------------------------------------------------------
class ProgressBar(tk.Canvas):
    """极简进度条：轨道 + 已完成填充 + 步骤圆点 + 步骤标签。

    替代旧版的横向 ``StepCircle`` 步进器（用户反馈原界面「太杂乱」）。只画一条细条，
    不堆大段文字；``set_fraction`` / ``set_step`` 驱动进度，``frac`` 供测试断言。

    设计：``steps`` 个步骤圆点均匀分布在轨道上（第 1 个在左端、最后一个在右端），
    填充比例 ``frac`` 从左向右生长，已越过 / 已完成的圆点刷成主色、其余为轨道灰；
    每个圆点下方一行小标签（Prepare / Check / Fix 之类），当前及已过步骤用主色。
    """

    def __init__(self, parent, steps: int = 3, bg: str = None, font=None,
                 height: int = 34, track: str = None, fill: str = None, **kw):
        bg = bg or parent.cget("bg")
        self._steps = max(1, int(steps))
        self._font = font
        self._track = track or theme.HAIRLINE
        self._fill = fill or theme.PRIMARY
        self._labels: list = []
        self.frac = 0.0                     # 公开属性，供测试断言
        super().__init__(parent, height=height, bg=bg,
                         highlightthickness=0, bd=0, **kw)
        self.bind("<Configure>", lambda e: self.draw())
        self.draw()

    def set_fraction(self, f: float):
        self.frac = max(0.0, min(1.0, float(f)))
        self.draw()
        return self.frac

    def set_step(self, index: int):
        """``index`` = 已完成步骤数（0..steps）。"""
        self.set_fraction(index / self._steps)

    def set_labels(self, labels):
        self._labels = list(labels) if labels else []
        self.draw()

    def draw(self):
        try:
            self.delete("all")
            w = self.winfo_width()
            h = self.winfo_height()
        except Exception:
            return
        if w <= 2 or h <= 2:
            return
        pad = 14
        bar_y = 9
        bar_h = 5
        radius = bar_h / 2.0
        n = self._steps

        # 轨道
        track_pts = rounded_points(pad, bar_y, w - pad, bar_y + bar_h, radius)
        self.create_polygon(track_pts, fill=self._track, outline=self._track,
                            tags="bar")
        # 填充
        fx = pad + (w - 2 * pad) * self.frac
        if fx > pad + 0.5:
            fill_pts = rounded_points(pad, bar_y, fx, bar_y + bar_h, radius)
            self.create_polygon(fill_pts, fill=self._fill, outline=self._fill,
                                tags="bar")

        # 步骤圆点 + 标签
        if n > 1:
            mid_y = bar_y + bar_h + 13
            for i in range(n):
                pos = i / (n - 1)
                cx = pad + (w - 2 * pad) * pos
                completed = (i + 1) <= round(self.frac * n)
                col = self._fill if completed else self._track
                self.create_oval(cx - 4, bar_y + bar_h / 2 - 4,
                                 cx + 4, bar_y + bar_h / 2 + 4,
                                 fill=col, outline=col, tags="bar")
                if self._labels and self._font and i < len(self._labels):
                    name = self._labels[i]
                    try:
                        tw = tkfont.Font(font=self._font).measure(name)
                    except Exception:
                        tw = len(name) * 7
                    cx = max(tw / 2.0 + 2, min(w - tw / 2.0 - 2, cx))
                    self.create_text(cx, mid_y, text=name, fill=col,
                                     font=self._font, anchor="n", tags="bar")


# ---------------------------------------------------------------------------
# 线描图标（纯 Canvas 自绘）
# ---------------------------------------------------------------------------
# 规格包里的图标全是单色线描件，且**不许引入第三方图形库**（禁 PIL、不许内嵌图片）。
# 这里用一个 24×24 的设计网格 + 统一缩放把常用图标画出来：调用方只给中心点与边长，
# 内部所有坐标都按 ``P(x, y)`` 换算，换尺寸不会走形。
#
# 为什么值得有：设计稿每个字段行、每个面板标题前都有一枚小图标，缺了它整页会显得
# 「只有字没有结构」。图标是**装饰**，任何一步失败都静默跳过（绝不能拖垮界面）。
def draw_icon(cv, name: str, cx: float, cy: float, size: float,
              color: str, width: int = 2, tags=()) -> bool:
    """在画布 ``cv`` 上以 (cx, cy) 为中心画一枚 ``size`` 边长的线描图标。"""
    s = float(size) / 24.0

    def P(x, y):
        return (cx + (x - 12.0) * s, cy + (y - 12.0) * s)

    def line(*pts):
        flat = []
        for x, y in pts:
            flat.extend(P(x, y))
        cv.create_line(*flat, fill=color, width=width, capstyle="round",
                       joinstyle="round", tags=tags)

    def oval(x0, y0, x1, y1, fill="", w=None):
        a, b = P(x0, y0), P(x1, y1)
        cv.create_oval(a[0], a[1], b[0], b[1], outline=color, fill=fill,
                       width=(width if w is None else w), tags=tags)

    def poly(*pts, fill="", outline=None, w=None):
        flat = []
        for x, y in pts:
            flat.extend(P(x, y))
        cv.create_polygon(*flat, fill=fill,
                          outline=(color if outline is None else outline),
                          width=(width if w is None else w), tags=tags)

    try:
        if name == "doc":
            poly((7, 3.5), (15.5, 3.5), (19, 7), (19, 20.5), (7, 20.5))
            line((15, 3.8), (15, 7.2), (18.6, 7.2))
            line((10, 12), (16, 12))
            line((10, 16), (16, 16))
        elif name == "quote":
            for dx in (0, 7.5):
                oval(5.5 + dx, 8.5, 11 + dx, 14, fill=color, w=0)
                poly((5.6 + dx, 13.4), (8.4 + dx, 13.4), (6.2 + dx, 17.6),
                     fill=color, outline="", w=0)
        elif name == "bank":
            poly((3, 10), (12, 4), (21, 10), fill="")
            line((3.5, 10.4), (20.5, 10.4))
            for x in (6.5, 10, 14, 17.5):
                line((x, 12.2), (x, 18.6))
            line((3.5, 20.5), (20.5, 20.5))
        elif name == "code":
            line((9, 8), (4.8, 12), (9, 16))
            line((15, 8), (19.2, 12), (15, 16))
        elif name == "gear":
            oval(6.5, 6.5, 17.5, 17.5)
            oval(10.2, 10.2, 13.8, 13.8)
            for ax, ay in ((12, 3.4), (12, 20.6), (3.4, 12), (20.6, 12),
                           (5.9, 5.9), (18.1, 5.9), (5.9, 18.1), (18.1, 18.1)):
                px, py = P(ax, ay)
                qx, qy = P(12 + (ax - 12) * 0.72, 12 + (ay - 12) * 0.72)
                cv.create_line(px, py, qx, qy, fill=color, width=width,
                               capstyle="round", tags=tags)
        elif name == "sliders":
            for y, knob in ((8, 9.5), (12, 15), (16, 11)):
                line((4, y), (20, y))
                oval(knob - 2.2, y - 2.2, knob + 2.2, y + 2.2,
                     fill=theme.SURFACE, w=width)
        elif name == "bulb":
            oval(8, 5.5, 16, 14)
            line((9.6, 15.6), (10.6, 17.4), (13.4, 17.4), (14.4, 15.6))
            line((10.4, 19.6), (13.6, 19.6))
            line((12, 1.8), (12, 3.6))
            line((4.6, 6.6), (6.2, 8.2))
            line((19.4, 6.6), (17.8, 8.2))
        elif name == "shield":
            poly((12, 3), (19.5, 6), (19.5, 12), (12, 21), (4.5, 12), (4.5, 6))
            line((9, 12), (11.4, 14.4), (15.4, 9.8))
        elif name == "search":
            oval(4, 4, 15, 15)
            line((13.8, 13.8), (19.5, 19.5))
        elif name == "wrench":
            line((4.6, 19.4), (13.4, 10.6))
            oval(12.4, 3.6, 20.4, 11.6)
            poly((18.6, 5.4), (20.4, 8.4), (17.4, 9.2), fill=theme.SURFACE,
                 outline=color, w=width)
        elif name == "cloud":
            for x0, y0, x1, y1 in ((3.5, 11.5, 12.5, 20.5), (8, 9, 18, 19),
                                   (13, 12.5, 21, 20.5)):
                oval(x0, y0, x1, y1, fill=color, w=0)
            flat = P(3.5, 16)
            cv.create_rectangle(flat[0], flat[1], P(21, 20.5)[0], P(21, 20.5)[1],
                                fill=color, outline=color, tags=tags)
            line((12, 18.6), (12, 11.6)) if False else None
            a = P(12, 18.4)
            b = P(12, 11.8)
            cv.create_line(a[0], a[1], b[0], b[1], fill="#FFFFFF",
                           width=max(2, width), capstyle="round", tags=tags)
            c, d, e = P(9.6, 14.2), P(12, 11.8), P(14.4, 14.2)
            cv.create_line(c[0], c[1], d[0], d[1], e[0], e[1], fill="#FFFFFF",
                           width=max(2, width), capstyle="round",
                           joinstyle="round", tags=tags)
        elif name == "mail":
            poly((3, 5.5), (21, 5.5), (21, 18.5), (3, 18.5))
            line((3.4, 6.4), (12, 13.2), (20.6, 6.4))
        elif name == "globe":
            oval(3, 3, 21, 21)
            line((3.4, 12), (20.6, 12))
            oval(8, 3, 16, 21)
        elif name == "chevron":
            line((9.5, 6), (15.5, 12), (9.5, 18))
        elif name == "cap":
            poly((12, 4), (22, 9), (12, 14), (2, 9), fill=color, outline="", w=0)
            poly((7, 11.4), (17, 11.4), (15.4, 18.6), (8.6, 18.6),
                 fill=color, outline="", w=0)
            a, b = P(21.2, 9.6), P(21.2, 16.4)
            cv.create_line(a[0], a[1], b[0], b[1], fill=color,
                           width=max(1, int(width * 0.8)), capstyle="round",
                           tags=tags)
            oval(20, 15.6, 22.4, 18, fill=color, w=0)
        else:
            return False
        return True
    except Exception:
        return False


class BrandMark(tk.Canvas):
    """头部品牌 mark：青蓝圆角方块 + 白色学士帽剪影（规格 GLOBAL_UI_REFERENCE.png）。

    ``GLOBAL_MARK.svg`` 里是线描书本，与设计稿不一致 —— 老板拍板以设计稿的实心学士帽
    为准，配色取 theme.PRIMARY_HOVER（青蓝方块，帽体白色）。整个 mark 只用
    ``create_polygon`` / ``create_oval`` / ``create_line`` 画，零第三方依赖、零图片资源。
    """

    def __init__(self, parent, size: int = theme.MARK_SIZE, bg: str = None, **kw):
        bg = bg or parent.cget("bg")
        size = int(size)
        super().__init__(parent, width=size, height=size, bg=bg,
                         highlightthickness=0, bd=0, **kw)
        self._size = size
        self.draw()

    def draw(self):
        try:
            self.delete("all")
        except Exception:
            return
        s = self._size
        pts = rounded_points(0.5, 0.5, s - 0.5, s - 0.5, theme.MARK_RADIUS)
        self.create_polygon(pts, fill=theme.MARK_BG, outline=theme.MARK_BG,
                            width=1, tags="mark")
        k = s / 48.0
        # 学士帽（48 网格）：帽板菱形 → 帽体梯形 → 右侧帽穗
        self.create_polygon(17 * k, 22.2 * k, 31 * k, 22.2 * k,
                            29.2 * k, 29.6 * k, 18.8 * k, 29.6 * k,
                            fill=theme.MARK_FG, outline="", tags="mark")
        self.create_polygon(24 * k, 12.6 * k, 37.4 * k, 19 * k,
                            24 * k, 25.4 * k, 10.6 * k, 19 * k,
                            fill=theme.MARK_FG, outline="", tags="mark")
        self.create_line(36.6 * k, 19.8 * k, 36.6 * k, 27.2 * k,
                         fill=theme.MARK_FG, width=max(1, int(round(1.6 * k))),
                         capstyle="round", tags="mark")
        self.create_oval(34.9 * k, 26.6 * k, 38.3 * k, 30 * k,
                         fill=theme.MARK_FG, outline="", tags="mark")


class IconBadge(tk.Canvas):
    """柔和蓝小圆徽标 + 线描图标（设计稿每个字段行 / 面板标题前的装饰件）。"""

    def __init__(self, parent, icon: str, size: int = theme.BADGE_SIZE,
                 bg: str = None, fill: str = theme.PRIMARY_SOFT,
                 fg: str = theme.PRIMARY, **kw):
        bg = bg or parent.cget("bg")
        size = int(size)
        super().__init__(parent, width=size, height=size, bg=bg,
                         highlightthickness=0, bd=0, **kw)
        self._size = size
        self._icon = icon
        self._fill = fill
        self._fg = fg
        self.draw()

    def draw(self):
        try:
            self.delete("all")
        except Exception:
            return
        s = self._size
        pad = 0.5
        self.create_oval(pad, pad, s - pad, s - pad, fill=self._fill,
                         outline=self._fill, width=1)
        draw_icon(self, self._icon, s / 2.0, s / 2.0, s * 0.62, self._fg, 2)
