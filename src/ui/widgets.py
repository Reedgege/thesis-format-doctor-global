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
        if w <= 2 or h <= 2:
            return
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
                 bg: str = None):
        bg = bg or parent.cget("bg")
        super().__init__(parent, bg=bg, highlightthickness=0, bd=0)
        self._style = style if style in self.STYLES else "primary"
        self._radius = int(radius)
        self._h = int(height)
        self._padx = int(padx)
        self._font = font
        self._text = text
        self._command = command
        self._enabled = True
        self._hover = False
        self._visible = True
        self._pack_opts: dict = {"fill": "x"}

        self._cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                             height=self._h, width=1, takefocus=1)
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
            cv.create_text(w / 2.0, h / 2.0, text=self._text, fill=fg,
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
            self.draw()

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
# 步进圆点
# ---------------------------------------------------------------------------
class StepCircle(tk.Canvas):
    """圆形步骤标记：``todo`` 空心浅蓝 / ``active`` 实心深蓝 / ``done`` 实心绿勾。"""

    SIZE = 34

    def __init__(self, parent, number: int, font=None, size: int = SIZE, **kw):
        bg = kw.pop("bg", None) or parent.cget("bg")
        super().__init__(parent, width=size, height=size, bg=bg,
                         highlightthickness=0, bd=0, **kw)
        self._n = int(number)
        self._size = int(size)
        self._font = font
        self._state = "todo"
        self.draw()

    def set_state(self, state: str):
        if state not in ("todo", "active", "done"):
            state = "todo"
        if state == self._state:
            return
        self._state = state
        self.draw()

    @property
    def state(self) -> str:
        return self._state

    def draw(self):
        try:
            self.delete("all")
        except Exception:
            return
        s = self._size
        pad = 1.5
        if self._state == "todo":
            fill, outline, fg, text = theme.SURFACE, theme.BORDER, theme.TEXT_3, str(self._n)
        elif self._state == "active":
            fill, outline, fg, text = theme.PRIMARY, theme.PRIMARY, "#FFFFFF", str(self._n)
        else:
            fill, outline, fg, text = theme.SUCCESS, theme.SUCCESS, "#FFFFFF", "✓"
        try:
            self.create_oval(pad, pad, s - pad, s - pad, fill=fill, outline=outline,
                             width=1)
            self.create_text(s / 2.0, s / 2.0 + 0.5, text=text, fill=fg,
                             font=self._font)
        except Exception:
            pass
