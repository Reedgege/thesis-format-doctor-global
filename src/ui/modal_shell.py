"""模态弹窗组件 —— 符合 modal-shell.svg 视觉规范。

规格：
- 宽度 560-640px
- cream/white surface (SURFACE)
- 12-14px radius
- thin cool-gray border (BORDER)
- subtle shadow
- title in academic serif (Georgia)
- body in clean sans-serif (Inter/Segoe UI)
- one clear primary action
- secondary actions visually quiet
- no red title bars
- no giant colored headers
- 24-32px internal spacing
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from . import theme, widgets, i18n


class ModalShell(tk.Toplevel):
    """符合规范的模态弹窗。

    参数
    ----
    parent  : 父窗口
    title   : 弹窗标题（用 academic serif）
    width   : 宽度（默认 600，规格要求 560-640）
    height  : 高度（可选，未指定则自动适应）
    """

    def __init__(self, parent, title: str, width: int = 600, height: int = None, **kw):
        bg = kw.pop("bg", None) or theme.BG
        super().__init__(parent, bg=bg, **kw)
        self._title_text = title
        self._width = int(width)
        self._height = height
        self._parent = parent
        self._primary_btn = None
        self._secondary_btns = []

        # 弹窗背景（与页面背景一致）
        self.configure(bg=bg)

        # 居中并设置尺寸
        self._center()
        self.resizable(False, False)  # 默认不可调整大小（保持简洁）
        self.title(title)

        # 画圆角边框和投影的 Canvas（作为底层，可纵向滚动）
        self._cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self._cv.pack(side="left", fill="both", expand=True)
        self._sb = ttk.Scrollbar(self, orient="vertical", command=self._cv.yview)
        self._cv.configure(yscrollcommand=self._on_scroll)
        self._cv.bind("<Configure>", self._draw)
        # 滚轮挂在顶层窗：子控件（Entry/Button/Label）的 MouseWheel 会冒泡到顶层窗
        # bindtag，折叠区/内容上方滚轮也能滚；return "break" 阻断继续冒泡到 root 的
        # bind_all，避免弹窗打开时主窗口在背后跟着滚。
        self.bind("<MouseWheel>", self._on_wheel)

        # 内容容器（嵌入 Canvas，超出可视区时由滚动条接管）
        self._inner = tk.Frame(self._cv, bg=theme.SURFACE)
        self._win = self._cv.create_window(0, 0, window=self._inner, anchor="nw")
        self._relaying = False

        # 标题
        F = i18n.resolve_font_spec(i18n.detect_lang())
        self._title_lbl = tk.Label(
            self._inner, text=title,
            bg=theme.SURFACE, fg=theme.TEXT,
            font=F["F_DIALOG_TITLE"], anchor="w"
        )
        self._title_lbl.pack(fill="x", padx=24, pady=(24, 0))

        # 分隔线（可选）
        self._separator = tk.Frame(self._inner, bg=theme.BORDER, height=1)
        # 默认隐藏，需要时调用 show_separator()

        # 内容区
        self._content = tk.Frame(self._inner, bg=theme.SURFACE)
        self._content.pack(fill="both", expand=True, padx=24, pady=16)

        # 按钮区（默认隐藏，需要时调用 add_primary_action / add_secondary_action）
        self._button_bar = tk.Frame(self._inner, bg=theme.SURFACE)
        self._button_bar.pack(fill="x", padx=24, pady=(0, 24))

        self.update_idletasks()

    def _center(self):
        """将弹窗居中显示。"""
        self.withdraw()
        self.update_idletasks()

        # 计算居中位置
        pw = self._parent.winfo_width()
        ph = self._parent.winfo_height()
        x = self._parent.winfo_rootx() + (pw - self._width) // 2
        if self._height:
            y = self._parent.winfo_rooty() + (ph - self._height) // 2
        else:
            y = self._parent.winfo_rooty() + ph // 4

        self.geometry(f"{self._width}x{self._height or 400}+{x}+{y}")
        self.deiconify()

    def _draw(self, e=None):
        """画圆角边框和投影。"""
        cv = self._cv
        try:
            w = e.width if e is not None else cv.winfo_width()
            h = e.height if e is not None else cv.winfo_height()
        except Exception:
            return
        if w <= 2 or h <= 2:
            return

        # 14px 圆角（modal-shell 规格）
        r = 14

        # 投影（下移 3px）
        sh = widgets.rounded_points(2, 3, w - 2, h - 2 + 3, r)
        cv.create_polygon(sh, fill=theme.SHADOW, outline=theme.SHADOW,
                          tags="modal")

        # 卡片主体
        pts = widgets.rounded_points(1, 1, w - 1, h - 1, r)
        cv.create_polygon(pts, fill=theme.SURFACE, outline=theme.BORDER,
                          width=1, tags="modal")
        cv.tag_lower("modal")

        # 更新内容容器尺寸 + 滚动区域（内容超出可视区时显示滚动条）
        self._sync_scroll(w, h)

    # ------------------------------------------------------------- 滚动支持
    def _sync_scroll(self, w: int, h: int):
        """把内容容器调到真实高度，并在超出可视区时启用纵向滚动条。

        关键修复：旧实现把 ``_inner`` 强制压成可视区高度（inner_h = h-2），
        内容一旦比弹窗高就被 Canvas 裁掉、且 Canvas 没有滚动条 → 离线激活
        折叠展开后整段落在可视区外、点 [+] 像没反应。现在让 ``_inner`` 按自然
        高度排布，滚动区覆盖全部内容，溢出才显示滚动条。
        """
        if self._relaying:
            return
        self._relaying = True
        try:
            self.update_idletasks()
            ih = int(self._inner.winfo_reqheight())
            inner_w = max(1, w - 2)
            # 至少填满可视区，避免内容短时下方留空；内容更高则按自然高度
            inner_h = max(ih, h - 2)
            self._cv.itemconfigure(self._win, width=inner_w, height=inner_h)
            self._cv.configure(scrollregion=(0, 0, inner_w, inner_h))
            if ih > h + 2:
                if not self._sb.winfo_ismapped():
                    self._sb.pack(side="right", fill="y")
            else:
                if self._sb.winfo_ismapped():
                    self._sb.pack_forget()
        except Exception:
            pass
        finally:
            self._relaying = False

    def _on_scroll(self, first: str, last: str):
        """Canvas 的 yscrollcommand 回调：同步滚动条并隐藏/显示。"""
        try:
            self._sb.set(first, last)
            f, l = float(first), float(last)
            if f <= 0.0 and l >= 1.0:
                if self._sb.winfo_ismapped():
                    self._sb.pack_forget()
            else:
                if not self._sb.winfo_ismapped():
                    self._sb.pack(side="right", fill="y")
        except Exception:
            pass

    def _on_wheel(self, event):
        """鼠标滚轮滚动内容（Windows 用 delta，跨平台兜底）。"""
        try:
            self._cv.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass
        return "break"

    # ---------------------------------------------------------------- 对外
    @property
    def content(self) -> tk.Frame:
        """内容容器：调用方把正文 pack/grid 进来（标题与按钮栏由本类负责）。"""
        return self._content

    def show_separator(self):
        """显示标题下方的分隔线。"""
        self._separator.pack(fill="x", padx=24, pady=(16, 0))

    def add_primary_action(self, text: str, command=None):
        """添加主操作按钮（视觉主导）。

        规格：一个弹窗只有一个主 CTA。
        """
        if self._primary_btn is not None:
            raise RuntimeError("Modal already has a primary action")

        F = i18n.resolve_font_spec(i18n.detect_lang())
        btn = widgets.RoundButton(
            self._button_bar, text=text, command=command,
            style="primary", font=F["F_BTN"], height=52
        )
        btn.pack(side="right", padx=(0, 0))
        self._primary_btn = btn

    def add_secondary_action(self, text: str, command=None):
        """添加次要操作按钮（视觉安静）。"""
        F = i18n.resolve_font_spec(i18n.detect_lang())
        btn = widgets.RoundButton(
            self._button_bar, text=text, command=command,
            style="secondary", font=F["F_BTN_S"], height=52
        )
        btn.pack(side="right", padx=(12, 0))
        self._secondary_btns.append(btn)

    def add_cancel_action(self, text: str, command=None):
        """添加取消按钮（ghost 样式）。"""
        if command is None:
            command = self.destroy

        F = i18n.resolve_font_spec(i18n.detect_lang())
        btn = widgets.RoundButton(
            self._button_bar, text=text, command=command,
            style="ghost", font=F["F_BTN_S"], height=52
        )
        btn.pack(side="right", padx=(12, 0))
        self._secondary_btns.append(btn)
