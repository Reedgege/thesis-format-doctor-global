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
from .wheel import bind_wheel, wheel_pixels


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
        # v2.3.64：吸底动作栏。长表单（手填问卷 16 个字段）内容比可视区高时，主按钮
        # 会被挤到折线以下 —— 客户看不到"保存"在哪。置 True 时动作栏钉在**窗口**底部
        # （真控件，不参与滚动，见下方 v2.3.81 注释）；默认 False，其余弹窗行为不变。
        sticky_actions = bool(kw.pop("sticky_actions", False))
        super().__init__(parent, bg=bg, **kw)
        self._sticky = sticky_actions
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

        # v2.3.81：吸底动作栏升级为**真控件**，先 pack 到窗口底部。
        # 旧版它是画布里的 create_window item —— 每滚一格都要 ``coords()`` 移动它，
        # 而移动子窗口必然触发重绘，于是底部按钮区跟着一起闪（老板实测圈出来的
        # 正是这一块）。现在它完全不参与滚动：画布视口天生就短了一截，滚多少格
        # 按钮栏都纹丝不动，也就不可能有重绘。
        if sticky_actions:
            self._button_bar = tk.Frame(self, bg=theme.SURFACE)
            self._button_bar.pack(side="bottom", fill="x")
            tk.Frame(self._button_bar, bg=theme.BORDER, height=1).pack(fill="x")
            self._bar_pad = tk.Frame(self._button_bar, bg=theme.SURFACE)
            self._bar_pad.pack(fill="x", padx=24, pady=(10, 16))

        # 画圆角边框和投影的 Canvas（作为底层，可纵向滚动）
        # v2.3.80：显式 ``yscrollincrement=1`` —— 让 ``yview_scroll`` 的 "units"
        # 明确等于 1 像素。不设时（默认 0）canvas 的 1 unit = **视口高度的 1/10**
        # （实测 720 高弹窗 ≈ 69px），而旧版每格滚 ±3 units ≈ 209px（弹窗的 30%）：
        # 每次滚动都新露出一大片区域要重绘，肉眼看到的就是「闪 + 跳」。
        self._cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                             yscrollincrement=1)
        self._cv.pack(side="left", fill="both", expand=True)
        self._sb = ttk.Scrollbar(self, orient="vertical", command=self._cv.yview)
        self._cv.configure(yscrollcommand=self._on_scroll)
        self._cv.bind("<Configure>", self._draw)
        # 滚轮挂在顶层窗：子控件（Entry/Button/Label）的 MouseWheel 会冒泡到顶层窗
        # bindtag，折叠区/内容上方滚轮也能滚；return "break" 阻断继续冒泡到 root 的
        # bind_all，避免弹窗打开时主窗口在背后跟着滚。
        bind_wheel(self, self._on_wheel)

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

        # 按钮区：吸底的在 __init__ 开头已钉到窗口底部（完全不参与滚动）；
        # 非吸底的仍随内容排在末尾，需要时调用 add_primary_action / add_secondary_action。
        if not sticky_actions:
            self._button_bar = tk.Frame(self._inner, bg=theme.SURFACE)
            self._button_bar.pack(fill="x", padx=24, pady=(0, 24))
            self._bar_pad = self._button_bar

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

        # v2.3.80：先清掉上一批"卡片+投影"多边形再画。旧版每次 <Configure>
        # 都 create_polygon 一遍却从不删旧的 —— 折叠展开 / 滚动条显隐 / DPI 变化
        # 都会让画布层层叠上几十上百个一模一样的多边形，之后每次重绘都要把它们
        # 全画一遍：越用越慢、滚动越闪。只删带 "modal" 标签的（按钮栏等不受影响）。
        try:
            cv.delete("modal")
        except Exception:
            pass

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

        # v2.3.81：吸底动作栏搬到窗口底部后，画布下沿不再是窗口下沿 —— 卡的
        # 下圆角会露在动作栏正上方，两角各留一个薄荷色的小豁口（卡内纯白、底
        # 下也是纯白，中间却凹进去一块，看着像画坏了）。这里把底部 r+2 像素补
        # 成实心白条、盖掉下边框，卡片就与动作栏白色连成一片，只剩动作栏自己
        # 那条分隔线。只在吸底模式下补（其余弹窗卡片本来就到底，不需要）。
        if self._sticky:
            cv.create_rectangle(0, h - (r + 2), w, h,
                                fill=theme.SURFACE, outline=theme.SURFACE,
                                tags="modal")
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
            # v2.3.81：吸底动作栏已是窗口底部的**真控件**（不进球、不滚动），
            # 旧版这里"按可视区宽度摆 item + 补占位高度 + 每次贴底"的一套全部作废。
            ih = int(self._inner.winfo_reqheight())
            inner_w = max(1, w - 2)
            # 至少填满可视区，避免内容短时下方留空；内容更高则按自然高度
            inner_h = max(ih, h - 2)
            self._cv.itemconfigure(self._win, width=inner_w, height=inner_h)
            self._cv.configure(scrollregion=(0, 0, inner_w, inner_h))
            # 显隐用**两个不同阈值**（滞回带）：滚动条是 pack 进窗口右侧的，它一出现
            # Canvas 就变窄 → 里面自动换行的文字重新折行 → 内容高度跟着变 → 又踩回
            # 阈值另一侧 → 滚动条再消失……卡在临界点上会一直来回抽（用户看到的是
            # "内容一直在抖/漂移"）。留 6px 的回差，一旦定了就咬住不动。
            if ih > h + 2:
                if not self._sb.winfo_ismapped():
                    self._sb.pack(side="right", fill="y")
            elif ih < h - 6:
                if self._sb.winfo_ismapped():
                    self._sb.pack_forget()
            # 落在 [h-6, h+2] 回差带里：维持现状，不动作
        except Exception:
            pass
        finally:
            self._relaying = False

    def _on_scroll(self, first: str, last: str):
        """Canvas 的 yscrollcommand 回调：只同步滚动条位置。

        显隐的唯一权威是 ``_sync_scroll``（它才知道内容自然高度 vs 可视高度）。
        旧版这里也自己 pack/pack_forget 一遍滚动条 —— 两处各判一次、判据还不一样
        （这边看 yview 比例，那边看像素高），必然互相拆台：内容刚好卡在临界高度时
        滚动条一会儿出来一会儿藏，Canvas 宽度跟着变，整个弹窗内容左右晃。

        v2.3.81：这里**不再**碰吸底动作栏。它已经是窗口底部的真控件，滚动与本方法
        完全无关；旧版每次滚轮都 ``coords()`` 挪一次 canvas item，等于每格移动一个
        子窗口 → 必然重绘 → 底部按钮区跟着闪。
        """
        try:
            self._sb.set(first, last)
        except Exception:
            pass

    def _on_wheel(self, event):
        """鼠标滚轮滚动内容（按**像素**滚，鼠标一格约 120px）。

        v2.3.80：画布已设 ``yscrollincrement=1`` → 1 unit == 1px。
        旧版 ``int(-delta/120)`` 只取 ±3 个 **unit**；在没设 increment 的
        canvas 上 1 unit = 视口 1/10（实测 ≈69px），于是每格跳 ≈209px（视口的
        30%）—— 跳得猛、每次新露出大片区域要重绘，看着就是「闪」。
        现在直接按 delta 滚像素：鼠标一格 ±120px（≈视口 1/5，与主窗口手感一致），
        触控板的小 delta（±30/±60）逐像素平滑。X11 少数环境只给 ±1（一格），
        归一成一格 120px，免得「滚不动」。
        """
        # 🔴 2026-11-17：X11 的 Button-4/5 没有 delta，统一由 wheel_pixels 归一
        try:
            px = wheel_pixels(event)
            if px:
                self._cv.yview_scroll(px, "units")
        except Exception:
            pass
        return "break"

    # ---------------------------------------------------------------- 对外
    def refresh_layout(self):
        """内容改动后强制重算滚动区与滚动条（不依赖 ``<Configure>`` 是否派发）。

        调用方往 ``content`` 里加完内容、或改动了内容高度之后调用它，保证
        ``scrollregion`` 等于内容此刻的自然高度 —— 否则会出现「内容被裁且滚不到」
        （小屏幕上内容超出可视区时尤其明显）。幂等、开销很小。
        """
        try:
            self.update_idletasks()
            self._sync_scroll(self._cv.winfo_width(), self._cv.winfo_height())
        except Exception:
            pass

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
            self._bar_pad, text=text, command=command,
            style="primary", font=F["F_BTN"], height=52
        )
        btn.pack(side="right", padx=(0, 0))
        self._primary_btn = btn

    def add_secondary_action(self, text: str, command=None):
        """添加次要操作按钮（视觉安静）。"""
        F = i18n.resolve_font_spec(i18n.detect_lang())
        btn = widgets.RoundButton(
            self._bar_pad, text=text, command=command,
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
            self._bar_pad, text=text, command=command,
            style="ghost", font=F["F_BTN_S"], height=52
        )
        btn.pack(side="right", padx=(12, 0))
        self._secondary_btns.append(btn)
