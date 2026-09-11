"""可缩放的 tkinter 字体对象。

与 ``i18n.py`` 的分工：那边是**纯数据**（字体族名 + 平台兜底规则，可离线单测），
这边只负责把族名变成真正的 ``tkfont.Font`` 命名对象，并处理随窗口缩放的联动。

为什么要"命名 Font 对象"而不是到处写 ``font=("Georgia", 21, "bold")``：
tkinter 里 tuple 形式的字体在 widget 建立时就固化下来了，窗口拉大后想整体放大字号
就得遍历所有 widget 逐个重设；而 ``Font`` 命名对象是**引用传递**的，改一次
``font.configure(size=...)``，所有引用它的 widget 立即跟着更新。
"""

from __future__ import annotations

import tkinter.font as tkfont

from . import i18n

BASE_WIDTH = 1180          # 设计基准窗口宽度（px），与主窗口默认尺寸对应
MIN_SCALE = 0.80           # 缩放下限（再小就不好读了）
MAX_SCALE = 1.00           # 上限 1.0：窗口拉大不放大字体，只让留白变宽
MIN_SIZE = 9               # 字号下限（磅）


class Fonts:
    """按语言构建的字体集合；``fonts["F_TITLE"]`` 取用，切换语言时 ``rebuild``。"""

    def __init__(self, root):
        self.root = root
        self.lang = i18n.DEFAULT_LANG
        self._fonts: dict = {}      # name -> (Font, base_size)
        self._scale = 1.0
        self.rebuild(self.lang)

    # ---------------------------------------------------------------- build
    def rebuild(self, lang: str):
        """按语言重建全部字体对象（旧对象交给 GC，切换语言是低频操作）。"""
        self.lang = lang if lang in i18n.LANGS else i18n.DEFAULT_LANG
        try:
            available = set(tkfont.families(self.root))
        except Exception:
            available = None
        spec = i18n.font_spec(self.lang, available)
        new_fonts: dict = {}
        for name, (family, size, weight) in spec.items():
            try:
                f = tkfont.Font(root=self.root, family=family, size=size, weight=weight)
            except Exception:
                f = tkfont.Font(root=self.root, size=size, weight=weight)
            new_fonts[name] = (f, size)
        self._fonts = new_fonts
        self._apply_size()

    def __getitem__(self, name: str):
        try:
            return self._fonts[name][0]
        except KeyError:
            # 未知名字不该让界面崩：返回正文字体兜底
            return self._fonts["F_BODY"][0]

    def get(self, name: str, default=None):
        return self._fonts.get(name, (default,))[0]

    # ---------------------------------------------------------------- scale
    def scale_for_width(self, width: int) -> float:
        """按窗口宽度算缩放系数（钳制在 MIN_SCALE~MAX_SCALE）。"""
        if not width or width < 60:
            return self._scale
        factor = width / BASE_WIDTH
        return max(MIN_SCALE, min(MAX_SCALE, factor))

    def apply_scale(self, factor: float):
        self._scale = max(MIN_SCALE, min(MAX_SCALE, factor))
        self._apply_size()

    def _apply_size(self):
        for f, base in self._fonts.values():
            try:
                f.configure(size=max(MIN_SIZE, int(round(base * self._scale))))
            except Exception:
                pass
