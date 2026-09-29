"""UI 小图标 PNG 加载器。

v2.3.17 起，设计稿里字段行 / 面板标题 / 按钮前的小图标逐步从 Canvas 自绘切换为
PNG 图片资源，调用方（widgets.IconBadge / RoundButton / app.py）保持原接口不变。

本模块只负责：
1. 在源码运行和 Nuitka 打包后都能定位到 ``assets/img/icons/*.png``；
2. 用 Pillow 加载并按需缩放；
3. 缓存 ``ImageTk.PhotoImage`` 避免重复读盘。

所有失败都静默返回 ``None``，调用方回退到自绘图标，保证「图标缺失不崩界面」。
"""

from __future__ import annotations

import os
import sys
from typing import Optional


def _norm(path: str) -> str:
    try:
        return os.path.normpath(os.path.abspath(path))
    except Exception:
        return path


def _runtime_bases() -> list:
    """打包后「可执行文件所在目录」候选（与 iconpath 同款逻辑，避免交叉导入）。"""
    bases = []
    argv = getattr(sys, "argv", None) or []
    raw_list = [argv[0] if argv else "", getattr(sys, "executable", "") or ""]
    for raw in raw_list:
        try:
            if raw:
                bases.append(os.path.dirname(_norm(raw)))
        except Exception:
            continue
    try:
        bases.append(os.getcwd())
    except Exception:
        pass
    out = []
    for b in bases:
        if b and b not in out:
            out.append(b)
    return out


def icon_dir_candidates() -> list:
    """返回 ``assets/img/icons`` 的可能路径（含不存在的），按命中概率排序。"""
    here = os.path.dirname(_norm(__file__))
    cands = [
        os.path.join(here, "..", "..", "assets", "img", "icons"),  # 源码：src/ui -> repo/assets/img/icons
        os.path.join(here, "..", "data", "img", "icons"),           # 备选
    ]
    for base in _runtime_bases():
        cands += [
            os.path.join(base, "assets", "img", "icons"),           # Nuitka：exe 同级
            os.path.join(base, "img", "icons"),
            os.path.join(base, "..", "assets", "img", "icons"),     # macOS .app 内 MacOS/../Resources?
        ]
    out = []
    for c in cands:
        c = _norm(c)
        if c not in out:
            out.append(c)
    return out


def find_icon(name: str) -> "str | None":
    """返回 ``{name}.png`` 第一个真实存在的路径；没有则返回 None。"""
    for d in icon_dir_candidates():
        path = os.path.join(d, name + ".png")
        try:
            if os.path.isfile(path):
                return path
        except Exception:
            continue
    return None


# Canvas 自绘图标名 → PNG 图标名映射。
# 只列出「有对应 PNG」的图标；未列出的（如 sliders/cloud/cap/chevron/mail/globe）
# 继续走 widgets.draw_icon 自绘，保证不因为没有图片就空白。
ICON_NAME_MAP = {
    # 字段行 / 面板标题
    "doc": "your_documents",        # 左卡大标题 "Your Documents"
    "your_documents": "your_documents",
    "paper": "paper",               # Paper 字段行（v2.3.17 与标题拆成两个 PNG）
    "quote": "citation_style",      # Citation style 字段行
    "citation": "citation_style",
    "bank": "university_template",  # University template 字段行
    "university": "university_template",
    "code": "latex_template",       # LaTeX template 字段行
    "latex": "latex_template",
    "gear": "advanced_options",     # Advanced options 折叠条标题
    "sliders": "review_and_fix",    # Review & Fix 标题（用户补生成的 PNG）
    # 顶部隐私 / 右侧信任区
    "thesis_local_privacy": "thesis_local_privacy",  # 顶部 "Your thesis never leaves your computer"
    "built_academic": "built_academic",  # 右卡信任卡片 "Built for academic work"
    # Advanced options 内部新增图标（原来自绘位置无图标，现在补上）
    "questionnaire": "questionnaire",
    "import_config": "import_config",
    # 按钮
    "search": "btn_check_formatting",
    "wrench": "btn_fix_issues",
}

# 每个图标在界面上常见的目标尺寸；未命中时默认按请求 size 缩放。
ICON_DEFAULT_SIZE = {
    "your_documents": 30,
    "paper": 26,
    "citation_style": 26,
    "advanced_options": 26,
    "university_template": 26,
    "latex_template": 26,
    "questionnaire": 26,
    "import_config": 26,
    "thesis_local_privacy": 22,
    "built_academic": 26,
    "review_and_fix": 26,
    "btn_check_formatting": 18,
    "btn_fix_issues": 18,
}


# 缓存：name@size -> PhotoImage
_cache: dict = {}


def _resolve_name(name: str) -> "str | None":
    """把调用方用的自绘图标名解析成 PNG 名；None 表示保留自绘。"""
    mapped = ICON_NAME_MAP.get(name)
    if mapped is None:
        return None
    return mapped


def load_icon(name: str, size: Optional[int] = None) -> Optional[object]:
    """加载图标并返回 ``ImageTk.PhotoImage``；失败返回 None。

    ``name`` 可以是：
    - 自绘图标名（如 ``"doc"`` / ``"gear"``）——按 ICON_NAME_MAP 解析；
    - 直接 PNG 名（如 ``"your_documents"``）——只要文件存在就加载。

    ``size`` 为 None 时使用 ICON_DEFAULT_SIZE 对应值；仍未命中则按图片原始尺寸。
    """
    # 支持直接传 PNG 名（如 app.py 里写死的图片名）
    png_name = _resolve_name(name) or name
    path = find_icon(png_name)
    if not path:
        return None

    target_size = size
    if target_size is None:
        target_size = ICON_DEFAULT_SIZE.get(png_name)

    cache_key = (png_name, target_size)
    if cache_key in _cache:
        return _cache[cache_key]

    try:
        from PIL import Image, ImageTk, ImageOps
        img = Image.open(path).convert("RGBA")
        if target_size:
            # 正方形缩放，避免用户图标不是严格 1:1 时变形
            img = ImageOps.contain(img, (target_size, target_size), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        _cache[cache_key] = photo
        return photo
    except Exception:
        return None


def load_icon_for(name: str, size: int) -> Optional[object]:
    """强制按 ``size`` 加载图标（用于调用方知道自己想要什么尺寸时）。"""
    return load_icon(name, size=size)


def preload_all() -> dict:
    """预加载全部已映射图标，返回 {name: photo}。"""
    out = {}
    for drawn_name in ICON_NAME_MAP:
        png_name = ICON_NAME_MAP[drawn_name]
        if png_name:
            photo = load_icon(drawn_name)
            if photo:
                out[png_name] = photo
    return out
