# -*- coding: utf-8 -*-
"""
修复桌面快捷方式图标：
  1) 浅色/高光在桌面上发灰；
  2) 图标外圈出现白边（抗锯齿边缘带着原来白色画布的浅色）。

真因：
  - 羽毛高光是半透明的（alpha≈227），桌面快捷方式把图标合成到壁纸后
    高光被背景「染灰」。
  - 圆角抗锯齿边缘像素是「原图标底色×alpha + 白色画布×(1-alpha)」，
    半透明地带着白色。放在深色壁纸上就显出白边。

修法：
  - 取不透明深色底色 bg。
  - 图标内部的半透明高光：预合成到 bg 上，并设 alpha=255。
  - 图标轮廓的抗锯齿边缘像素：把 RGB 设为 bg，保留原 alpha，
    让圆角继续平滑地融入任何背景，而不是带着白色。
  - 圆角外仍保持透明。
"""
from __future__ import annotations
import os
import shutil
import statistics
import subprocess
import sys

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
ASSETS = os.path.join(REPO, "assets")
DATA = os.path.join(REPO, "src", "data")

ICON_PNG = os.path.join(ASSETS, "icon.png")
TMP_SOURCE = os.path.join(ASSETS, "_tmp_icon_source.png")
MAKE_ICON = os.path.join(HERE, "make_icon.py")


def body_color(img: Image.Image) -> tuple[int, int, int]:
    """取不透明像素中、不太浅的像素的 RGB 中位数作为底色。"""
    px = img.load()
    W, H = img.size
    rs, gs, bs = [], [], []
    for y in range(H):
        for x in range(W):
            r, g, b, a = px[x, y]
            if a < 250:
                continue
            lum = (r + g + b) / 3.0
            if lum > 200:          # 排除高光
                continue
            rs.append(r)
            gs.append(g)
            bs.append(b)
    if not rs:
        return (90, 105, 118)      # 兜底品牌色 Global #5A6976
    return (int(statistics.median(rs)),
            int(statistics.median(gs)),
            int(statistics.median(bs)))


def _is_edge(spx, x, y, W, H) -> bool:
    """判断 (x,y) 是否处于图标轮廓边缘（8邻域内有透明像素）。"""
    for dy in (-1, 0, 1):
        ny = y + dy
        if ny < 0 or ny >= H:
            return True
        for dx in (-1, 0, 1):
            if dx == 0 and dy == 0:
                continue
            nx = x + dx
            if nx < 0 or nx >= W:
                return True
            if spx[nx, ny][3] == 0:
                return True
    return False


def make_opaque(src: Image.Image, bg: tuple[int, int, int]) -> Image.Image:
    """
    把半透明高光预合成到 bg 上（变不透明）；
    把抗锯齿轮廓边缘的 RGB 替换为 bg、保留原 alpha，避免白色画布白边；
    alpha=0 的像素保持透明。
    """
    out = Image.new("RGBA", src.size, (0, 0, 0, 0))
    spx = src.load()
    opx = out.load()
    br, bg_, bb = bg
    W, H = src.size
    for y in range(H):
        for x in range(W):
            r, g, b, a = spx[x, y]
            if a == 0:
                continue
            if _is_edge(spx, x, y, W, H):
                # 轮廓边缘：用 body 颜色代替原来带白色的抗锯齿像素，
                # 保留原 alpha 让圆角仍然平滑。
                opx[x, y] = (br, bg_, bb, a)
            else:
                # 内部像素：高光预合成到 bg 上，设为不透明。
                aa = a / 255.0
                inv = 1.0 - aa
                opx[x, y] = (
                    int(round(r * aa + br * inv)),
                    int(round(g * aa + bg_ * inv)),
                    int(round(b * aa + bb * inv)),
                    255,
                )
    return out


def main() -> int:
    src = Image.open(ICON_PNG).convert("RGBA")
    print("original size:", src.size)

    bg = body_color(src)
    print("body color (RGB):", bg)

    # 1) 预合成 -> 内部实心、圆角外透明
    solid = make_opaque(src, bg)

    # 2) 放到白色画布中央，让 make_icon.py 自动裁剪
    pad = 120
    W, H = src.size
    canvas = Image.new("RGB", (W + pad * 2, H + pad * 2), (255, 255, 255))
    canvas.paste(solid, (pad, pad), solid)
    canvas.save(TMP_SOURCE)
    print("tmp source:", TMP_SOURCE)

    # 3) 生成新的 icon 三件套
    cmd = [
        sys.executable,
        MAKE_ICON,
        TMP_SOURCE,
        "--outdir", ASSETS,
        "--name", "icon",
        "--preview", os.path.join(ASSETS, "_preview_new.png"),
    ]
    print("run:", " ".join(cmd))
    rc = subprocess.call(cmd)
    if rc != 0:
        print("make_icon.py failed")
        return rc

    # 4) 同步运行时窗口图标
    shutil.copy(os.path.join(ASSETS, "icon.png"), os.path.join(DATA, "icon.png"))
    print("copied new icon.png -> src/data/icon.png")

    # 5) 清理临时源图
    os.remove(TMP_SOURCE)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
