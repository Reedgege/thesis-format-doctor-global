# -*- coding: utf-8 -*-
"""
修复桌面快捷方式图标「白色/浅色部分变灰」问题。

真因：现有 icon.ico 中羽毛高光是半透明的（alpha≈227）。
桌面快捷方式把图标合成到桌面壁纸上，半透明高光会被壁纸颜色「染灰」；
任务栏背景深，所以看起来正常。

修法：把现有 icon.png 中的高光按它自己的不透明深色底色进行预合成，
让图标内部（圆角矩形以内）全部变成不透明，只有圆角外面保持透明。
然后用这个「实心版」重新生成 icon.png / icon.ico / icon.icns。
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


def make_opaque(src: Image.Image, bg: tuple[int, int, int]) -> Image.Image:
    """
    把 src 中 alpha>0 的像素都预合成到 bg 上，并设 alpha=255；
    alpha=0 的像素保持透明。这样图标内部完全实心，圆角外仍透明。
    """
    out = Image.new("RGBA", src.size, (0, 0, 0, 0))
    spx = src.load()
    opx = out.load()
    br, bg_, bb = bg
    for y in range(src.height):
        for x in range(src.width):
            r, g, b, a = spx[x, y]
            if a == 0:
                continue
            # 把半透明像素按 alpha 合成到 bg，结果设为不透明
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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
