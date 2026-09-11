# -*- coding: utf-8 -*-
"""从「圆角方形 App 图标」源图生成一套规范图标（PNG / ICO / ICNS）。

源图长这样：白色画布 + 居中的圆角方形插画（圆角外面是白底，还常带一圈多余的
白边）。本脚本一次做完三件事：

  1) 自动找到圆角方形的外接框（与画布底色不同的区域），把多余白边裁掉；
  2) 量出源图的圆角半径，在目标尺寸上用超采样重画一个抗锯齿的圆角矩形 alpha
     蒙版 —— 于是圆角外面变成真透明，而不是白角；
  3) 输出 1024×1024 PNG（Tk 窗口图标 / 网页 / macOS 派生 icns 的素材）、多尺寸 ICO
     （Windows exe 图标，最高 256）、以及 ICNS（macOS，纯 Python 手写容器，
     不依赖 macOS 的 iconutil）。

用法：
  python tools/make_icon.py 源图.png --outdir assets --name icon
  python tools/make_icon.py 源图.png --outdir assets --preview 预览.png

依赖：Pillow。
"""

from __future__ import annotations

import argparse
import io
import os
import struct
import sys

from PIL import Image, ImageChops, ImageDraw

# 各平台尺寸表。
# PNG 出 1024：macOS 的 .icns 需要 1024 素材（CI 用 Pillow 从 icon.png 派生 icns，
# 只给 512 的话会被放大插值，图标发糊），Tk 的 iconphoto 也能直接吃这个尺寸。
PNG_SIZE = 1024
# Windows exe 图标最大只认 256，再大是浪费体积。
ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]
# ICNS 分块类型 -> 边长（ICNS 允许直接把 PNG 塞进分块）
ICNS_CHUNKS = [
    ("icp4", 16), ("icp5", 32), ("icp6", 64),
    ("ic07", 128), ("ic08", 256), ("ic09", 512), ("ic10", 1024),
]
SUPERSAMPLE = 4
DEFAULT_RADIUS_FRAC = 0.2237          # 兜底：iOS 图标圆角占比


# --------------------------------------------------------------------- 探测
def _canvas_color(img: Image.Image) -> tuple:
    """画布底色 = 四角像素的中位数（源图四角一定是画布，不会是插画）。"""
    w, h = img.size
    px = img.load()
    corners = [px[0, 0], px[w - 1, 0], px[0, h - 1], px[w - 1, h - 1]]
    return tuple(sorted(c[i] for c in corners)[len(corners) // 2] for i in range(3))


def _content_mask(img: Image.Image, tol: int = 14) -> Image.Image:
    """与画布底色差异 > tol 的像素 -> 255。用于找圆角方形轮廓。"""
    bg = Image.new("RGB", img.size, _canvas_color(img))
    diff = ImageChops.difference(img, bg)
    r, g, b = diff.split()
    peak = ImageChops.lighter(ImageChops.lighter(r, g), b)
    return peak.point(lambda v: 255 if v > tol else 0)


def _inset_at(mask: Image.Image, box: tuple, k: int):
    """在距外接框顶部第 k 行，从左边界往里数到第一个实心像素的距离。"""
    x0, y0, x1, y1 = box
    row = y0 + k
    if row >= y1:
        return None
    data = mask.load()
    for x in range(x0, x1):
        if data[x, row]:
            return x - x0
    return None


def _measure_radius(mask: Image.Image, box: tuple) -> int:
    """量圆角半径：用「圆角 = 半径 r 的 1/4 圆」模型拟合多行内缩量。

    只读最顶一行会被抗锯齿像素骗大（顶部往上 1px，内缩量就多 sqrt(2r) 那么多），
    所以这里在第 0/1/2/3/4/6/8/12/16 行各量一次，再对 r 做最小二乘拟合。

    圆角圆心在 (r, r)，位于第 k 行的边界内缩量 = r - sqrt(r^2 - (r-k)^2)。
    """
    side = box[2] - box[0]
    samples = []
    for k in (0, 1, 2, 3, 4, 6, 8, 12, 16):
        inset = _inset_at(mask, box, k)
        if inset is not None:
            samples.append((k, inset))
    if not samples:
        return 0
    best_r, best_err = 0, None
    for r in range(4, max(5, side // 2)):
        err = 0.0
        for k, inset in samples:
            if k >= r:
                pred = 0.0
            else:
                pred = r - (r * r - (r - k) ** 2) ** 0.5
            err += (pred - inset) ** 2
        if best_err is None or err < best_err:
            best_r, best_err = r, err
    return best_r


def _square_box(box: tuple, size: tuple) -> tuple:
    """把外接框补成正方形（以中心为准），免得输出被拉变形。"""
    x0, y0, x1, y1 = box
    bw, bh = x1 - x0, y1 - y0
    side = max(bw, bh)
    cx = (x0 + x1) / 2.0
    cy = (y0 + y1) / 2.0
    nx0 = int(round(cx - side / 2.0))
    ny0 = int(round(cy - side / 2.0))
    nx1 = nx0 + side
    ny1 = ny0 + side
    W, H = size
    # 越界时整体平移回图内，而不是裁掉一边（保持正方形）
    if nx0 < 0:
        nx1 -= nx0
        nx0 = 0
    if ny0 < 0:
        ny1 -= ny0
        ny0 = 0
    if nx1 > W:
        nx0 -= (nx1 - W)
        nx1 = W
    if ny1 > H:
        ny0 -= (ny1 - H)
        ny1 = H
    return (max(0, nx0), max(0, ny0), min(W, nx1), min(H, ny1))


def _rounded_alpha(size: int, radius: float, ss: int = SUPERSAMPLE) -> Image.Image:
    """抗锯齿圆角矩形 alpha：先放大 ss 倍画硬边，再缩回来当抗锯齿。"""
    big = size * ss
    r = max(1, int(round(radius * ss)))
    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, big - 1, big - 1], radius=r, fill=255)
    return mask.resize((size, size), Image.LANCZOS)


# --------------------------------------------------------------------- 主流程
def build_crop(src: str) -> tuple:
    """返回 (裁掉白边后的方图, 圆角半径占边长比例)。"""
    img = Image.open(src).convert("RGB")
    mask = _content_mask(img)
    box = mask.getbbox()
    if not box:
        raise SystemExit("源图 %s 里没找到任何内容（整张都是纯色）。" % src)
    radius_src = _measure_radius(mask, box)
    box = _square_box(box, img.size)
    crop = img.crop(box)
    radius_frac = radius_src / float(max(1, box[2] - box[0]))
    if not (0.0 < radius_frac < 0.5):
        radius_frac = DEFAULT_RADIUS_FRAC
    return crop, radius_frac


def render(crop: Image.Image, radius_frac: float, size: int) -> Image.Image:
    """把裁好的源图缩到 size 并套上圆角透明蒙版。"""
    rgb = crop.resize((size, size), Image.LANCZOS).convert("RGBA")
    rgb.putalpha(_rounded_alpha(size, radius_frac * size))
    return rgb


def write_icns(pngs: dict, out_path: str) -> None:
    """纯 Python 写 ICNS：'icns' 头 + 若干「类型(4B) + 长度(4B) + PNG 数据」分块。"""
    body = b""
    for ctype, size in ICNS_CHUNKS:
        img = pngs.get(size)
        if img is None:
            continue
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        data = buf.getvalue()
        body += ctype.encode("ascii") + struct.pack(">I", len(data) + 8) + data
    with open(out_path, "wb") as fh:
        fh.write(b"icns" + struct.pack(">I", len(body) + 8) + body)


def make_preview(crop: Image.Image, radius_frac: float, path: str,
                 sizes=(256, 128, 64, 48, 32, 16)) -> None:
    """把各尺寸图标摆到灰白棋盘格上，用来肉眼验收圆角与透明度。"""
    pad = 16
    cell = 16
    w = pad + sum(s + pad for s in sizes)
    h = 256 + pad * 2
    canvas = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    d = ImageDraw.Draw(canvas)
    for y in range(0, h, cell):
        for x in range(0, w, cell):
            if (x // cell + y // cell) % 2 == 0:
                d.rectangle([x, y, x + cell - 1, y + cell - 1], fill=(228, 228, 228, 255))
    x = pad
    for s in sizes:
        canvas.alpha_composite(render(crop, radius_frac, s), (x, pad + (256 - s) // 2))
        x += s + pad
    canvas.convert("RGB").save(path)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="圆角方形 App 图标 -> PNG/ICO/ICNS")
    ap.add_argument("src", help="源 PNG（白底圆角方形插画）")
    ap.add_argument("--outdir", default=".", help="输出目录")
    ap.add_argument("--name", default="icon", help="输出文件名（不含扩展名）")
    ap.add_argument("--preview", help="额外输出一张验收预览图（棋盘格 + 多尺寸）")
    args = ap.parse_args(argv)

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass

    crop, radius_frac = build_crop(args.src)
    master = render(crop, radius_frac, PNG_SIZE)

    os.makedirs(args.outdir, exist_ok=True)
    png_path = os.path.join(args.outdir, args.name + ".png")
    ico_path = os.path.join(args.outdir, args.name + ".ico")
    icns_path = os.path.join(args.outdir, args.name + ".icns")

    master.save(png_path)
    master.save(ico_path, format="ICO", sizes=[(s, s) for s in ICO_SIZES])
    write_icns({s: render(crop, radius_frac, s) for _, s in ICNS_CHUNKS}, icns_path)

    if args.preview:
        make_preview(crop, radius_frac, args.preview)

    with Image.open(args.src) as orig:
        print("source        : %s (%dx%d)" % (args.src, orig.size[0], orig.size[1]))
    print("corner radius : %.1f%% of edge" % (radius_frac * 100))
    print("png  -> %s" % png_path)
    print("ico  -> %s  (%s)" % (ico_path, ", ".join(str(s) for s in ICO_SIZES)))
    print("icns -> %s  (%s)" % (icns_path, ", ".join(str(s) for _, s in ICNS_CHUNKS)))
    if args.preview:
        print("preview -> %s" % args.preview)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
