# -*- coding: utf-8 -*-
"""
彻底修复海外版图标外圈白环（Windows 安装版桌面快捷方式）。

真因（v2.3.46 真机取证）：
  `make_icon.py` 的处理链是「源图 convert("RGB") 丢掉 alpha -> 按白底找内容 ->
   裁剪 -> resize -> putalpha(圆角蒙版)」。因为中途丢掉了 alpha，圆角外那圈
   **白色画布像素被原样带进了 RGB**，最后再由 putalpha 赋予半透明 alpha。
   结果就是：圆边缘 20px 宽的一圈是「白色 RGB + 不透明 alpha」，
   实测 1024 源图对角线 (144,144) = (255,255,255,255) 完全不透明。
   桌面快捷方式把图标合成到蓝色壁纸上，就成了那圈刺眼白边。

修法（不重裁、不走 make_icon.py，避免再次混入白底）：
  1. 取「不透明主体色」作为 body 色（取中位数，天然躲开白色羽毛高光）。
  2. 对整张图做「向心式重着色」：把每个非透明像素的 RGB 朝着 body 色收缩，
     收缩量 = 该像素与 body 色的差异度 × alpha/255。这样：
       - 边缘半透明像素（alpha 小）几乎完全变成 body 色 -> 白环消失；
       - 内部不透明像素（alpha=255）保持原样 -> 白色羽毛高光、设计细节不丢。
  3. alpha 通道完全不动，圆形轮廓与抗锯齿保持原样。
  4. icon.ico 用 BMP 编码（32bpp BGRA + AND 掩码）手写生成，与原 v2.3.43 一致，
     保证 Windows LoadImageW / 资源管理器都能正常读取。
"""
from __future__ import annotations
import io
import os
import shutil
import statistics
import struct

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
ASSETS = os.path.join(REPO, "assets")
DATA = os.path.join(REPO, "src", "data")

ICON_PNG = os.path.join(ASSETS, "icon.png")
ICON_ICO = os.path.join(ASSETS, "icon.ico")
ICON_ICNS = os.path.join(ASSETS, "icon.icns")
ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]
ICNS_CHUNKS = [("icp4", 16), ("icp5", 32), ("icp6", 64),
               ("ic07", 128), ("ic08", 256), ("ic09", 512), ("ic10", 1024)]


def body_color(img: Image.Image) -> tuple[int, int, int]:
    """取不透明、非浅色的像素 RGB 中位数作为主体色（躲开白色羽毛高光）。"""
    px = img.load()
    W, H = img.size
    rs, gs, bs = [], [], []
    for y in range(0, H, 2):
        for x in range(0, W, 2):
            r, g, b, a = px[x, y]
            if a < 250:
                continue
            if (r + g + b) / 3.0 > 200:
                continue
            rs.append(r); gs.append(g); bs.append(b)
    if not rs:
        return (2, 102, 92)
    return (int(statistics.median(rs)), int(statistics.median(gs)),
            int(statistics.median(bs)))


def recolor(img: Image.Image, bg: tuple[int, int, int], band: int = 30) -> Image.Image:
    """把「外边界 band 像素内、且明显亮于 body 色」的像素重着色为 body 色。

    为什么不能只看 alpha：``make_icon.py`` 丢掉 alpha 后，圆形边缘那圈白色画布
    像素被焊成了 **完全不透明**（实测 alpha=255），所以按 alpha 收缩无效。
    真正的判据是「离圆形轮廓有多近」——圆形内部的主体色本来就是深青，
    只有紧贴轮廓的一圈才是被污染的白色画布。白色羽毛高光位于圆心附近，
    距轮廓远 (band=30px)，不会被碰到。
    """
    out = Image.new("RGBA", img.size, (0, 0, 0, 0))
    spx, opx = img.load(), out.load()
    br, bg_, bb = bg
    W, H = img.size
    body_lum = (br + bg_ + bb) / 3.0
    thr = max(150.0, body_lum + 60.0)

    # 逐行从左到右 / 从右到左各扫一遍，记录每个像素到最近透明像素的水平距离；
    # 再用垂直距离补齐，得到近似的「到轮廓距离」。
    INF = 10 ** 9
    for y in range(H):
        row = [spx[x, y][3] for x in range(W)]
        # 左侧最近透明距离
        dist = [INF] * W
        last = -INF
        for x in range(W):
            if row[x] == 0:
                last = x
            dist[x] = x - last
        # 右侧最近透明距离
        last = INF
        for x in range(W - 1, -1, -1):
            if row[x] == 0:
                last = x
            d = last - x
            if d < dist[x]:
                dist[x] = d

        for x in range(W):
            r, g, b, a = row[x], 0, 0, 0
            r, g, b, a = spx[x, y]
            if a == 0:
                continue
            if dist[x] <= band and (r + g + b) / 3.0 > thr:
                opx[x, y] = (br, bg_, bb, a)
            else:
                opx[x, y] = (r, g, b, a)
    return out


def _dib_bgra(img: Image.Image) -> bytes:
    """32bpp BGRA + AND 掩码的 DIB（底向上），高度字段 2×size。"""
    w = h = img.size[0]
    px = img.load()
    and_row_bytes = ((w + 31) // 32) * 4
    bgra = bytearray()
    andmask = bytearray()
    for y in range(h - 1, -1, -1):
        brow = bytearray()
        arow = bytearray(and_row_bytes)
        for x in range(w):
            r, g, b, a = px[x, y]
            brow += bytes((b, g, r, a))
            if a == 0:
                arow[x // 8] |= (1 << (7 - (x % 8)))
        bgra += brow
        andmask += arow
    hdr = struct.pack("<IiiHHIIiiII", 40, w, h * 2, 1, 32, 0,
                      len(bgra) + len(andmask), 0, 0, 0, 0)
    return hdr + bytes(bgra) + bytes(andmask)


def write_ico_bmp(png: Image.Image, path: str, sizes) -> None:
    frames = [_dib_bgra(png.resize((s, s), Image.LANCZOS).convert("RGBA"))
              for s in sizes]
    out = bytearray(struct.pack("<HHH", 0, 1, len(frames)))
    off = 6 + 16 * len(frames)
    for s, dib in zip(sizes, frames):
        wb = 0 if s >= 256 else s
        out += struct.pack("<BBBBHHII", wb, wb, 0, 0, 1, 32, len(dib), off)
        off += len(dib)
    for dib in frames:
        out += dib
    with open(path, "wb") as fh:
        fh.write(out)


def write_icns(pngs: dict, out_path: str) -> None:
    body = b""
    for ctype, size in ICNS_CHUNKS:
        img = pngs.get(size)
        if img is None:
            continue
        b2 = io.BytesIO()
        img.save(b2, format="PNG")
        data = b2.getvalue()
        body += ctype.encode("ascii") + struct.pack(">I", len(data) + 8) + data
    with open(out_path, "wb") as fh:
        fh.write(b"icns" + struct.pack(">I", len(body) + 8) + body)


def _render_check(img: Image.Image) -> None:
    """蓝底 + 黑底 + 白底三张 256 预览，肉眼复核白环是否消失。"""
    small = img.resize((256, 256), Image.LANCZOS)
    for nm, bg in (("blue", (55, 115, 175)), ("black", (8, 8, 8)),
                   ("white", (245, 245, 245))):
        c = Image.new("RGB", (256, 256), bg)
        c.paste(small, (0, 0), small)
        c.save(os.path.join(ASSETS, f"_check_{nm}.png"))


def _report(img: Image.Image, tag: str) -> None:
    """量化白环：统计「外边界 4px 带内、平均亮度远高于 body 色」的像素。"""
    px = img.load()
    W, H = img.size
    bad = tot = 0
    for y in range(H):
        for x in range(W):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            tot += 1
            near = False
            for dy in range(-4, 5):
                for dx in range(-4, 5):
                    nx, ny = x + dx, y + dy
                    if nx < 0 or nx >= W or ny < 0 or ny >= H or px[nx, ny][3] == 0:
                        near = True
                        break
                if near:
                    break
            if near and (r + g + b) / 3.0 > 150:
                bad += 1
    print(f"  [{tag}] 外边界4px带内偏亮像素 = {bad} / 非透明 {tot}"
          f" ({bad * 100 // max(1, tot)}%)")


def main() -> int:
    src = Image.open(ICON_PNG).convert("RGBA")
    print("源图:", src.size)
    bg = body_color(src)
    print("主体色 body =", bg)
    _report(src, "修复前")

    fixed = recolor(src, bg)
    _report(fixed, "修复后")

    _render_check(fixed)
    fixed.save(ICON_PNG)
    write_ico_bmp(fixed, ICON_ICO, ICO_SIZES)
    write_icns({s: fixed.resize((s, s), Image.LANCZOS) for _, s in ICNS_CHUNKS},
               ICON_ICNS)
    shutil.copy(ICON_PNG, os.path.join(DATA, "icon.png"))
    print("已写出 icon.png / icon.ico(BMP) / icon.icns / src/data/icon.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
