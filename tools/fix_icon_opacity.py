# -*- coding: utf-8 -*-
"""
修复海外版 Windows 安装图标的两类问题（直接在原 icon.png 上改，不重裁）：

  1) 桌面快捷方式图标外圈白边：圆角抗锯齿边缘原本是「body×alpha + 白色画布
     ×(1-alpha)」，半透明地带着白色。改成 RGB=body 色、保留原 alpha，让圆角平滑
     融入任意背景而不带白色光晕。
  2) 任务栏图标丢失：原 v2.3.43 的 icon.ico 是 BMP 编码、任务栏正常；重生成时
     PIL 写成了 PNG 编码，LoadImageW 取不到图标导致 _set_win_icon_from_ico 兜底
     失败。这里用 BMP 编码（32bpp BGRA + AND 掩码）重新生成，与 v2.3.43 同格式。

内部半透明高光也预合成到 body 色并设为不透明，避免被桌面壁纸染灰。
"""
from __future__ import annotations
import io
import os
import shutil
import statistics
import struct

from PIL import Image, ImageDraw

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
    px = img.load()
    W, H = img.size
    rs, gs, bs = [], [], []
    for y in range(H):
        for x in range(W):
            r, g, b, a = px[x, y]
            if a < 250:
                continue
            if (r + g + b) / 3.0 > 200:     # 排除高光
                continue
            rs.append(r); gs.append(g); bs.append(b)
    if not rs:
        return (2, 102, 92)
    return (int(statistics.median(rs)), int(statistics.median(gs)),
            int(statistics.median(bs)))


def _is_edge(px, x, y, W, H) -> bool:
    for dy in (-1, 0, 1):
        ny = y + dy
        if ny < 0 or ny >= H:
            return True
        for dx in (-1, 0, 1):
            if dx == 0 and dy == 0:
                continue
            nx = x + dx
            if nx < 0 or nx >= W or px[nx, ny][3] == 0:
                return True
    return False


def fix_icon(src: Image.Image, bg: tuple[int, int, int]) -> Image.Image:
    """直接在原图上修正：轮廓边缘 RGB=body、保留 alpha；内部半透明高光合成到 body
    并设不透明；alpha=0 保持透明。不重裁，保留原透明圆角。"""
    out = Image.new("RGBA", src.size, (0, 0, 0, 0))
    spx, opx = src.load(), out.load()
    br, bg_, bb = bg
    W, H = src.size
    for y in range(H):
        for x in range(W):
            r, g, b, a = spx[x, y]
            if a == 0:
                continue
            if _is_edge(spx, x, y, W, H):
                opx[x, y] = (br, bg_, bb, a)          # 平滑深色边缘
            elif a < 250:                              # 内部半透明高光
                aa = a / 255.0; inv = 1.0 - aa
                opx[x, y] = (int(round(r * aa + br * inv)),
                             int(round(g * aa + bg_ * inv)),
                             int(round(b * aa + bb * inv)), 255)
            else:
                opx[x, y] = (r, g, b, 255)
    return out


def _dib_bgra(img: Image.Image) -> bytes:
    """32bpp BGRA + AND 掩码的 DIB（底向上），高度字段为 2×size。"""
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
    """手写 BMP 编码 ICO（与 v2.3.43 同格式，任务栏 LoadImageW 可识别）。"""
    frames = []
    for s in sizes:
        frames.append(_dib_bgra(png.resize((s, s), Image.LANCZOS).convert("RGBA")))
    out = bytearray()
    out += struct.pack("<HHH", 0, 1, len(frames))
    off = 6 + 16 * len(frames)
    for s, dib in zip(sizes, frames):
        wb = 0 if s >= 256 else s          # ICO 目录里 0 表示 256
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


def _preview(png: Image.Image, out_path: str) -> None:
    """棋盘格 + 蓝底两张预览，便于人工复核。"""
    small = png.resize((256, 256), Image.LANCZOS)
    # 棋盘格
    chk = Image.new("RGBA", (256, 256), (255, 255, 255, 255))
    d = ImageDraw.Draw(chk)
    for y in range(0, 256, 16):
        for x in range(0, 256, 16):
            if ((x // 16 + y // 16) % 2) == 0:
                d.rectangle([x, y, x + 15, y + 15], fill=(214, 214, 214, 255))
    chk.alpha_composite(small, (0, 0))
    chk.convert("RGB").save(os.path.join(ASSETS, "_preview_fixed_checker.png"))
    # 蓝底
    blue = Image.new("RGB", (256, 256), (55, 115, 175))
    blue.paste(small, (0, 0), small)
    blue.save(os.path.join(ASSETS, "_preview_fixed_blue.png"))


def main() -> int:
    src = Image.open(ICON_PNG).convert("RGBA")
    print("src size:", src.size)
    bg = body_color(src)
    print("body color:", bg)

    fixed = fix_icon(src, bg)
    _preview(fixed, os.path.join(ASSETS, "_preview_fixed.png"))

    fixed.save(ICON_PNG)
    write_ico_bmp(fixed, ICON_ICO, ICO_SIZES)
    print("ico ->", ICON_ICO)
    write_icns({s: fixed.resize((s, s), Image.LANCZOS) for _, s in ICNS_CHUNKS},
               ICON_ICNS)
    print("icns ->", ICON_ICNS)
    shutil.copy(ICON_PNG, os.path.join(DATA, "icon.png"))
    print("copied -> src/data/icon.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
