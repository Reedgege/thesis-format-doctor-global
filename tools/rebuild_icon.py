# -*- coding: utf-8 -*-
"""
从原始设计源图重建海外版图标（不再补任何白底）。

设计事实（源图实测确认，老板诊断正确）：
  - 源图 `03_圆形徽章logo_circular_badge.png` 里，白色羽毛是**镂空**的：
    圆心白色像素平均 alpha 仅 6.5，98% 半透明/全透明。
  - 源图本身 alpha 干净：四角 alpha=0，圆形边缘是
    「透明 -> 青色」直接过渡，**中间没有白色过渡带**。
  - 历次白环的真正来源：拿到源图后用 make_icon.py 走
    「convert("RGB") 丢 alpha -> 按白底找内容 -> putalpha」的链路，
    白色画布被焊进圆形边缘；再加上给半透明羽毛补不透明白底的处理，
    最终在轮廓外堆出一圈不透明白环。

本脚本（按老板要求：白羽镂空保留、白底不加、范围严格小于图标）：
  1. 源图 alpha 已是干净的「实心圆 + 镂空羽」，直接沿用；
  2. 只做「补间」：把抗锯齿边缘的 RGB 归一到主体色（避免半透明像素
     带着源图的白/浅色在桌面上显出白边），alpha 一律不动；
  3. 圆形之外保持 alpha=0，**任何位置都不补不透明白底**；
  4. 输出 icon.png（1024）/ icon.ico（BMP 编码 16~256）/ icon.icns。
"""
from __future__ import annotations
import io
import os
import shutil
import statistics
import struct
from collections import deque

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
ASSETS = os.path.join(REPO, "assets")
DATA = os.path.join(REPO, "src", "data")

# 老板提供的原始设计源图（白羽镂空、alpha 干净）
SRC = r"D:\AgentSpace\素材中心\UI设计\03_圆形徽章logo_circular_badge.png"
if not os.path.isfile(SRC):
    SRC = os.path.join(REPO, "assets", "_base_icon.png")

ICON_PNG = os.path.join(ASSETS, "icon.png")
ICON_ICO = os.path.join(ASSETS, "icon.ico")
ICON_ICNS = os.path.join(ASSETS, "icon.icns")
ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]
ICNS_CHUNKS = [("icp4", 16), ("icp5", 32), ("icp6", 64),
               ("ic07", 128), ("ic08", 256), ("ic09", 512), ("ic10", 1024)]
PNG_SIZE = 1024


def body_color(img: Image.Image) -> tuple[int, int, int]:
    """主体色 = 不透明、非浅色像素的 RGB 中位数（避开白色羽毛）。"""
    px = img.load()
    W, H = img.size
    rs, gs, bs = [], [], []
    for y in range(0, H, 2):
        for x in range(0, W, 2):
            r, g, b, a = px[x, y]
            if a < 250:
                continue
            if (r + g + b) / 3.0 > 190:
                continue
            rs.append(r); gs.append(g); bs.append(b)
    if not rs:
        return (2, 102, 92)
    return (int(statistics.median(rs)), int(statistics.median(gs)),
            int(statistics.median(bs)))


def remove_outer_halo(img: Image.Image, bg: tuple[int, int, int],
                      band: int = 20, alpha_thr: int = 20,
                      light_thr: int = 110) -> Image.Image:
    """只修复外圈白/灰光晕，不误伤内部羽毛镂空。

    步骤：
      1. 从四角 flood-fill 出「外部透明区」（alpha < alpha_thr）。
      2. BFS 计算每个像素到外部透明区的距离。
      3. 在 band 距离内、颜色偏浅（> light_thr）的像素 RGB 刷成主体色。

    这样把源图/缩放后外圈残留的白/灰 RGB 全部改为主体色，合成到深色背景上
    就不会再显白/灰边；内部羽毛镂空离外圈远，不会被误刷。alpha 一律保持原样，
    形状不变。
    """
    px = img.load()
    W, H = img.size
    br, bg_, bb = bg

    # 1) 外部透明区 mask
    out = [[False] * W for _ in range(H)]
    q = deque()
    for y in range(H):
        for x in (0, W - 1):
            if px[x, y][3] < alpha_thr and not out[y][x]:
                out[y][x] = True
                q.append((x, y))
    for x in range(W):
        for y in (0, H - 1):
            if px[x, y][3] < alpha_thr and not out[y][x]:
                out[y][x] = True
                q.append((x, y))
    while q:
        x, y = q.popleft()
        for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            nx, ny = x + dx, y + dy
            if (0 <= nx < W and 0 <= ny < H and not out[ny][nx]
                    and px[nx, ny][3] < alpha_thr):
                out[ny][nx] = True
                q.append((nx, ny))

    # 2) 到外部透明区的 BFS 距离
    INF = 10 ** 9
    dist = [[INF] * W for _ in range(H)]
    q2 = deque()
    for y in range(H):
        for x in range(W):
            if out[y][x]:
                dist[y][x] = 0
                q2.append((x, y))
    while q2:
        x, y = q2.popleft()
        d = dist[y][x]
        for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            nx, ny = x + dx, y + dy
            if 0 <= nx < W and 0 <= ny < H and dist[ny][nx] > d + 1:
                dist[ny][nx] = d + 1
                q2.append((nx, ny))

    # 3) 把 band 内的浅色像素刷成主体色
    fixed = img.copy()
    px2 = fixed.load()
    for y in range(H):
        for x in range(W):
            r, g, b, a = px2[x, y]
            if a == 0:
                continue
            if dist[y][x] <= band and (r + g + b) / 3.0 > light_thr:
                px2[x, y] = (br, bg_, bb, a)
    return fixed


def alpha_flatten(img: Image.Image, bg: tuple[int, int, int]) -> Image.Image:
    """把所有透明像素的 RGB 压平成主体色（alpha 不变），避免缩放时把透明区
    里可能残留的浅色 RGB 渗进半透明边缘。"""
    out = img.copy()
    px = out.load()
    br, bg_, bb = bg
    W, H = img.size
    for y in range(H):
        for x in range(W):
            r, g, b, a = px[x, y]
            if a == 0:
                px[x, y] = (br, bg_, bb, 0)
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


def _stats(img: Image.Image, tag: str):
    px = img.load()
    W, H = img.size
    # 白环指标：外边界 4px 带内的不透明浅色像素
    bad = 0
    for y in range(H):
        for x in range(W):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            near = False
            for dy in range(-4, 5):
                for dx in range(-4, 5):
                    nx, ny = x + dx, y + dy
                    if nx < 0 or nx >= W or ny < 0 or ny >= H or px[nx, ny][3] == 0:
                        near = True
                        break
                if near:
                    break
            if near and a > 200 and (r + g + b) / 3.0 > 150:
                bad += 1
    # 镂空校验：圆心白色像素的平均 alpha（应远小于 255）
    cx, cy = W // 2, H // 2
    vals = [px[x, y][3] for y in range(cy - H // 7, cy + H // 7)
            for x in range(cx - W // 7, cx + W // 7)
            if (px[x, y][0] + px[x, y][1] + px[x, y][2]) / 3 > 200]
    avg = sum(vals) / len(vals) if vals else -1
    print(f"  [{tag}] 边界不透明白像素={bad}   羽毛区平均 alpha={avg:.1f}"
          f"（<255 = 镂空保留）")


def main() -> int:
    src = Image.open(SRC).convert("RGBA")
    print("源图:", SRC)
    print("源图尺寸:", src.size)
    # 主体色从源图取（圆形本身的颜色）
    bg = body_color(src)
    print("主体色 =", bg)

    # 第一遍：源图上去除外圈白/灰光晕（只对外圈透明区距离内的浅色像素动手）
    fixed = remove_outer_halo(src, bg, band=20, alpha_thr=20, light_thr=110)
    # alpha 压平：所有透明像素的 RGB 改主体色，避免 LANCZOS 缩放时把透明区里的
    # 浅色渗进半透明边缘，再次形成白/灰边。
    fixed = alpha_flatten(fixed, bg)

    # 源图非正方形（220x240），先补成正方形画布再缩放，避免变形
    W, H = src.size
    side = max(W, H)
    sq_fixed = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    sq_fixed.paste(fixed, ((side - W) // 2, (side - H) // 2), fixed)
    # alpha 通道用**原图**的，保证形状和镂空羽毛与源图完全一致
    sq_alpha = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    sq_alpha.paste(src, ((side - W) // 2, (side - H) // 2), src)

    base_rgb = sq_fixed.convert("RGB").resize((PNG_SIZE, PNG_SIZE),
                                               Image.LANCZOS)
    base_alpha = sq_alpha.split()[-1].resize((PNG_SIZE, PNG_SIZE),
                                              Image.LANCZOS)
    base = base_rgb.convert("RGBA")
    base.putalpha(base_alpha)

    # 第二遍：1024 输出上再修一次光晕（缩放可能带出新的浅色渗色）
    base = remove_outer_halo(base, bg, band=30, alpha_thr=40, light_thr=110)
    print("归一为:", base.size)
    _stats(base, "最终 1024")

    # 预览（蓝/黑/白底）供肉眼复核
    small = base.resize((256, 256), Image.LANCZOS)
    for nm, bgc in (("blue", (55, 115, 175)), ("black", (8, 8, 8)),
                    ("white", (245, 245, 245))):
        c = Image.new("RGB", (256, 256), bgc)
        c.paste(small, (0, 0), small)
        c.save(os.path.join(ASSETS, f"_prev_{nm}.png"))

    base.save(ICON_PNG)
    write_ico_bmp(base, ICON_ICO, ICO_SIZES)
    write_icns({s: base.resize((s, s), Image.LANCZOS) for _, s in ICNS_CHUNKS},
               ICON_ICNS)
    shutil.copy(ICON_PNG, os.path.join(DATA, "icon.png"))
    print("已写出 icon.png / icon.ico(BMP) / icon.icns / src/data/icon.png")
    print("预览: assets/_prev_blue.png / _prev_black.png / _prev_white.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
