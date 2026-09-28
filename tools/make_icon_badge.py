# -*- coding: utf-8 -*-
"""从「圆形徽章 PNG」生成 app 图标三件套：icon.png / icon.ico / icon.icns。

源图应为带透明背景的圆形徽章（如羽毛笔+青圆底）。脚本会：
1. 把它居中缩放后贴到 1024×1024 透明画布（留出 iOS/macOS 安全边距）；
2. 输出 1024×1024 PNG；
3. 输出多尺寸 ICO（Windows，最高 256）；
4. 输出多尺寸 ICNS（macOS）。

用法：
  python tools/make_icon_badge.py "源图.png" --outdir assets --name icon
"""

from __future__ import annotations

import argparse
import io
import os
import shutil
import struct
import sys

from PIL import Image, ImageOps

PNG_SIZE = 1024
SAFE_FIT = 0.84                          # 84% 安全区，圆角/阴影不会被裁
ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]
ICNS_CHUNKS = [
    ("icp4", 16), ("icp5", 32), ("icp6", 64),
    ("ic07", 128), ("ic08", 256), ("ic09", 512), ("ic10", 1024),
]


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


def make_badge_icon(src_path: str, out_dir: str, name: str = "icon") -> None:
    os.makedirs(out_dir, exist_ok=True)
    src = Image.open(src_path).convert("RGBA")
    # 统一缩放到安全区内，保持比例，允许放大/缩小
    fit_size = int(PNG_SIZE * SAFE_FIT)
    scaled = ImageOps.contain(src, (fit_size, fit_size), Image.LANCZOS)
    # 居中贴到 1024×1024 透明画布
    canvas = Image.new("RGBA", (PNG_SIZE, PNG_SIZE), (0, 0, 0, 0))
    x = (PNG_SIZE - scaled.width) // 2
    y = (PNG_SIZE - scaled.height) // 2
    canvas.paste(scaled, (x, y), scaled)

    # 1) PNG
    png_path = os.path.join(out_dir, f"{name}.png")
    canvas.save(png_path, "PNG")
    print("png ->", png_path, "size:", os.path.getsize(png_path))

    # 2) ICO
    ico_path = os.path.join(out_dir, f"{name}.ico")
    canvas.save(ico_path, format="ICO", sizes=[(s, s) for s in ICO_SIZES])
    print("ico ->", ico_path, "sizes:", ICO_SIZES)

    # 3) ICNS：为每个 chunk 生成对应尺寸 PNG 再写入容器
    icns_pngs = {PNG_SIZE: canvas}
    for _, size in ICNS_CHUNKS:
        if size == PNG_SIZE:
            continue
        icns_pngs[size] = canvas.resize((size, size), Image.LANCZOS)
    icns_path = os.path.join(out_dir, f"{name}.icns")
    write_icns(icns_pngs, icns_path)
    print("icns ->", icns_path, "chunks:", [c for c, _ in ICNS_CHUNKS])


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="圆形徽章 PNG -> PNG/ICO/ICNS")
    ap.add_argument("src", nargs="?", help="源 PNG（带透明背景的圆形徽章）")
    ap.add_argument("--outdir", default="assets", help="输出目录")
    ap.add_argument("--name", default="icon", help="输出文件名（不含扩展名）")
    ap.add_argument("--copy-source", default="assets/logo_badge.png",
                      help="把源图复制到仓库的哪个路径（留档/CI 复用）")
    args = ap.parse_args(argv)

    # 默认源图路径（本次用户发的文件），可被命令行覆盖
    src = args.src or r"D:\AgentSpace\素材中心\UI设计\03_圆形徽章logo_circular_badge.png"
    if not os.path.exists(src):
        print("源图不存在:", src, file=sys.stderr)
        return 1

    outdir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), args.outdir)
    repo_src = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), args.copy_source)
    os.makedirs(os.path.dirname(repo_src), exist_ok=True)
    shutil.copy2(src, repo_src)
    print("source copied ->", repo_src)

    make_badge_icon(repo_src, outdir, args.name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
