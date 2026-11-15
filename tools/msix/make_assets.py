# -*- coding: utf-8 -*-
"""从各产品的图标生成 MSIX 所需的商店素材。

MSIX 对素材有硬性要求（必须是真 PNG、必须是指定尺寸），缺一个
makeappx 或商店校验就会报错。这里一次性生成齐，产物提交进仓库，
CI 里就不需要装 Pillow。

三个产品**品牌图标各不相同**，绝不能共用一套（商店磁贴会显示错图标）：
    paperformatpro -> 本仓 assets/icon.png
    mdf-stu        -> tfd-desktop 的 tfd_app/assets/icon.png
    mdf-advisor    -> tfd-mentor 的 tfd_app/assets/icon.png

用法（在仓库根目录）：
    python tools/msix/make_assets.py                                  # 默认 paperformatpro
    python tools/msix/make_assets.py --product mdf-stu --src <icon.png>
    python tools/msix/make_assets.py --product mdf-advisor --src <icon.png>
产物落在 tools/msix/assets/<product>/，make_msix.py 会优先用这个目录。
"""
import argparse
import os

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
DEFAULT_SRC = os.path.join(ROOT, "assets", "icon.png")

# 文件名 -> (宽, 高)
TARGETS = {
    "StoreLogo.png": (50, 50),              # Properties/Logo
    "Square44x44Logo.png": (44, 44),        # 开始菜单小图标
    "Square150x150Logo.png": (150, 150),    # 主磁贴
    "Square71x71Logo.png": (71, 71),        # 小磁贴
    "Square310x310Logo.png": (310, 310),    # 大磁贴
    "Wide310x150Logo.png": (310, 150),      # 宽磁贴（居中摆放，不拉伸）
    "SplashScreen.png": (620, 300),         # 启动画面（同样居中）
}


def fit(im, w, h):
    """等比缩放后居中贴到 w×h 透明画布上，绝不拉伸变形。"""
    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    scale = min(w / im.width, h / im.height)
    nw, nh = max(1, int(im.width * scale)), max(1, int(im.height * scale))
    mini = im.resize((nw, nh), Image.LANCZOS)
    canvas.paste(mini, ((w - nw) // 2, (h - nh) // 2), mini)
    return canvas


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--product", default="paperformatpro", help="输出到 tools/msix/assets/<product>/")
    ap.add_argument("--src", default=None, help="源图标 PNG（缺省用本仓 assets/icon.png）")
    ap.add_argument("--out", default=None, help="直接指定输出目录（覆盖 --product 推出的路径）")
    args = ap.parse_args()

    src_path = args.src or DEFAULT_SRC
    if not os.path.isfile(src_path):
        raise SystemExit("找不到源图标：%s" % src_path)
    out_dir = args.out or os.path.join(HERE, "assets", args.product)
    os.makedirs(out_dir, exist_ok=True)

    src = Image.open(src_path).convert("RGBA")
    print("源图标 %s  %dx%d" % (src_path, src.width, src.height))
    for name, (w, h) in TARGETS.items():
        out = os.path.join(out_dir, name)
        fit(src, w, h).save(out, "PNG", optimize=True)
        print("  %-24s %dx%d  %d 字节" % (name, w, h, os.path.getsize(out)))
    print("输出目录 %s" % out_dir)


if __name__ == "__main__":
    main()
