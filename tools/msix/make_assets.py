# -*- coding: utf-8 -*-
"""从 assets/icon.png 生成 MSIX 需要的商店素材。

MSIX 对素材有硬性要求（必须是真 PNG、必须是指定尺寸），缺一个
makeappx 或商店校验就会报错。这里一次性生成齐，产物提交进仓库，
CI 里就不需要装 Pillow。

用法（在仓库根目录）：
    python tools/msix/make_assets.py
"""
import os

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
SRC = os.path.join(ROOT, "assets", "icon.png")
OUT = os.path.join(HERE, "assets")

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
    if not os.path.isfile(SRC):
        raise SystemExit("找不到源图标：%s" % SRC)
    os.makedirs(OUT, exist_ok=True)
    src = Image.open(SRC).convert("RGBA")
    print("源图标 %s  %dx%d" % (SRC, src.width, src.height))
    for name, (w, h) in TARGETS.items():
        out = os.path.join(OUT, name)
        fit(src, w, h).save(out, "PNG", optimize=True)
        print("  %-24s %dx%d  %d 字节" % (name, w, h, os.path.getsize(out)))


if __name__ == "__main__":
    main()
