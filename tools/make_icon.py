# -*- coding: utf-8 -*-
"""Regenerate the app icon (PNG + ICO) for Thesis Format Doctor Global.

Run:  python tools/make_icon.py
Needs Pillow:  python -m pip install pillow
"""
import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
ASSETS = os.path.join(ROOT, "assets")
W = 1024


def _rounded(draw, box, r, fill):
    draw.rounded_rectangle(box, r, fill=fill)


def make():
    os.makedirs(ASSETS, exist_ok=True)
    base = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    d = ImageDraw.Draw(base)
    bg = (31, 78, 70, 255)  # deep reed green #1F4E46
    _rounded(d, [40, 40, W - 40, W - 40], 180, bg)

    pg = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    pd = ImageDraw.Draw(pg)
    _rounded(pd, [300, 210, W - 300, W - 150], 28, (255, 255, 255, 238))
    for i, y in enumerate(range(330, 720, 72)):
        w = 300 if i % 3 == 2 else 384
        _rounded(pd, [362, y, 362 + w, y + 34], 12, (176, 196, 186, 255))

    cc = (W - 248, W - 248)
    pd.ellipse([cc[0] - 112, cc[1] - 112, cc[0] + 112, cc[1] + 112], fill=(46, 160, 110, 255))
    pd.line([cc[0] - 58, cc[1], cc[0] - 12, cc[1] + 48], fill=(255, 255, 255, 255), width=36, joint="curve")
    pd.line([cc[0] - 12, cc[1] + 48, cc[0] + 78, cc[1] - 58], fill=(255, 255, 255, 255), width=36, joint="curve")

    img = Image.alpha_composite(base, pg)
    img.save(os.path.join(ASSETS, "icon.png"))
    img.save(os.path.join(ASSETS, "icon.ico"),
             sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    print("icons written to", ASSETS)


if __name__ == "__main__":
    make()
