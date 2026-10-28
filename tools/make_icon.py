#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成应用图标 assets/studio.ico（蓝底 + PS 字标）。"""
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(os.path.dirname(HERE), "assets")
S = 512


def gradient(size, top, bottom):
    base = Image.new("RGB", (1, 2))
    base.putpixel((0, 0), top)
    base.putpixel((0, 1), bottom)
    return base.resize((size, size), Image.BICUBIC)


def build():
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    grad = gradient(S, (91, 156, 255), (24, 61, 168))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.23), fill=255)
    img.paste(grad, (0, 0), mask)

    # 内侧高光
    hi_mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(hi_mask).rounded_rectangle(
        [int(S * 0.05), int(S * 0.05), int(S * 0.95), int(S * 0.95)], radius=int(S * 0.19), fill=255
    )
    hi = Image.new("RGBA", (S, S), (255, 255, 255, 26))
    img.paste(hi, (0, 0), hi_mask)

    # 光圈装饰
    d = ImageDraw.Draw(img)
    cx, cy, r = S / 2, S / 2, int(S * 0.37)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=(255, 255, 255, 60), width=int(S * 0.018))

    # PS 字标
    font_path = r"C:\Windows\Fonts\segoeuib.ttf"
    font = ImageFont.truetype(font_path, int(S * 0.44))
    d.text((cx, cy - S * 0.01), "PS", font=font, fill=(255, 255, 255, 255), anchor="mm")

    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, "studio.ico")
    img.save(path, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("icon:", path)
    return path


if __name__ == "__main__":
    build()
