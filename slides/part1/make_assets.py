"""voice-deck スキルのトンマナに合わせた背景・飾りの画像を生成する（外部画像は使わない）。"""
import math
import os

from PIL import Image, ImageDraw, ImageFilter

OUT = os.path.join(os.path.dirname(__file__), "assets")
W, H = 1920, 1080


def hex_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def lerp(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def gradient():
    # 左下 #A6BEF7 → #D6CDF8 → #E8CCF4 → 右上 #FBBCEC、左上に淡い水色 #E6FAFA を重ねる
    stops = [(0.0, hex_rgb("A6BEF7")), (0.38, hex_rgb("D6CDF8")),
             (0.68, hex_rgb("E8CCF4")), (1.0, hex_rgb("FBBCEC"))]
    img = Image.new("RGB", (W, H))
    px = img.load()
    for y in range(H):
        for x in range(W):
            t = (x / W + (1 - y / H)) / 2  # 左下=0, 右上=1
            for i in range(len(stops) - 1):
                t0, c0 = stops[i]
                t1, c1 = stops[i + 1]
                if t <= t1:
                    c = lerp(c0, c1, (t - t0) / (t1 - t0))
                    break
            # 左上の淡い水色
            d = math.hypot(x / W, y / H) / 0.75
            a = max(0.0, 1 - d) ** 1.6 * 0.85
            px[x, y] = lerp(c, hex_rgb("E6FAFA"), a)
    return img


def deco_line():
    # 手書き風の飾り線（細い曲線）
    s = 4
    w, h = 900 * s, 90 * s
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    col = hex_rgb("897D74") + (255,)
    pts = []
    # 左端の小さなループ → ゆるい波 → 右へ細く抜ける
    for i in range(0, 121):
        t = i / 120 * 2 * math.pi
        pts.append((120 * s + 70 * s * math.cos(t + math.pi) * 0.55 + i * 0.35 * s,
                    48 * s + 26 * s * math.sin(t + math.pi)))
    x0, y0 = pts[-1]
    for i in range(1, 401):
        u = i / 400
        x = x0 + u * (820 * s - x0)
        y = y0 - 22 * s * math.sin(u * math.pi * 1.15) * (1 - u * 0.4)
        pts.append((x, y))
    for i in range(len(pts) - 1):
        u = i / len(pts)
        width = max(1, round((5.5 - 4.2 * u) * s))
        d.line([pts[i], pts[i + 1]], fill=col, width=width)
    return img.resize((w // s, h // s), Image.LANCZOS)


def gold_dot(r):
    s = 4
    img = Image.new("RGBA", (r * 2 * s, r * 2 * s), (0, 0, 0, 0))
    ImageDraw.Draw(img).ellipse([0, 0, r * 2 * s - 1, r * 2 * s - 1],
                                fill=hex_rgb("F5DE81") + (255,))
    return img.resize((r * 2, r * 2), Image.LANCZOS)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    gradient().filter(ImageFilter.GaussianBlur(2)).save(os.path.join(OUT, "bg-gradient.png"))
    deco_line().save(os.path.join(OUT, "deco-line.png"))
    print("ok")
