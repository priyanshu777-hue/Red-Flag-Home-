"""Backgrounds for The Nights Promise · USA deck: aurora gradient meshes (Pillow) and a glowing moon orb.
No stock photography. Run: python make_art.py"""
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ART = Path(__file__).resolve().parent / "art"
SIZE = (2400, 1350)


def mesh(name, base, blobs, blur=260, grain=True, seed=1):
    random.seed(seed)
    img = Image.new("RGB", SIZE, base)
    layer = Image.new("RGB", SIZE, base)
    d = ImageDraw.Draw(layer)
    for (cx, cy, r, col) in blobs:
        d.ellipse((cx * SIZE[0] - r, cy * SIZE[1] - r, cx * SIZE[0] + r, cy * SIZE[1] + r), fill=col)
    img = layer.filter(ImageFilter.GaussianBlur(blur))
    if grain:
        noise = Image.effect_noise(SIZE, 10).convert("L")
        img = Image.composite(Image.new("RGB", SIZE, (255, 255, 255)), img, noise.point(lambda v: 10 if v > 140 else 0))
    img.save(ART / f"{name}.jpg", quality=92)


def orb():
    s = 900
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    glow = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse((90, 90, s - 90, s - 90), fill=(255, 196, 120, 120))
    glow = glow.filter(ImageFilter.GaussianBlur(70))
    body = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    bd = ImageDraw.Draw(body)
    for i in range(260, 0, -1):                       # radial shading, light from top-left
        t = i / 260
        c = (int(255 - 40 * t), int(236 - 70 * t), int(200 - 110 * t), 255)
        off = (1 - t) * -60
        bd.ellipse((s / 2 - i + off, s / 2 - i + off, s / 2 + i + off * .2, s / 2 + i + off * .2), fill=c)
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).ellipse((s / 2 - 260, s / 2 - 260, s / 2 + 260, s / 2 + 260), fill=255)
    img = Image.alpha_composite(glow, Image.composite(body, Image.new("RGBA", (s, s), (0, 0, 0, 0)), mask))
    img.save(ART / "orb.png")


if __name__ == "__main__":
    ART.mkdir(exist_ok=True)
    mesh("aurora", (9, 12, 28), [(.85, .15, 520, (88, 70, 229)), (.62, .95, 600, (255, 92, 77)),
                                 (1.0, .65, 420, (139, 92, 246)), (.1, .1, 380, (20, 40, 110))])
    mesh("aurora2", (9, 12, 28), [(.1, .9, 560, (88, 70, 229)), (.95, .05, 480, (255, 92, 77)),
                                  (.5, .5, 300, (40, 30, 110))], seed=2)
    mesh("aurora3", (9, 12, 28), [(.5, 1.05, 700, (255, 92, 77)), (.15, .2, 480, (88, 70, 229)),
                                  (.9, .25, 420, (139, 92, 246))], seed=3)
    mesh("dawn", (247, 245, 241), [(.95, .0, 520, (255, 214, 205)), (.0, 1.0, 520, (221, 216, 255))], blur=300,
         grain=False)
    orb()
    print(sorted(p.name for p in ART.iterdir()))
