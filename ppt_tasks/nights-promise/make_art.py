"""Art for The Nights Promise deck: midnight-graded frame from Red Flag's own whyhost.mp4, a drawn moon
(SVG via headless Chromium) and a generated star field. No stock photography. Run: python make_art.py"""
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageOps
from playwright.sync_api import sync_playwright

ART = Path(__file__).resolve().parent / "art"
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
NAVY, MOON = (14, 20, 36), (241, 235, 221)


def night_photo():
    src = Image.open(ART / "night_8.jpg").convert("L")
    ImageOps.colorize(src, black=(8, 12, 24), mid=(52, 64, 92), white=(246, 236, 210)).save(ART / "night.jpg", quality=90)


def stars(name, size=(2400, 1350), n=900, seed=4):
    random.seed(seed)
    img = Image.new("RGB", size, NAVY)
    glow = Image.new("RGB", size, (0, 0, 0))
    d, g = ImageDraw.Draw(img), ImageDraw.Draw(glow)
    for _ in range(n):
        x, y = random.uniform(0, size[0]), random.uniform(0, size[1])
        r = random.choice([.6, .8, 1, 1, 1.2, 1.6, 2.2])
        a = random.randint(70, 230)
        c = tuple(int(NAVY[i] + (MOON[i] - NAVY[i]) * a / 255) for i in range(3))
        d.ellipse((x - r, y - r, x + r, y + r), fill=c)
        if r > 1.5:
            g.ellipse((x - 6, y - 6, x + 6, y + 6), fill=(60, 60, 70))
    img = Image.blend(img, Image.eval(Image.composite(glow, img, glow.convert("L")), lambda v: v), .0)
    img.paste(Image.new("RGB", size, (0, 0, 0)), mask=None) if False else None
    img = Image.composite(Image.new("RGB", size, (40, 46, 66)), img, glow.convert("L").filter(ImageFilter.GaussianBlur(5)).point(lambda v: v // 3))
    img.save(ART / f"{name}.jpg", quality=90)


MOON_SVG = """<svg xmlns="http://www.w3.org/2000/svg" width="800" height="800" viewBox="0 0 800 800">
<defs>
 <radialGradient id="halo"><stop offset=".55" stop-color="#F1E3BC" stop-opacity=".35"/><stop offset="1" stop-color="#F1E3BC" stop-opacity="0"/></radialGradient>
 <radialGradient id="body" cx=".38" cy=".35" r=".75"><stop offset="0" stop-color="#FBF5E6"/><stop offset=".6" stop-color="#EADBB4"/><stop offset="1" stop-color="#C9B07A"/></radialGradient>
 <filter id="b"><feGaussianBlur stdDeviation="11"/></filter>
</defs>
<circle cx="400" cy="400" r="390" fill="url(#halo)"/>
<circle cx="400" cy="400" r="230" fill="url(#body)"/>
<g fill="#B79E6A" opacity=".2" filter="url(#b)">
 <circle cx="330" cy="330" r="46"/><circle cx="470" cy="300" r="26"/><circle cx="455" cy="460" r="58"/>
 <circle cx="320" cy="470" r="22"/><circle cx="520" cy="380" r="16"/><circle cx="380" cy="560" r="30"/></g>
</svg>"""


def main():
    night_photo()
    stars("stars")
    stars("stars_soft", n=380, seed=9)
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME)
        pg = b.new_page(device_scale_factor=2, viewport={"width": 800, "height": 800})
        pg.set_content(f'<html><body style="margin:0;background:transparent">{MOON_SVG}</body></html>')
        pg.screenshot(path=str(ART / "moon.png"), omit_background=True)
        b.close()
    print(sorted(x.name for x in ART.iterdir()))


if __name__ == "__main__":
    main()
