"""Bespoke illustrations for the Red Flag Ritual deck — no stock photography.

Every image is drawn here as SVG (bottles, mist, monogram box, shelf, sachet, icons) and rendered to
transparent PNG with headless Chromium; the paper texture is generated with Pillow.
Run: python make_art.py   (writes art/*.png)
"""
import random
from pathlib import Path

from PIL import Image, ImageFilter
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
ART = HERE / "art"
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"

IVORY, INK, OX, GREEN, BRASS = "#F4EFE4", "#1B1A17", "#5E1A1D", "#24382B", "#A88A57"
SERIF, CASLON = "EB Garamond", "Libre Caslon Display"

DEFS = f"""
<defs>
  <linearGradient id="brass" x1="0" x2="1"><stop offset="0" stop-color="#7C5E2C"/><stop offset=".35" stop-color="#E4C98F"/>
    <stop offset=".55" stop-color="#B8954F"/><stop offset="1" stop-color="#6E5225"/></linearGradient>
  <linearGradient id="brassV" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#E9D29C"/><stop offset=".5" stop-color="#B8954F"/>
    <stop offset="1" stop-color="#7C5E2C"/></linearGradient>
  <linearGradient id="amber" x1="0" x2="1"><stop offset="0" stop-color="#3E1A07"/><stop offset=".25" stop-color="#8A4513"/>
    <stop offset=".5" stop-color="#B8691F"/><stop offset=".8" stop-color="#7A3A10"/><stop offset="1" stop-color="#3A1806"/></linearGradient>
  <linearGradient id="frost" x1="0" x2="1"><stop offset="0" stop-color="#D9D5CB"/><stop offset=".3" stop-color="#FBFAF6"/>
    <stop offset=".7" stop-color="#F1EEE7"/><stop offset="1" stop-color="#CFCAC0"/></linearGradient>
  <linearGradient id="glass" x1="0" x2="1"><stop offset="0" stop-color="#C9C3B6" stop-opacity=".95"/>
    <stop offset=".12" stop-color="#F7F4EC" stop-opacity=".9"/><stop offset=".5" stop-color="#E9E3D6" stop-opacity=".55"/>
    <stop offset=".88" stop-color="#F7F4EC" stop-opacity=".9"/><stop offset="1" stop-color="#BDB6A8" stop-opacity=".95"/></linearGradient>
  <linearGradient id="liquid" x1="0" x2="1"><stop offset="0" stop-color="#D8C49A"/><stop offset=".5" stop-color="#F0E2BE"/>
    <stop offset="1" stop-color="#CDB485"/></linearGradient>
  <linearGradient id="green" x1="0" x2="1"><stop offset="0" stop-color="#0F1B13"/><stop offset=".3" stop-color="#2C4535"/>
    <stop offset=".55" stop-color="#3B5A46"/><stop offset="1" stop-color="#101C14"/></linearGradient>
  <linearGradient id="black" x1="0" x2="1"><stop offset="0" stop-color="#0B0B0A"/><stop offset=".45" stop-color="#2A2926"/>
    <stop offset="1" stop-color="#0B0B0A"/></linearGradient>
  <linearGradient id="shine" x1="0" x2="1"><stop offset="0" stop-color="#fff" stop-opacity="0"/>
    <stop offset=".5" stop-color="#fff" stop-opacity=".55"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></linearGradient>
  <filter id="soft" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="9"/></filter>
  <filter id="soft2" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="3"/></filter>
  <filter id="mist" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="14"/></filter>
</defs>"""


def shadow(cx, y, w):
    return (f'<ellipse cx="{cx}" cy="{y}" rx="{w * .62}" ry="{w * .07}" fill="#3A2E20" opacity=".28" filter="url(#soft)"/>'
            f'<ellipse cx="{cx}" cy="{y}" rx="{w * .45}" ry="{w * .03}" fill="#2A2116" opacity=".35" filter="url(#soft2)"/>')


def label(cx, y, w, h, top, bottom, dark=False):
    fg = IVORY if dark else INK
    bg = "none" if dark else "#F7F2E8"
    return (f'<rect x="{cx - w / 2}" y="{y}" width="{w}" height="{h}" fill="{bg}" stroke="{BRASS}" stroke-width="1"/>'
            f'<path d="M{cx - 5} {y + h * .2} v{h * .16} M{cx - 5} {y + h * .2} l11 3.5 -11 3.5" stroke="{OX}" '
            f'stroke-width="1.6" fill="{OX}"/>'
            f'<text x="{cx}" y="{y + h * .58}" font-family="{CASLON}" font-size="{h * .2}" letter-spacing="{h * .06}" '
            f'fill="{fg}" text-anchor="middle">{top}</text>'
            f'<text x="{cx}" y="{y + h * .82}" font-family="{SERIF}" font-style="italic" font-size="{h * .13}" '
            f'fill="{fg}" text-anchor="middle" opacity=".8">{bottom}</text>')


# ---- bottle shapes (cx = centre x, base = y of the bottom) ---------------------------------------------
def vial(cx, base, s=1.0, code="SUD", name="Late Checkout"):
    w, h = 70 * s, 190 * s
    x, y = cx - w / 2, base - h
    return (shadow(cx, base, w) +
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{w * .22}" fill="url(#frost)"/>'
            f'<rect x="{x + w * .14}" y="{y + 8 * s}" width="{w * .12}" height="{h * .8}" rx="4" fill="url(#shine)" opacity=".6"/>'
            f'<rect x="{x - 2 * s}" y="{y - 42 * s}" width="{w + 4 * s}" height="{46 * s}" rx="{8 * s}" fill="#EFEBE2" stroke="#D3CEC3"/>'
            f'<rect x="{x - 2 * s}" y="{y - 42 * s}" width="{w + 4 * s}" height="{9 * s}" rx="{4 * s}" fill="#E2DDD2"/>'
            + label(cx, y + h * .3, w * .86, h * .38, code, name))


def amber(cx, base, s=1.0, code="DEW", name="Late Checkout"):
    w, h = 120 * s, 230 * s
    x, y = cx - w / 2, base - h
    neck = 34 * s
    body = (f'<path d="M{x} {base - 14 * s} V{y + 46 * s} Q{x} {y + 6 * s} {cx - neck / 2} {y - 4 * s} V{y - 22 * s} H{cx + neck / 2} '
            f'V{y - 4 * s} Q{x + w} {y + 6 * s} {x + w} {y + 46 * s} V{base - 14 * s} Q{x + w} {base} {x + w - 14 * s} {base} '
            f'H{x + 14 * s} Q{x} {base} {x} {base - 14 * s}Z" fill="url(#amber)"/>')
    cap = "".join(f'<rect x="{cx - 26 * s}" y="{y - 74 * s + i * 7 * s}" width="{52 * s}" height="{4 * s}" fill="#2B2A27"/>'
                  for i in range(8))
    return (shadow(cx, base, w) + body +
            f'<rect x="{x + w * .1}" y="{y + 30 * s}" width="{w * .1}" height="{h * .75}" rx="5" fill="url(#shine)" opacity=".5"/>'
            f'<rect x="{cx - 28 * s}" y="{y - 78 * s}" width="{56 * s}" height="{60 * s}" rx="{6 * s}" fill="url(#black)"/>' + cap
            + label(cx, y + h * .32, w * .78, h * .38, code, name))


def designer(cx, base, s=1.0, code="VEL", name="Late Checkout", tint="url(#liquid)"):
    w, h = 140 * s, 210 * s
    x, y = cx - w / 2, base - h
    t = 16 * s
    return (shadow(cx, base, w) +
            f'<path d="M{x + 18 * s} {y} H{x + w - 18 * s} L{x + w} {y + 18 * s} V{base - 18 * s} L{x + w - 18 * s} {base} '
            f'H{x + 18 * s} L{x} {base - 18 * s} V{y + 18 * s}Z" fill="url(#glass)" stroke="#B9B1A1" stroke-width="1.2"/>'
            f'<rect x="{x + t}" y="{y + h * .3}" width="{w - 2 * t}" height="{h * .7 - t * 1.6}" rx="{4 * s}" fill="{tint}" opacity=".85"/>'
            f'<rect x="{x + t * .6}" y="{y + t}" width="{w * .06}" height="{h - 2 * t}" fill="url(#shine)" opacity=".8"/>'
            f'<rect x="{cx - 22 * s}" y="{y - 14 * s}" width="{44 * s}" height="{16 * s}" fill="#D8D1C2"/>'
            f'<rect x="{cx - 34 * s}" y="{y - 78 * s}" width="{68 * s}" height="{66 * s}" rx="{3 * s}" fill="url(#brass)"/>'
            f'<rect x="{cx - 34 * s}" y="{y - 78 * s}" width="{68 * s}" height="{6 * s}" fill="#F0DCAA" opacity=".6"/>'
            + label(cx, y + h * .36, w * .66, h * .36, code, name))


def spray(cx, base, s=1.0, code="AIR", name="Room spray", body="url(#green)", dark=True):
    w, h = 104 * s, 250 * s
    x, y = cx - w / 2, base - h
    return (shadow(cx, base, w) +
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{10 * s}" fill="{body}"/>'
            f'<rect x="{x + w * .1}" y="{y + 10 * s}" width="{w * .1}" height="{h * .85}" rx="4" fill="url(#shine)" opacity=".45"/>'
            f'<rect x="{cx - 20 * s}" y="{y - 24 * s}" width="{40 * s}" height="{26 * s}" fill="url(#brass)"/>'
            f'<rect x="{cx - 15 * s}" y="{y - 64 * s}" width="{30 * s}" height="{40 * s}" rx="{4 * s}" fill="url(#brass)"/>'
            f'<rect x="{cx + 15 * s}" y="{y - 54 * s}" width="{9 * s}" height="{7 * s}" fill="#7C5E2C"/>'
            + label(cx, y + h * .34, w * .8, h * .32, code, name, dark=dark))


def pump(cx, base, s=1.0, code="SUD", name="Shampoo", body="url(#frost)", dark=False, tall=1.0):
    w, h = 104 * s, 240 * s * tall
    x, y = cx - w / 2, base - h
    return (shadow(cx, base, w) +
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{18 * s}" fill="{body}"/>'
            f'<rect x="{x + w * .1}" y="{y + 12 * s}" width="{w * .1}" height="{h * .82}" rx="4" fill="url(#shine)" opacity=".5"/>'
            f'<rect x="{cx - 16 * s}" y="{y - 30 * s}" width="{32 * s}" height="{32 * s}" fill="url(#brass)"/>'
            f'<rect x="{cx - 6 * s}" y="{y - 64 * s}" width="{12 * s}" height="{36 * s}" fill="url(#brass)"/>'
            f'<path d="M{cx - 12 * s} {y - 70 * s} H{cx + 46 * s} V{y - 58 * s} H{cx - 12 * s}Z" fill="url(#brassV)"/>'
            + label(cx, y + h * .3, w * .78, min(h * .36, 92 * s), code, name, dark=dark))


def diffuser(cx, base, s=1.0, code="LNG", name="Diffuser"):
    w, h = 130 * s, 150 * s
    x, y = cx - w / 2, base - h
    reeds = "".join(f'<line x1="{cx - 8 * s + i * 4 * s}" y1="{y - 6 * s}" x2="{cx - 70 * s + i * 30 * s}" y2="{y - 230 * s + (i % 2) * 20 * s}" '
                    f'stroke="#3A2C1E" stroke-width="{3 * s}" stroke-linecap="round"/>' for i in range(5))
    return (shadow(cx, base, w) + reeds +
            f'<path d="M{x + 10 * s} {base} Q{x - 6 * s} {y + 40 * s} {cx - 20 * s} {y} H{cx + 20 * s} Q{x + w + 6 * s} {y + 40 * s} '
            f'{x + w - 10 * s} {base}Z" fill="url(#amber)"/>'
            f'<rect x="{cx - 22 * s}" y="{y - 12 * s}" width="{44 * s}" height="{14 * s}" fill="url(#brass)"/>'
            + label(cx, y + h * .36, w * .66, h * .44, code, name))


def tube(cx, base, s=1.0, code="SUD", fill="#EDE2C8"):
    w, h = 34 * s, 200 * s
    x, y = cx - w / 2, base - h
    return (f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{w / 2}" fill="url(#glass)" stroke="#B9B1A1"/>'
            f'<rect x="{x + 4 * s}" y="{y + h * .32}" width="{w - 8 * s}" height="{h * .66}" rx="{(w - 8 * s) / 2}" fill="{fill}"/>'
            f'<rect x="{x + 5 * s}" y="{y + 8 * s}" width="{4 * s}" height="{h * .85}" fill="url(#shine)" opacity=".9"/>'
            f'<rect x="{x - 2 * s}" y="{y - 22 * s}" width="{w + 4 * s}" height="{26 * s}" rx="{4 * s}" fill="#B08B5E"/>'
            f'<rect x="{x - 2 * s}" y="{y - 22 * s}" width="{w + 4 * s}" height="{5 * s}" fill="#C9A77A"/>'
            f'<text x="{cx}" y="{y + h * .22}" font-family="{CASLON}" font-size="{12 * s}" fill="{INK}" text-anchor="middle" '
            f'letter-spacing="{1.5 * s}">{code}</text>')


def flag(x, y, s=1.0, color=OX):
    return (f'<line x1="{x}" y1="{y}" x2="{x}" y2="{y + 120 * s}" stroke="{color}" stroke-width="{5 * s}"/>'
            f'<path d="M{x} {y} L{x + 78 * s} {y + 22 * s} L{x} {y + 46 * s}Z" fill="{color}"/>')


# ---- scenes --------------------------------------------------------------------------------------------
def svg(w, h, body):
    return f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}">{DEFS}{body}</svg>'


def scenes():
    out = {}
    out["heirlooms"] = svg(900, 520, vial(170, 470, 1.3, "SUD", "House") + amber(450, 470, 1.25, "DEW", "Manor")
                           + designer(740, 470, 1.25, "VEL", "Estate"))
    out["tier_house"] = svg(300, 460, vial(150, 430, 1.45, "SUD", "House"))
    out["tier_manor"] = svg(300, 460, amber(150, 430, 1.1, "DEW", "Manor"))
    out["tier_estate"] = svg(300, 460, designer(150, 430, 1.15, "VEL", "Estate"))
    # the collection: six bottles, codes printed natively on the slide
    coll = [("SUD", "Shampoo", pump(0, 0, 1, "SUD", "Shampoo")),
            ("DEW", "Body wash", pump(0, 0, 1, "DEW", "Body wash", body="url(#amber)", dark=True)),
            ("VEL", "Body lotion", pump(0, 0, 1, "VEL", "Body lotion", tall=.82)),
            ("SLK", "Conditioner", pump(0, 0, 1, "SLK", "Conditioner", body="url(#green)", dark=True)),
            ("AIR", "Room spray", spray(0, 0, 1, "AIR", "Room spray")),
            ("LNG", "Diffuser", diffuser(0, 0, 1, "LNG", "Diffuser"))]
    for code, name, _ in coll:
        pass
    out["c_sud"] = svg(260, 420, pump(130, 400, 1, "SUD", "Shampoo"))
    out["c_dew"] = svg(260, 420, pump(130, 400, 1, "DEW", "Body wash", body="url(#amber)", dark=True))
    out["c_vel"] = svg(260, 420, pump(130, 400, 1, "VEL", "Body lotion", tall=.82))
    out["c_slk"] = svg(260, 420, pump(130, 400, 1, "SLK", "Conditioner", body="url(#green)", dark=True))
    out["c_air"] = svg(260, 420, spray(130, 400, 1, "AIR", "Room spray"))
    out["c_lng"] = svg(260, 420, diffuser(130, 400, 1, "LNG", "Diffuser"))
    # AIR mid-mist
    random.seed(7)
    dots = "".join(f'<circle cx="{random.gauss(560, 70) + i * 1.6}" cy="{random.gauss(250, 38 + i * .4)}" '
                   f'r="{random.uniform(1, 3.2)}" fill="#fff" opacity="{random.uniform(.35, .9):.2f}"/>' for i in range(260))
    out["mist"] = svg(900, 620,
                      '<path d="M345 245 C 460 170, 650 150, 860 200 L 860 320 C 650 360, 460 330, 345 262Z" '
                      'fill="#FFFFFF" opacity=".75" filter="url(#mist)"/>'
                      '<ellipse cx="640" cy="255" rx="200" ry="70" fill="#fff" opacity=".55" filter="url(#mist)"/>'
                      + dots + spray(260, 590, 1.6, "AIR", "Late Checkout"))
    # Ritual Classic kit: ink box, brass RFR foil, four tubes
    tubes = "".join(tube(330 + i * 80, 430, 1.0, c, f) for i, (c, f) in
                    enumerate((("SUD", "#EFE6CF"), ("DEW", "#E3C99A"), ("VEL", "#F4EEE2"), ("AIR", "#CFDCCB"))))
    out["kit"] = svg(900, 620,
                     shadow(450, 560, 600) +
                     '<rect x="170" y="10" width="560" height="450" fill="url(#black)"/>'            # lid, standing behind
                     '<rect x="190" y="30" width="520" height="410" fill="none" stroke="#8C6E3A" stroke-width="1.5"/>'
                     f'<text x="450" y="135" font-family="{CASLON}" font-size="96" fill="url(#brass)" text-anchor="middle" '
                     'letter-spacing="10">RFR</text>'
                     f'<text x="450" y="172" font-family="{SERIF}" font-size="17" fill="url(#brass)" text-anchor="middle" '
                     'letter-spacing="8">RED FLAG RITUAL</text>'
                     + tubes +
                     '<rect x="250" y="400" width="400" height="160" fill="url(#black)"/>'           # tray front
                     '<rect x="250" y="400" width="400" height="4" fill="#3A3833"/>'
                     f'<text x="450" y="500" font-family="{CASLON}" font-size="30" fill="url(#brass)" text-anchor="middle" '
                     'letter-spacing="6">RITUAL CLASSIC</text>')
    # bathroom shelf, morning light
    beams = "".join(f'<polygon points="{60 + i * 140},0 {150 + i * 140},0 {480 + i * 140},620 {390 + i * 140},620" '
                    f'fill="#FFF8E6" opacity=".22"/>' for i in range(4))
    towel = "".join(f'<rect x="640" y="{356 - i * 22}" width="200" height="22" rx="10" fill="{c}"/>'
                    for i, c in enumerate(("#E8E1D3", "#F4EFE5", "#E8E1D3", "#F7F3EA")))
    out["shelf"] = svg(900, 620,
                       f'<rect width="900" height="620" fill="#EEE7DA"/>' + beams +
                       '<rect x="0" y="378" width="900" height="34" fill="#D9D0C0"/><rect x="0" y="412" width="900" height="8" fill="#C7BCA9"/>'
                       '<rect x="0" y="378" width="900" height="3" fill="#F6F1E7"/>'
                       '<rect x="120" y="420" width="16" height="70" fill="url(#brassV)"/><rect x="760" y="420" width="16" height="70" fill="url(#brassV)"/>'
                       + pump(140, 378, 1.05, "SUD", "Shampoo") + pump(280, 378, 1.05, "DEW", "Body wash", body="url(#amber)", dark=True)
                       + pump(415, 378, .95, "VEL", "Lotion", tall=.82) + spray(540, 378, .95, "AIR", "Room spray")
                       + towel)
    # a plain, forgettable sachet — slightly out of focus
    out["sachet"] = svg(700, 520,
                        '<g filter="url(#soft2)" transform="rotate(-9 350 260)">'
                        '<path d="M190 120 h320 l-6 10 6 10 v260 l-6 10 6 10 h-320 l6 -10 -6 -10 v-260 l6 -10z" fill="#E7E6E3" stroke="#CFCDC8"/>'
                        '<rect x="190" y="150" width="320" height="6" fill="#D6D4CF"/>'
                        '<text x="350" y="280" font-family="Liberation Sans" font-size="34" fill="#9E9C97" text-anchor="middle">SHAMPOO</text>'
                        '<text x="350" y="318" font-family="Liberation Sans" font-size="20" fill="#B2B0AB" text-anchor="middle">10 ml</text></g>')
    # thin brass line icons
    st = f'fill="none" stroke="{BRASS}" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"'
    out["i_plan"] = svg(200, 200, f'<g {st}><rect x="35" y="70" width="34" height="96" rx="8"/><rect x="83" y="50" width="34" height="116" rx="8"/>'
                                  f'<rect x="131" y="80" width="34" height="86" rx="8"/><path d="M95 30 v20 M105 30 v20"/></g>')
    out["i_deliver"] = svg(200, 200, f'<g {st}><path d="M30 70 L100 38 L170 70 L100 102Z"/><path d="M30 70 V140 L100 172 V102"/>'
                                     f'<path d="M170 70 V140 L100 172"/><path d="M65 54 L135 86"/></g>')
    out["i_notice"] = svg(200, 200, f'<g {st}><path d="M40 50 h120 a12 12 0 0 1 12 12 v60 a12 12 0 0 1 -12 12 h-70 l-30 26 v-26 h-20 '
                                    f'a12 12 0 0 1 -12 -12 v-60 a12 12 0 0 1 12 -12z"/><path d="M100 72 l7 14 15 2 -11 10 3 15 -14 -7 -14 7 3 -15 -11 -10 15 -2z"/></g>')
    out["i_repeat"] = svg(200, 200, f'<g {st}><path d="M152 84 A56 56 0 0 0 52 70"/><path d="M48 116 A56 56 0 0 0 148 130"/>'
                                    f'<path d="M40 52 L52 70 L70 60"/><path d="M160 148 L148 130 L130 140"/></g>')
    out["flag"] = svg(120, 140, flag(20, 10, 1.0))
    out["flag_ivory"] = svg(120, 140, flag(20, 10, 1.0, IVORY))
    return out


def paper(name, base, size=(2400, 1350), grain=6, seed=3):
    random.seed(seed)
    img = Image.new("RGB", size, base)
    noise = Image.effect_noise(size, 18).convert("L").filter(ImageFilter.GaussianBlur(.6))
    tint = Image.new("RGB", size, tuple(max(0, c - grain * 2) for c in Image.new("RGB", (1, 1), base).getpixel((0, 0))))
    img = Image.composite(tint, img, noise.point(lambda v: int(v * .22)))
    img.save(ART / f"{name}.jpg", quality=90)


def main():
    ART.mkdir(exist_ok=True)
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME)
        pg = b.new_page(device_scale_factor=3)
        for name, s in scenes().items():
            w = int(s.split('width="')[1].split('"')[0])
            h = int(s.split('height="')[1].split('"')[0])
            pg.set_viewport_size({"width": w, "height": h})
            pg.set_content(f'<html><body style="margin:0;background:transparent">{s}</body></html>')
            pg.wait_for_timeout(120)
            pg.screenshot(path=str(ART / f"{name}.png"), omit_background=True,
                          clip={"x": 0, "y": 0, "width": w, "height": h})
        b.close()
    paper("paper_ivory", "#F4EFE4")
    paper("paper_oxblood", "#5E1A1D", seed=5)
    paper("paper_ink", "#1B1A17", seed=9)
    paper("paper_green", "#24382B", seed=11)
    paper("paper_stone", "#D9D0C0", seed=13)
    print(sorted(x.name for x in ART.iterdir()))


if __name__ == "__main__":
    main()
