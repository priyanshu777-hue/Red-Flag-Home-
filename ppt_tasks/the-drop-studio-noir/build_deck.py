"""Build The Drop · Studio Noir customer deck — Red Flag Homes Network.

Build Mode with the pptx-designer public API. Fonts: Cormorant Garamond + Plus Jakarta Sans (see fonts/).
Motion: motion.py (shared house style). Run: python build_deck.py
"""
from io import BytesIO
from pathlib import Path

from PIL import Image
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from pptx_designer import Presentation
from pptx_designer.tools.images import gradient_mask_image
from pptx_designer.tools.layout import add_slide, clean_save
from pptx_designer.tools.shapes import oval, rect, rrect
from pptx_designer.tools.text import text

import motion

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
IMG = HERE / "assets"
OUT = HERE / "output"
LOGO = str(ROOT / "logobg.png")

SERIF, SANS = "Cormorant Garamond", "Plus Jakarta Sans"
W, H, M = 13.333, 7.5, 0.75
TOTAL = 18

NOIR = {"bg": "#151311", "paper": "#201C19", "ink": "#EFE6D8", "body": "#CFC5B6", "muted": "#958A7B",
        "hair": "#39322B"}
OAT = {"bg": "#EDE6DA", "paper": "#E2D8C8", "ink": "#1B1815", "body": "#4A423A", "muted": "#8A7F71",
       "hair": "#D2C6B4"}
ACC = {"caramel": "#B8845A", "brass": "#C4A06A", "green": "#6E7B5F", "red": "#D93A2F", "white": "#FFFFFF",
       "charcoal": "#2A2724", "wood": "#8A5F3C", "cream": "#E9DFCC", "black": "#0E0D0C"}


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.pal = NOIR
        self.plan = {}

    def col(self, key):
        return key if key.startswith("#") else (self.pal.get(key) or ACC[key])

    @property
    def C(self):
        return {"background": self.pal["bg"], "text_body": self.pal["body"], "text_dark": self.pal["ink"],
                "text_muted": self.pal["muted"], "font_body": SANS, "font_heading": SERIF}

    # ---- primitives -------------------------------------------------------------------------
    def T(self, s, x, y, w, h, txt, size=12, font=SANS, color="body", italic=False, bold=False,
          spacing=None, align="left", leading=None, caps=False, strike=False):
        box = text(s, x, y, w, h, txt.upper() if caps else txt, font_size=size, color=self.col(color),
                   bold=bold, align=align, font_name=font, C=self.C)
        tf = box.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for p in tf.paragraphs:
            if leading:
                p.line_spacing = leading
            for r in p.runs:
                r.font.italic = italic
                if spacing is not None:
                    r.font._element.set("spc", str(spacing))
                if strike:
                    r.font._element.set("strike", "sngStrike")
        return box

    def H1(self, s, x, y, w, h, txt, size=40, color="ink", italic=False, align="left", leading=0.92, bold=False):
        return self.T(s, x, y, w, h, txt, size=size, font=SERIF, color=color, italic=italic, align=align,
                      leading=leading, bold=bold)

    def label(self, s, x, y, w, txt, color="caramel", size=8, align="left"):
        return self.T(s, x, y, w, 0.25, txt, size=size, color=color, bold=True, spacing=300, caps=True,
                      align=align)

    def bullets(self, s, x, y, w, h, items, size=11, color="body", gap=6, mark="—", mark_color="caramel"):
        box = self.T(s, x, y, w, h, "", size=size, color=color)
        tf = box.text_frame
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(gap)
            p.line_spacing = 1.25
            for t, c, b in ((f"{mark}   ", mark_color, True), (it, color, False)):
                r = p.add_run()
                r.text = t
                r.font.size, r.font.name, r.font.bold = Pt(size), SANS, b
                r.font.color.rgb = RGBColor.from_string(self.col(c)[1:])
        return box

    def box(self, s, x, y, w, h, fill, line=None, round_=False, radius=0.06):
        f = (rrect if round_ else rect)(s, x, y, w, h, self.col(fill), line=self.col(line) if line else None,
                                         C=self.C)
        if round_:
            f.adjustments[0] = radius
        return f

    def rule(self, s, x, y, w, color="hair", t=0.012):
        return self.box(s, x, y, w, t, color)

    def photo(self, s, x, y, w, h, name, anchor=0.5, tag=None):
        """Cover-fit crop (as pptx_designer's cover_image) embedded as JPEG to keep the file small."""
        img = Image.open(IMG / f"{name}.jpg").convert("RGB")
        iw, ih = img.size
        r = w / h
        if iw / ih > r:
            cw, ch = int(ih * r), ih
            x0, y0 = (iw - cw) // 2, 0
        else:
            cw, ch = iw, int(iw / r)
            x0, y0 = 0, int((ih - ch) * anchor)
        img = img.crop((x0, y0, x0 + cw, y0 + ch))
        px = min(cw, int(w * 180))
        img = img.resize((px, max(1, int(px / r))), Image.LANCZOS)
        buf = BytesIO()
        img.save(buf, "JPEG", quality=87, optimize=True, progressive=True)
        buf.seek(0)
        pic = s.shapes.add_picture(buf, Inches(x), Inches(y), Inches(w), Inches(h))
        pic.name = "static-photo"
        if tag:
            t = self.T(s, x + 0.15, y + h - 0.35, w - 0.3, 0.25, tag, size=7, color="white", bold=True,
                       spacing=250, caps=True)
            t.name = "static-tag"
        return pic

    def mask(self, s, x, y, w, h, direction, a0=100, a1=0, color=None):
        shp = gradient_mask_image(s, x, y, w, h, bg_color=color or self.pal["bg"], direction=direction,
                                  alpha_start=a0, alpha_end=a1)
        shp.name = "static-mask"
        return shp

    def shade(self, s, x, y, w, h, pct):
        shp = rect(s, x, y, w, h, "#000000", C=self.C)
        shp.name = "static-shade"
        clr = shp.fill._xPr.find(qn("a:solidFill"))[0]
        clr.append(clr.makeelement(qn("a:alpha"), {"val": str(pct * 1000)}))
        return shp

    def logo(self, s, x, y, size):
        pic = s.shapes.add_picture(LOGO, Inches(x), Inches(y), Inches(size), Inches(size))
        pic.name = "!!logo"
        return pic

    # ---- scaffolding -------------------------------------------------------------------------
    def slide(self, light=False, chrome=True, transition=("fade", False)):
        self.pal = OAT if light else NOIR
        s = add_slide(self.prs)
        n = len(self.prs.slides)
        self.plan[n] = transition
        rect(s, 0, 0, W, H, self.pal["bg"], C=self.C).name = "static-bg"
        if chrome:
            self.logo(s, W - M - 0.55, 0.32, 0.55)
            for shp in (self.label(s, M, 0.52, 6, "The Drop  ·  Studio Noir", color="muted", size=7),
                        self.label(s, M, H - 0.5, 4, "redflaghomes.in", color="muted", size=7),
                        self.T(s, W - M - 1.5, H - 0.5, 1.5, 0.25, f"{n:02d} / {TOTAL:02d}", size=7,
                               color="muted", bold=True, spacing=200, align="right")):
                shp.name = "chrome"
        return s

    def header(self, s, kicker, title, x=M, y=1.1, w=10, size=40):
        self.label(s, x, y, 8, kicker)
        self.H1(s, x, y + 0.3, w, 1.0, title, size=size)

    def save(self):
        for sl in self.prs.slides:
            tree = sl.shapes._spTree
            for shp in [x for x in sl.shapes if x.name in ("!!logo", "static-tag")]:
                tree.remove(shp._element)
                tree.append(shp._element)
        for i, sl in enumerate(self.prs.slides, start=1):
            kind, black = self.plan.get(i, ("fade", False))
            motion.transition(sl, kind, through_black=black)
            motion.choreograph(sl)
        OUT.mkdir(exist_ok=True)
        path = OUT / "TheDrop_Studio_Noir.pptx"
        clean_save(self.prs, str(path))
        print(path)


SAMPLE = "Sample image"


def build():
    d = Deck()
    T, H1, label = d.T, d.H1, d.label

    # 1 — Cover -------------------------------------------------------------------------------
    s = d.slide(chrome=False, transition=("fade", True))
    d.photo(s, 7.35, 0, W - 7.35, H, "sample_studio", anchor=0.35, tag=SAMPLE)
    d.mask(s, 7.35, 0, 1.3, H, "right")
    d.logo(s, M - 0.08, 0.55, 1.25)
    label(s, M, 2.25, 6, "Red Flag Homes Network  ·  The Drop")
    H1(s, M, 2.5, 6.4, 1.6, "The Drop", size=104, leading=0.85)
    H1(s, M, 3.95, 6, 0.8, "Studio · Noir", size=40, italic=True, color="caramel")
    T(s, M, 4.95, 5.6, 0.8, "One room. One delivery.\nGuest-ready in fifteen days.", size=15, color="ink",
      leading=1.35)
    d.rule(s, M, 6.3, 5.4, "caramel")
    label(s, M, 6.45, 6, "85+ pieces  ·  delivered  ·  installed  ·  photographed", color="brass", size=7.5)

    # 2 — What this is (oatmeal) --------------------------------------------------------------------
    s = d.slide(light=True, transition=("morph", False))
    d.photo(s, 8.4, 0, W - 8.4, H, "dark_bedroom", anchor=0.5)
    label(s, M, 1.15, 6, "What this is")
    H1(s, M, 1.45, 7.4, 1.7, "Fifteen days after you book, your studio looks like a boutique hotel room.",
       size=33, leading=1.0)
    T(s, M, 3.3, 7.0, 0.8, "You have a studio with an AC, a fan, a wardrobe and white tubelights. "
      "We turn it into a room that is ready to take bookings.", size=12.5, color="body", leading=1.45)
    for i, (big, small) in enumerate((("One", "order"), ("One", "delivery"), ("One", "install day"))):
        x = M + i * 2.4
        d.rule(s, x, 4.3, 2.1, "ink")
        H1(s, x, 4.42, 2.2, 0.6, big, size=34, italic=True, color="caramel")
        T(s, x, 5.05, 2.2, 0.3, small, size=11, color="ink", bold=True)
    T(s, M, 5.65, 7.2, 0.6, "We source everything, bring it, assemble it, style it, deep clean it and photograph it. "
      "You unlock the door once and we hand it back finished.", size=11, color="body", leading=1.4)
    label(s, M, 6.55, 7.5, "No market trips  ·  No fifteen deliveries  ·  No carpenter who doesn’t turn up",
          color="red", size=7)

    # 3 — The look ---------------------------------------------------------------------------------
    s = d.slide()
    label(s, M, 1.1, 6, "The look")
    H1(s, M, 1.4, 6.4, 1.2, "Noir is warm,\nnot cold.", size=48, leading=0.92)
    sw = [("charcoal", "Charcoal"), ("wood", "Warm wood"), ("cream", "Cream & oatmeal"), ("black", "Black metal"),
          ("brass", "Brass"), ("green", "Green")]
    for i, (c, n) in enumerate(sw):
        x = M + i * 1.0
        o = oval(s, x, 3.05, 0.72, 0.72, ACC[c], line=NOIR["hair"], C=d.C)
        T(s, x - 0.1, 3.88, 0.95, 0.4, n, size=7.5, color="muted", align="center", leading=1.1)
    d.bullets(s, M, 4.6, 5.8, 2.2, ["Deep charcoal surfaces in matte — never gloss",
                                    "Jute and wood texture to keep the room soft",
                                    "Cream and oatmeal bedding layered over dark furniture",
                                    "Black-framed art, terracotta pots, real plants, brass accents"],
              size=11, gap=5)
    d.photo(s, 7.3, 1.1, 2.9, 5.6, "rattan", anchor=0.4)
    d.photo(s, 10.35, 1.1, W - M - 10.35, 2.72, "candle", anchor=0.5)
    d.photo(s, 10.35, 3.98, W - M - 10.35, 2.72, "caramel_bed", anchor=0.5)

    # 4 — The light ----------------------------------------------------------------------------------
    s = d.slide(transition=("fade", True))
    d.photo(s, 6.6, 0, W - 6.6, H, "brass_pendants", anchor=0.3)
    d.mask(s, 6.6, 0, 1.4, H, "right")
    label(s, M, 1.1, 6, "The single biggest change")
    H1(s, M, 1.4, 5.8, 1.0, "It’s the light.", size=52)
    H1(s, M, 2.55, 5.8, 1.6, "2700K", size=110, color="brass", leading=0.85)
    T(s, M, 4.3, 5.4, 0.3, "warm, in every fitting we install", size=11, color="ink", bold=True)
    d.rule(s, M, 4.75, 5.3)
    T(s, M, 4.95, 5.5, 1.6, "White tubelights make any room photograph like a clinic. We light from lamps, sconces "
      "and a pendant instead of one ceiling tube — the difference between a dark room and an expensive one.",
      size=12, color="body", leading=1.45)

    # 5 — Sample gallery ---------------------------------------------------------------------------
    s = d.slide()
    label(s, M, 1.1, 8, "Sample images  ·  How your Airbnb could look")
    H1(s, M, 1.4, 10, 0.9, "Picture your studio like this.", size=40)
    gw, gy, gh = (W - 2 * M - 0.5) / 3, 2.4, 4.1
    for i, (img, cap, anc) in enumerate((("sample_studio", "Living, sleeping, dining", 0.45),
                                         ("sample_plan", "The whole studio, planned", 0.5),
                                         ("sample_kitchen", "Kitchenette and bed", 0.5))):
        x = M + i * (gw + 0.25)
        d.photo(s, x, gy, gw, gh, img, anchor=anc, tag=SAMPLE)
        T(s, x, gy + gh + 0.12, gw, 0.3, cap, size=15, font=SERIF, color="ink", italic=True)

    # 6–9 — Everything included (oatmeal) -------------------------------------------------------------
    cats = {
        "Sleep": ("black_bed", [("Queen bed frame with box storage, dark finish", "78 × 60 in"),
                                ("Medium-firm mattress", "78 × 60 × 6 in"),
                                ("Mattress protectors", "78 × 60 in"),
                                ("Pillows — soft and firm, with protectors", "18 × 28 in"),
                                ("Full linen sets — fitted, flat, pillow covers, 300TC", "Flat 90 × 108 in"),
                                ("Knit throw and waffle throw", "50 × 60 in"),
                                ("Cushion covers with inserts", "45 × 45 cm"),
                                ("Black metal bedside tables", "40 × 40 × 55 cm")]),
        "Lighting": ("rattan", [("Rattan or smoked-glass pendant over the bed", "40 cm dia"),
                                ("Black metal wall sconces with brass inner", "15 × 25 cm"),
                                ("Table lamp with amber shade", "35 cm high"),
                                ("Slim black floor lamp", "150 cm high"),
                                ("Under-cabinet warm LED strip", "1 m"),
                                ("Warm 2700K bulbs throughout, plus a night light", "2700K")]),
        "Furniture": ("dark_chairs", [("Console or open shelf unit, black metal and wood", "100 × 35 × 75 cm"),
                                      ("Accent chair, cane or wishbone style", "55 × 55 × 80 cm"),
                                      ("Round side table, dark wood", "45 cm dia"),
                                      ("Full-length mirror, black frame", "150 × 40 cm")]),
        "Floor and windows": (None, [("Jute rug under the bed", "6 × 4 ft"),
                                     ("Blackout curtains, oatmeal, with black rod", "4.5 × 9 ft panels")]),
        "Bathroom": ("dark_bath", [("Bath towels", "75 × 150 cm"),
                                   ("Hand towels and face towels", "40 × 60 · 30 × 30 cm"),
                                   ("Bath mats", "50 × 80 cm"),
                                   ("Refillable dispensers — shampoo, conditioner, body wash, hand wash", "300 ml"),
                                   ("Mirror, shelf, towel rail and hooks", "60 × 80 cm mirror"),
                                   ("Bin, toilet brush, bucket and mug", "—"),
                                   ("Hair dryer", "1200 W"),
                                   ("Upgraded shower head", "—")]),
        "Kitchenette": ("sample_kitchen", [("Induction hob", "2000 W"),
                                           ("Electric kettle", "1.5 L"),
                                           ("Crockery, cutlery and glassware for two", "—"),
                                           ("Pan and pot", "24 cm · 2 L"),
                                           ("Chopping board and knife", "30 × 20 cm"),
                                           ("Storage jars, tea and coffee station on a brass tray", "35 cm tray"),
                                           ("Dish rack and bin", "—")]),
        "Safety, utility and access": ("green_wall", [
            ("Keyless smart lock — fingerprint, PIN and key backup", "Fits 35–60 mm doors"),
            ("Water purifier", "7 L RO"),
            ("Smoke alarm, fire extinguisher, fire blanket, first-aid kit", "2 kg ABC"),
            ("Iron and board, drying rack, laundry basket", "—"),
            ("Hangers, extension boards, cleaning set", "4-socket")]),
        "Styling and welcome": ("candle2", [("Framed prints, black frames", "A2 · 42 × 59 cm"),
                                            ("Plants in pots, plus a woven basket planter", "60–120 cm"),
                                            ("Tray, vase, books, candle and coasters, styled on the day", "—"),
                                            ("Scent diffuser with first refill", "200 ml"),
                                            ("Welcome basket and a printed house guide", "A5, bound")]),
    }
    pages = [(["Sleep"], ["Lighting"]), (["Furniture", "Floor and windows"], ["Bathroom"]),
             (["Kitchenette"], ["Safety, utility and access"]), (["Styling and welcome"], None)]
    cw = (W - 2 * M - 0.6) / 2
    for pi, (left, right) in enumerate(pages):
        s = d.slide(light=True)
        label(s, M, 1.0, 8, f"Everything included, with sizes  ·  {pi + 1} of 4")
        for ci, group in enumerate((left, right)):
            x = M + ci * (cw + 0.6)
            if group is None:  # last page: mood image + note on what is left out on purpose
                d.photo(s, x, 1.45, cw, 3.55, "green_sofa", anchor=0.5)
                d.box(s, x, 5.2, cw, 1.55, "paper")
                H1(s, x + 0.3, 5.35, cw - 0.6, 0.4, "Fridge and microwave?", size=20)
                T(s, x + 0.3, 5.8, cw - 0.6, 0.9, "Not included as standard — most studios don’t need them, and "
                  "it keeps your price down. Add either as an add-on if you want them.", size=10.5,
                  color="body", leading=1.4)
                continue
            y = 1.45
            img = cats[group[0]][0]
            d.photo(s, x, y, cw, 1.15, img, anchor=0.72 if img == "black_bed" else 0.5, tag=SAMPLE if img.startswith("sample") else None)
            y += 1.35
            for g in group:
                H1(s, x, y, cw, 0.45, g, size=24)
                y += 0.52
                for item, size in cats[g][1]:
                    d.rule(s, x, y, cw)
                    T(s, x, y + 0.08, cw - 1.9, 0.3, item, size=10, color="ink", leading=1.1)
                    T(s, x + cw - 1.85, y + 0.08, 1.85, 0.3, size, size=10, color="caramel", bold=True,
                      align="right")
                    y += 0.36 if len(item) < 52 else 0.5
                d.rule(s, x, y, cw)
                y += 0.25

    # 10 — The part nobody else does ---------------------------------------------------------------
    s = d.slide(transition=("fade", True))
    d.photo(s, 0, 0, 5.2, H, "sample_studio", anchor=0.55, tag=SAMPLE)
    x0 = 5.85
    label(s, x0, 1.1, 6, "And the part nobody else does")
    H1(s, x0, 1.4, 6.8, 0.9, "Not a kit. A finished room.", size=38)
    svc = [("Design plan", "A full layout before anything is ordered, so the room works as a whole"),
           ("Sourcing and delivery", "One order, one contact — not fifteen sellers"),
           ("Full install day", "Assembled, hung, placed and styled by our team"),
           ("Deep clean", "Guest-ready, not builder-dusty"),
           ("Professional shoot", "20+ edited images, ready to upload the same evening")]
    for i, (k, v) in enumerate(svc):
        y = 2.5 + i * 0.66
        d.rule(s, x0, y, W - M - x0)
        H1(s, x0, y + 0.12, 2.6, 0.4, k, size=19, color="ink")
        T(s, x0 + 2.75, y + 0.18, W - M - x0 - 2.75, 0.45, v, size=10.5, color="body", leading=1.3)
    d.rule(s, x0, 2.5 + 5 * 0.66, W - M - x0)
    H1(s, x0, 6.0, 6.8, 0.6, "Photographs are 80% of the booking decision. A kit without a shoot is half a job.",
       size=17, italic=True, color="brass", leading=1.05)

    # 11 — Two versions / price -------------------------------------------------------------------
    s = d.slide(transition=("fade", True))
    label(s, M, 1.0, 6, "Two versions")
    H1(s, M, 1.28, 11, 0.9, "Choose your Noir.", size=40)
    feats = ["Everything listed in this brochure", "Charcoal matte paint, two walls", "Warm wood vinyl plank flooring",
             "Wardrobe refinished in matte black"]
    tiers = [("Core", "₹3,21,538", "₹2,09,000", [True, False, False, False],
              "If your floor and walls are already presentable.", "paper", "caramel_bed"),
             ("Full Noir", "₹4,29,231", "₹2,79,000", [True, True, True, True],
              "If the room still looks like a builder handed it over.", "paper", "dark_bedroom")]
    tw = (W - 2 * M - 0.4) / 2
    for i, (name, was, now, inc, when, fill, img) in enumerate(tiers):
        x = M + i * (tw + 0.4)
        top = 2.3
        d.box(s, x, top, tw, 4.4, fill, line="brass" if i == 1 else None)
        d.photo(s, x + tw - 1.6, top, 1.6, 4.4, img, anchor=0.5)
        badge = d.box(s, x + 0.35, top + 0.35, 1.15, 0.34, "red", round_=True, radius=0.5)
        T(s, x + 0.35, top + 0.41, 1.15, 0.25, "35% OFF", size=8.5, color="white", bold=True, align="center",
          spacing=150)
        d.box(s, x + 1.6, top + 0.35, 1.95, 0.34, "brass", round_=True, radius=0.5)
        T(s, x + 1.6, top + 0.41, 1.95, 0.25, "+ FREE RED FLAG GIFT", size=8, color="black", bold=True,
          align="center", spacing=100)
        H1(s, x + 0.35, top + 0.82, 3.5, 0.6, name, size=30, italic=i == 1, color="brass" if i == 1 else "ink")
        T(s, x + 0.35, top + 1.45, 3.4, 0.35, was, size=15, color="muted", strike=True)
        H1(s, x + 0.35, top + 1.72, 3.6, 0.8, now, size=46, color="ink", leading=0.9)
        T(s, x + 0.35, top + 2.5, 3.4, 0.25, "plus GST", size=8.5, color="muted")
        for j, f in enumerate(feats):
            yy = top + 2.9 + j * 0.3
            T(s, x + 0.35, yy, 0.25, 0.25, "✓" if inc[j] else "—", size=10, color="brass" if inc[j] else "muted",
              bold=True)
            T(s, x + 0.62, yy, 3.3, 0.25, f, size=9.5, color="ink" if inc[j] else "muted")
        T(s, x + 0.35, top + 4.02, 3.6, 0.3, when, size=9, color="caramel", italic=True)
    T(s, M, 6.78, 11.8, 0.25, "Plus applicable GST. For studios up to 350 sq ft. Delivery and installation included in Goa, "
      "Mumbai, Bengaluru, Delhi NCR and Lucknow; outside these, at cost.", size=7.5, color="muted")

    # 12 — The gift ------------------------------------------------------------------------------------
    s = d.slide(transition=("fade", True))
    px, py, pw, ph = M + 0.2, 2.2, 5.4, 2.6          # the brass name plate, drawn natively
    d.box(s, M, 1.1, 5.8, 5.6, "paper")
    plate = d.box(s, px, py, pw, ph, "brass", line="#E2C48E", round_=True, radius=0.05)
    plate.name = "static-plate"
    d.box(s, px + 0.18, py + 0.18, pw - 0.36, ph - 0.36, "brass", line="#8E6C3A", round_=True, radius=0.04).name = "static-plate"
    for sx, sy in ((px + 0.35, py + 0.35), (px + pw - 0.5, py + 0.35), (px + 0.35, py + ph - 0.5), (px + pw - 0.5, py + ph - 0.5)):
        oval(s, sx, sy, 0.15, 0.15, "#7A5C2E", C=d.C).name = "static-screw"
    T(s, px, py + 0.75, pw, 0.4, "RED FLAG  ×", size=15, color="black", bold=True, align="center", spacing=500)
    H1(s, px, py + 1.2, pw, 0.8, "Your Property", size=40, color="black", align="center", italic=True)
    T(s, M, 5.3, 5.8, 0.3, "Branded brass name plate", size=10, color="muted", align="center", italic=True)
    x0 = 7.2
    label(s, x0, 1.1, 6, "Free with both packages", color="brass")
    H1(s, x0, 1.4, W - M - x0, 1.4, "A welcome gift\nfrom Red Flag Homes.", size=38, leading=0.98)
    gifts = [("Branded brass name plate", "Engraved, for your front door"),
             ("Branded tissues", "For the room and the bathroom"),
             ("Branded slippers", "Waiting by the bed for every guest"),
             ("Branded glass water bottles", "On the bedside, ready for check-in")]
    for i, (k, v) in enumerate(gifts):
        y = 3.05 + i * 0.85
        d.rule(s, x0, y, W - M - x0)
        T(s, x0, y + 0.16, 0.6, 0.4, f"0{i + 1}", size=12, color="brass", bold=True)
        H1(s, x0 + 0.7, y + 0.1, W - M - x0 - 0.7, 0.4, k, size=20, color="ink")
        T(s, x0 + 0.7, y + 0.5, W - M - x0 - 0.7, 0.3, v, size=10, color="muted")
    d.rule(s, x0, 3.05 + 4 * 0.85, W - M - x0)
    T(s, x0, 6.62, W - M - x0, 0.3, "Included free with Core and Full Noir.", size=10.5, color="caramel", bold=True)

    # 12 — Add what you need (oatmeal) ------------------------------------------------------------------
    s = d.slide(light=True)
    d.photo(s, 8.9, 0, W - 8.9, H, "candle2", anchor=0.5)
    label(s, M, 1.1, 6, "Add what you need")
    H1(s, M, 1.4, 7.5, 0.9, "Make it yours.", size=40)
    adds = [("Mini fridge, 45–50 L", "₹14,000"), ("Microwave, 20 L", "₹10,000"),
            ("WiFi router with guest network, set up", "₹4,500"),
            ("Smart TV 32\" with streaming stick, wall-mounted", "₹22,000"),
            ("Round table with chair — eating or working", "₹9,500"),
            ("Branded amenity kit — tissue, toiletries, coasters, welcome card", "₹8,000"),
            ("Reels package — vertical video for Instagram and your listing", "₹15,000"),
            ("Restock subscription — linen, toiletries, consumables", "₹2,500 / month")]
    for i, (k, v) in enumerate(adds):
        y = 2.5 + i * 0.5
        d.rule(s, M, y, 7.4)
        T(s, M, y + 0.14, 5.7, 0.3, k, size=11, color="ink")
        H1(s, M + 5.6, y + 0.06, 1.8, 0.35, v, size=17, color="caramel", align="right")
    d.rule(s, M, 2.5 + 8 * 0.5, 7.4)
    T(s, M, 6.7, 7.4, 0.25, "Add-ons plus applicable GST. Installed on the same day.", size=8, color="muted")

    # 13 — Not included ---------------------------------------------------------------------------
    s = d.slide()
    label(s, M, 1.1, 6, "What’s not included")
    H1(s, M, 1.4, 6, 1.4, "Said plainly,\nup front.", size=46, leading=0.95)
    T(s, M, 3.0, 4.6, 1.0, "We work with what you already have — and quote anything bigger separately.",
      size=12, color="body", leading=1.45)
    notinc = ["Structural, plumbing or electrical work, false ceilings, rewiring",
              "Air conditioner, fan and wardrobe — we work with what you already have",
              "Washing machine and dishwasher", "Additional or sofa beds — quoted separately",
              "Studios above 350 sq ft — quoted separately", "Your society or building permissions"]
    for i, t in enumerate(notinc):
        y = 1.45 + i * 0.8
        d.rule(s, 6.3, y, W - M - 6.3)
        T(s, 6.3, y + 0.22, 0.4, 0.3, "✕", size=11, color="caramel", bold=True)
        T(s, 6.75, y + 0.22, W - M - 6.75, 0.5, t, size=12.5, color="ink", leading=1.3)
    d.rule(s, 6.3, 1.45 + 6 * 0.8, W - M - 6.3)

    # 14 — How it works ------------------------------------------------------------------------------
    s = d.slide(light=True)
    label(s, M, 1.1, 6, "How it works")
    H1(s, M, 1.4, 11, 0.9, "Fifteen days from booking to a bookable room.", size=38)
    steps = [("Day 0", "Send photos and measurements. We confirm the price within 24 hours.", "10% to book"),
             ("Day 1", "Approve the design plan with the actual items. Sourcing starts.", "30%"),
             ("Day 2–12", "Everything is made and collected at our end. Nothing lands at your door piecemeal.",
              "30% by day 7"),
             ("Day 13", "Everything is packed and dispatched to your studio.", "30% before delivery"),
             ("Day 14", "Install day. We assemble, place, hang and style the entire room.", None),
             ("Day 15", "Deep clean and professional shoot. Images sent that evening.", "Nothing to pay")]
    sw_ = (W - 2 * M) / 6
    d.rule(s, M, 2.85, W - 2 * M, "ink", 0.016)
    for i, (dday, txt, pay) in enumerate(steps):
        x = M + i * sw_
        oval(s, x, 2.75, 0.22, 0.22, ACC["red"] if i == 5 else ACC["caramel"], C=d.C)
        T(s, x, 3.25, sw_ - 0.15, 0.5, dday, size=21, color="ink", bold=True)
        T(s, x, 3.9, sw_ - 0.3, 1.5, txt, size=10, color="body", leading=1.45)
        if pay:
            fill = "ink" if i == 0 else ("caramel" if pay.startswith("30") else "paper")
            d.box(s, x, 5.35, sw_ - 0.25, 0.36, fill, round_=True, radius=0.5)
            T(s, x, 5.42, sw_ - 0.25, 0.25, pay, size=8.5, color="bg" if fill == "ink" else
              ("white" if fill == "caramel" else "ink"), bold=True, align="center")
    d.box(s, M, 6.1, W - 2 * M, 0.55, "ink")
    T(s, M + 0.3, 6.25, 11.5, 0.3, "10% to book  ·  the balance in three equal parts, all before delivery  ·  "
      "nothing to pay on handover day", size=10.5, color="bg", bold=True)

    # 15 — Questions --------------------------------------------------------------------------------
    s = d.slide()
    label(s, M, 1.0, 6, "Questions")
    H1(s, M, 1.28, 10, 0.8, "Everything owners ask us.", size=36)
    faq = [("Why no fridge or microwave?", "Most studio guests don’t cook, and leaving them out keeps your price lower. Add them and we install them the same day."),
           ("Can I change items?", "Yes, within the palette. Swap the chair, choose a different pendant, change the art. We keep the room coherent."),
           ("Can I see it before you order?", "Yes. You approve the design plan with the actual items before anything is bought."),
           ("What if something arrives damaged?", "We replace it. Every appliance carries its manufacturer warranty, registered in your name."),
           ("Will it fit my studio?", "Send measurements with your photos. If the room is tight we adjust the layout, or suggest a smaller bed."),
           ("Do you do 1BHK and villas?", "Yes. The Drop comes in Studio, Villa and Estate. Tell us the property and we’ll quote it."),
           ("I want you to run it too.", "That’s Red Flag Outpost Classic — we design, brand, launch and operate it for 8% of bookings. Join within 90 days and we deduct what you spent on The Drop from the Outpost setup fee.")]
    fw = (W - 2 * M - 0.9) / 4
    for i, (q, a) in enumerate(faq):
        x = M + (i % 4) * (fw + 0.3)
        y = 2.3 + (i // 4) * 2.3
        d.rule(s, x, y, fw, "caramel")
        H1(s, x, y + 0.12, fw, 0.6, q, size=17, color="ink", leading=1.0)
        T(s, x, y + 0.78, fw, 1.4, a, size=9.5, color="body", leading=1.4)

    # 16 — Order yours --------------------------------------------------------------------------------
    s = d.slide(transition=("fade", True))
    d.photo(s, 7.6, 0, W - 7.6, H, "sample_kitchen", anchor=0.4, tag=SAMPLE)
    d.mask(s, 7.6, 0, 1.3, H, "right")
    label(s, M, 1.1, 6, "Order yours")
    H1(s, M, 1.4, 6.6, 2.2, "A price in 24 hours.\nA finished room\nin fifteen days.", size=46, leading=0.98)
    T(s, M, 3.95, 6.2, 0.8, "Send photos of the studio and its measurements. That’s all we need to start.",
      size=13, color="body", leading=1.45)
    for i, (now, was, n) in enumerate((("₹2,09,000", "₹3,21,538", "Core"), ("₹2,79,000", "₹4,29,231", "Full Noir"))):
        x = M + i * 3.2
        label(s, x, 4.9, 3, n, color="brass")
        T(s, x, 5.18, 2.9, 0.3, was, size=11, color="muted", strike=True)
        H1(s, x, 5.42, 3.0, 0.6, now, size=32, color="ink")
    T(s, M, 6.2, 6.4, 0.3, "35% off  ·  free Red Flag gift  ·  10% to book  ·  plus GST", size=9, color="red", bold=True, spacing=150)

    # 17 — Contact ----------------------------------------------------------------------------------
    s = d.slide(chrome=False, transition=("morph", False))
    d.photo(s, 0, 0, W, H, "dark_bedroom", anchor=0.5)
    d.shade(s, 0, 0, W, H, 62)
    d.logo(s, W / 2 - 0.7, 1.0, 1.4)
    H1(s, 1.5, 2.75, W - 3, 1.0, "redflaghomes.in", size=64, align="center")
    T(s, 1.5, 3.85, W - 3, 0.35, "hello@redflaghomes.in", size=15, color="ink", bold=True, align="center")
    T(s, 1.5, 4.3, W - 3, 0.3, "The Drop by Red Flag  ·  Studio  ·  Villa  ·  Estate", size=10, color="brass",
      align="center", spacing=150)
    T(s, 1.5, 4.65, W - 3, 0.3, "Red Flag Homes Network · a company of Red Flag World Holdings Ltd, London",
      size=9, color="body", align="center")
    T(s, 1.8, 6.15, W - 3.6, 0.8, "Prices plus applicable GST, for studios up to 350 sq ft in our delivery cities. Sizes are "
      "standard and may vary slightly by supplier. Specific items may be substituted for equivalents of the same quality "
      "and palette where stock requires it — the design plan you approve is what we deliver. Images marked “Sample "
      "image” are for illustration.", size=7.5, color="body", align="center", leading=1.35)

    assert len(d.prs.slides) == TOTAL, len(d.prs.slides)
    d.save()


if __name__ == "__main__":
    build()
