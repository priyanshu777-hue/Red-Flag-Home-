"""Build the Outpost Classic Partner Guide deck — Red Flag Homes Network.

Build Mode with the pptx-designer public API. Theme lock: theme-lock.yaml v1 ("Sunlit hospitality").
Fonts: Fraunces + Manrope (see fonts/). Run: python build_deck.py
"""
from pathlib import Path

from pptx.util import Inches

from pptx_designer import Presentation
from pptx_designer.tools.images import gradient_mask_image
from pptx_designer.tools.layout import add_slide, clean_save
from pptx_designer.tools.shapes import oval, rect, rrect
from pptx_designer.tools.text import text
from pptx.oxml.ns import qn

import motion

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
IMG = HERE / "assets"
OUT = HERE / "output"
LOGO = str(ROOT / "logobg.png")

SERIF, SANS = "Fraunces", "Manrope"
W, H, M = 13.333, 7.5, 0.75
TOTAL = 24

LIGHT = {"bg": "#F6F1EA", "paper": "#ECE4D8", "ink": "#1C1A17", "body": "#4A443D", "muted": "#8A8074",
         "hair": "#D9CFC1"}
DARK = {"bg": "#161311", "paper": "#221E1A", "ink": "#F3ECE1", "body": "#D6CCBE", "muted": "#A89E91",
        "hair": "#3B352E"}
ACC = {"brass": "#A8834A", "brass_lt": "#C9A66B", "red": "#D93A2F", "white": "#FFFFFF"}


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.pal = LIGHT
        self.plan = {}

    # ---- colour + primitives ------------------------------------------------------------
    def col(self, key):
        return self.pal.get(key) or ACC[key]

    @property
    def C(self):
        return {"background": self.pal["bg"], "text_body": self.pal["body"], "text_dark": self.pal["ink"],
                "text_muted": self.pal["muted"], "font_body": SANS, "font_heading": SERIF}

    def T(self, s, x, y, w, h, txt, size=12, font=SANS, color="body", italic=False, bold=False,
          spacing=None, align="left", leading=None, caps=False, anchor="top"):
        box = text(s, x, y, w, h, txt.upper() if caps else txt, font_size=size, color=self.col(color),
                   bold=bold, align=align, font_name=font, C=self.C, anchor=anchor)
        tf = box.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for p in tf.paragraphs:
            if leading:
                p.line_spacing = leading
            for r in p.runs:
                r.font.italic = italic
                if spacing is not None:
                    r.font._element.set("spc", str(spacing))
        return box

    def H1(self, s, x, y, w, h, txt, size=40, color="ink", italic=False, align="left", leading=0.95):
        return self.T(s, x, y, w, h, txt, size=size, font=SERIF, color=color, italic=italic, align=align,
                      leading=leading)

    def label(self, s, x, y, w, txt, color="brass", size=8.5, align="left"):
        return self.T(s, x, y, w, 0.25, txt, size=size, color=color, bold=True, spacing=250, caps=True,
                      align=align)

    def bullets(self, s, x, y, w, h, items, size=11, color="body", gap=6, mark="—", mark_color="brass"):
        box = self.T(s, x, y, w, h, "", size=size, color=color)
        tf = box.text_frame
        from pptx.util import Pt
        from pptx.dml.color import RGBColor
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(gap)
            p.line_spacing = 1.2
            if mark:
                r0 = p.add_run()
                r0.text = f"{mark}  "
                r0.font.size, r0.font.name = Pt(size), SANS
                r0.font.color.rgb = RGBColor.from_string(self.col(mark_color)[1:])
                r0.font.bold = True
            r = p.add_run()
            r.text = it
            r.font.size, r.font.name = Pt(size), SANS
            r.font.color.rgb = RGBColor.from_string(self.col(color)[1:])
        return box

    def box(self, s, x, y, w, h, fill, line=None, round_=False):
        f = (rrect if round_ else rect)(s, x, y, w, h, self.col(fill), line=self.col(line) if line else None,
                                         C=self.C)
        if round_:
            f.adjustments[0] = 0.08
        return f

    def rule(self, s, x, y, w, color="hair", t=0.012):
        return self.box(s, x, y, w, t, color)

    def vrule(self, s, x, y, h, color="hair"):
        return self.box(s, x, y, 0.012, h, color)

    def photo(self, s, x, y, w, h, name):
        """Cover-fit like pptx_designer's cover_image, but embedded as JPEG.

        cover_image writes each crop as a lossless PNG (this deck came out at 44 MB) and caches crops by
        path only, so an edited asset can be served stale. Same centred crop, sized for 16:9 projection.
        """
        from io import BytesIO
        from PIL import Image
        img = Image.open(IMG / f"{name}.jpg").convert("RGB")
        iw, ih = img.size
        r = w / h
        cw, ch = (int(ih * r), ih) if iw / ih > r else (iw, int(iw / r))
        img = img.crop(((iw - cw) // 2, (ih - ch) // 2, (iw - cw) // 2 + cw, (ih - ch) // 2 + ch))
        px = min(cw, int(w * 180))  # ~180 dpi at slide size
        img = img.resize((px, int(px / r)), Image.LANCZOS)
        buf = BytesIO()
        img.save(buf, "JPEG", quality=86, optimize=True, progressive=True)
        buf.seek(0)
        pic = s.shapes.add_picture(buf, Inches(x), Inches(y), Inches(w), Inches(h))
        pic.name = "static-photo"
        return pic

    def mask(self, s, x, y, w, h, direction, a0=100, a1=0, color=None):
        shp = gradient_mask_image(s, x, y, w, h, bg_color=color or self.pal["bg"], direction=direction,
                                  alpha_start=a0, alpha_end=a1)
        shp.name = "static-mask"
        return shp

    def veil(self, s, pct, color=None):
        shp = rect(s, 0, 0, W, H, color or self.pal["bg"], C=self.C)
        shp.name = "static-veil"
        clr = shp.fill._xPr.find(qn("a:solidFill"))[0]
        clr.append(clr.makeelement(qn("a:alpha"), {"val": str(pct * 1000)}))
        return shp

    def logo(self, s, x, y, size):
        pic = s.shapes.add_picture(LOGO, Inches(x), Inches(y), Inches(size), Inches(size))
        pic.name = "!!logo"
        return pic

    # ---- slide scaffolding ---------------------------------------------------------------
    def slide(self, dark=False, chrome=True, transition=("fade", False)):
        self.pal = DARK if dark else LIGHT
        s = add_slide(self.prs)
        n = len(self.prs.slides)
        self.plan[n] = transition
        rect(s, 0, 0, W, H, self.pal["bg"], C=self.C).name = "static-bg"
        if chrome:
            self.logo(s, W - M - 0.55, 0.32, 0.55)
            for shp in (
                self.label(s, M, 0.52, 6, "Outpost Classic  ·  Partner Guide", color="muted", size=7.5),
                self.label(s, M, H - 0.5, 4, "redflaghomes.in", color="muted", size=7.5),
                self.T(s, W - M - 1.5, H - 0.5, 1.5, 0.25, f"{n:02d} / {TOTAL:02d}", size=7.5,
                       color="muted", bold=True, spacing=200, align="right"),
            ):
                shp.name = "chrome"
        return s

    def header(self, s, kicker, title, y=1.1, w=9.5, size=36, sub=None, sub_w=8.5):
        self.label(s, M, y, 8, kicker)
        self.H1(s, M, y + 0.32, w, 1.0, title, size=size)
        if sub:
            self.T(s, M, y + 0.32 + size / 72 * 1.25 + 0.12, sub_w, 0.6, sub, size=13, color="body",
                   leading=1.3)

    def save(self):
        for sl in self.prs.slides:  # keep the logo mark above photos (and visible for Morph)
            tree = sl.shapes._spTree
            for shp in [x for x in sl.shapes if x.name == "!!logo"]:
                tree.remove(shp._element)
                tree.append(shp._element)
        for i, sl in enumerate(self.prs.slides, start=1):
            kind, black = self.plan.get(i, ("fade", False))
            motion.transition(sl, kind, through_black=black)
            motion.choreograph(sl)
        OUT.mkdir(exist_ok=True)
        path = OUT / "Outpost_Classic_Partner_Guide.pptx"
        clean_save(self.prs, str(path))
        print(path)


def build():
    d = Deck()
    T, H1, label = d.T, d.H1, d.label

    # 1 — Cover ------------------------------------------------------------------------------
    s = d.slide(chrome=False, transition=("fade", True))
    d.photo(s, 5.3, 0, W - 5.3, H, "aerial_pool")
    d.mask(s, 5.3, 0, 1.2, H, "right")
    d.logo(s, M - 0.08, 0.55, 1.3)
    label(s, M, 2.3, 4.5, "Red Flag Homes Network  ·  Partner Guide", color="brass")
    H1(s, M, 2.55, 4.6, 2.4, "Outpost\nClassic", size=80, leading=0.84)
    H1(s, M, 5.15, 4.4, 1.0, "You bring the property.\nWe build it, brand it and run it.", size=21,
       italic=True, color="body", leading=1.1)
    d.rule(s, M, 6.45, 4.1)
    label(s, M, 6.6, 4.75, "8% commission  ·  200 founding allocations  ·  India", color="red", size=7.5)
    card = d.box(s, 8.85, 5.55, 3.75, 1.3, "bg")
    T(s, 9.1, 5.72, 3.3, 0.3, "The Earn-Back Promise", size=8, color="brass", bold=True, spacing=200, caps=True)
    H1(s, 9.1, 6.0, 3.35, 0.8, "Earn back your programme fee in 12 months — or we pay you the difference.",
       size=13.5, leading=1.1)

    # 2 — The model ---------------------------------------------------------------------------
    s = d.slide(transition=("morph", False))
    d.photo(s, 7.55, 0, W - 7.55, H, "villa_pool")
    label(s, M, 1.2, 6, "The model, in one line")
    H1(s, M, 1.55, 6.6, 1.6, "You own or lease the property.\nWe do everything else.", size=30, leading=1.05)
    T(s, M, 2.95, 6.1, 1.2,
      "Red Flag designs it, brands it, launches it through creators and runs it completely — pricing, "
      "guests, cleaning, repairs. You pay one programme fee and 8% of what it books.", size=13.5, leading=1.4)
    steps = [("01", "Design"), ("02", "Brand"), ("03", "Launch"), ("04", "Run")]
    for i, (n, t) in enumerate(steps):
        x = M + i * 1.58
        d.rule(s, x, 5.3, 1.4, "ink")
        T(s, x, 5.45, 1.4, 0.3, n, size=9, color="brass", bold=True, spacing=150)
        H1(s, x, 5.72, 1.5, 0.5, t, size=22)

    # 3 — Why now (dark) ------------------------------------------------------------------------
    s = d.slide(dark=True, transition=("fade", True))
    label(s, M, 1.2, 6, "Why now")
    H1(s, M, 1.55, 11, 2.0, "Earn it back in 12 months,\nor we pay the difference.", size=52, leading=0.98)
    stats = [("200", "founding allocations", "across India, reviewed in the order they arrive."),
             ("8%", "for life", "for founding partners. After the 200, the rate rises."),
             ("1 Oct", "2026", "is when allocations open.")]
    for i, (big, k, v) in enumerate(stats):
        x = M + i * 3.95
        d.rule(s, x, 4.25, 3.55, "brass")
        H1(s, x, 4.45, 3.6, 1.1, big, size=60, color="brass_lt", leading=0.9)
        T(s, x, 5.5, 3.5, 0.3, k, size=13, color="ink", bold=True)
        T(s, x, 5.82, 3.4, 0.7, v, size=11.5, color="muted", leading=1.3)

    # 4 — Two ways in ---------------------------------------------------------------------------
    s = d.slide()
    d.header(s, "Two ways in", "Two ways in. One system.", size=36,
             sub="Same fee. Same 8%. Same promise. The only difference is whether you pay a landlord.")
    xa, xb, cw = 4.55, 8.7, 3.95
    for x, img, t in ((xa, "house_dusk", "I own a property"), (xb, "living", "I don’t own one")):
        d.photo(s, x, 2.55, cw, 0.95, img)
        d.box(s, x, 3.2, cw, 0.3, "ink").name = "static-tag"
        T(s, x + 0.15, 3.25, cw - 0.3, 0.25, t, size=9, color="white", bold=True, spacing=150, caps=True)
    rows = [("Franchise fee, one time", "₹59,999", None),
            ("Setup & procuring, up to 2 keys", "₹1,11,111", None),
            ("Programme fee", "₹1,71,110", None),
            ("Rent and deposit", "None", "Paid by you to your landlord"),
            ("Electricity, water, internet", "You", None),
            ("Red Flag commission", "8%, min ₹4,999 a month", None),
            ("Cleaning & minor maintenance", "From booking revenue, at cost", None),
            ("Earn-Back Promise", "Yes", None)]
    y0, rh = 3.62, 0.355
    for i, (k, a, b) in enumerate(rows):
        y = y0 + i * rh
        d.rule(s, M, y, W - 2 * M)
        if b:
            d.box(s, xa, y + 0.012, W - M - xa, rh - 0.012, "paper")
        T(s, M, y + 0.09, 3.6, 0.3, k, size=11, color="ink", bold=True)
        if b:
            T(s, xa, y + 0.09, cw, 0.3, a, size=11, color="ink", align="center")
            T(s, xb, y + 0.09, cw, 0.3, b, size=11, color="red", bold=True, align="center")
        else:
            T(s, xa, y + 0.09, W - M - xa, 0.3, a, size=11, color="ink", align="center")
    d.rule(s, M, y0 + len(rows) * rh, W - 2 * M)
    T(s, M, 6.62, 11.5, 0.3, "All prices plus applicable GST. Setup covers up to 2 keys; each extra key ₹45,000. "
      "Furniture for an unfurnished property is quoted separately, itemised, at cost.", size=8.5, color="muted")

    # 5 — The price -----------------------------------------------------------------------------
    s = d.slide()
    d.header(s, "The price", "One fee. Then 8% of what it books.", size=36)
    eq = [("Franchise fee", "₹59,999", "one time"), ("Setup & procuring", "₹1,11,111", "up to 2 keys"),
          ("Programme fee", "₹1,71,110", "one time, plus GST")]
    for i, (k, v, sub) in enumerate(eq):
        y = 2.45 + i * 1.22
        if i:
            H1(s, M, y - 0.12, 0.5, 0.6, "+" if i == 1 else "=", size=30, color="brass")
        label(s, M + 0.6, y, 3, k, color="muted")
        H1(s, M + 0.6, y + 0.25, 4.2, 0.8, v, size=44 if i < 2 else 50, color="red" if i == 2 else "ink",
           leading=0.9)
        T(s, M + 4.4, y + 0.5, 2.0, 0.3, sub, size=10.5, color="muted")
        if i == 1:
            d.rule(s, M + 0.6, y + 1.1, 5.6, "ink", 0.02)
    px, pw = 7.35, W - M - 7.35
    d.box(s, px, 2.35, pw, 3.55, "paper")
    label(s, px + 0.35, 2.6, 4, "Setup Flex")
    H1(s, px + 0.35, 2.9, pw - 0.7, 0.6, "Start with less at signing.", size=22)
    tot = pw - 0.7
    a_w = tot * 115555 / 171110
    d.box(s, px + 0.35, 3.75, a_w, 0.42, "ink")
    d.box(s, px + 0.35 + a_w + 0.04, 3.75, tot - a_w - 0.04, 0.42, "brass")
    T(s, px + 0.5, 3.84, a_w - 0.2, 0.3, "₹1,15,555 at signing", size=10.5, color="white", bold=True)
    T(s, px + 0.35 + a_w + 0.15, 3.84, 1.7, 0.3, "₹55,556", size=10.5, color="white", bold=True)
    T(s, px + 0.35, 4.35, tot, 0.9, "The balance comes from your first six payouts — at no extra charge.",
      size=12, color="body", leading=1.35)
    d.rule(s, px + 0.35, 5.0, tot)
    label(s, px + 0.35, 5.15, 4, "Commission", color="muted")
    T(s, px + 0.35, 5.42, tot, 0.4, "8% of booking revenue, min ₹4,999 a month", size=13, color="ink", bold=True)
    d.box(s, M, 6.18, W - 2 * M, 0.62, "ink")
    T(s, M + 0.35, 6.33, 11, 0.4, "You pay nothing until your property passes assessment.  We look first, then you decide.",
      size=13, color="white", bold=True)

    # 6 — Earn-Back Promise (dark) ----------------------------------------------------------------
    s = d.slide(dark=True, transition=("fade", True))
    d.photo(s, 8.6, 0, W - 8.6, H, "bonfire")
    d.mask(s, 8.6, 0, 1.6, H, "right")
    label(s, M, 1.15, 6, "The Earn-Back Promise", color="brass_lt")
    H1(s, M, 1.5, 7.6, 1.4, "If it doesn’t earn back its fee in year one, we pay the difference.", size=34,
       leading=1.02)
    T(s, M, 2.95, 7.4, 0.9, "If your Outpost’s net operating profit over its first 12 full months after going live is "
      "less than the programme fee you paid, Red Flag pays you the difference.", size=12, color="body", leading=1.35)
    label(s, M, 3.95, 5, "Worked example", color="muted")
    bw = 7.3
    bars = [("Programme fee you paid", 171110, "₹1,71,110", "hair"), ("Your Outpost nets in 12 months", 120000,
                                                                    "₹1,20,000", "ink")]
    for i, (k, v, t, c) in enumerate(bars):
        y = 4.3 + i * 0.72
        T(s, M, y, 4, 0.25, k, size=9.5, color="muted")
        d.box(s, M, y + 0.26, bw * v / 171110, 0.3, "brass" if i == 0 else "ink")
        T(s, M + 0.12, y + 0.3, 2.5, 0.25, t, size=10, color="bg" if i else "white", bold=True)
    gx = M + bw * 120000 / 171110
    d.box(s, gx + 0.03, 5.02 + 0.26, bw - (gx - M) - 0.03, 0.3, "red")
    T(s, gx + 0.1, 5.32, 2.2, 0.25, "We pay ₹51,110", size=10, color="white", bold=True)
    d.bullets(s, M, 5.95, 7.6, 0.9, [
        "Net operating profit = your payouts, minus the rent and utilities you paid.",
        "Maximum payment is the programme fee you paid. Everything installed stays yours either way."],
        size=10, color="muted", gap=3)

    # 7 — Conditions ---------------------------------------------------------------------------
    s = d.slide()
    d.header(s, "The Earn-Back Promise", "The conditions, stated plainly.", size=36)
    conds = ["The property stays bookable at least 27 nights a month, excluding your agreed owner nights.",
             "We set the pricing, and you don’t override it.",
             "The property passed our assessment — and for a leased property, the rent is within the range we assessed.",
             "You approve major repairs within 7 days of us asking.",
             "You claim in writing within 30 days of month 12."]
    for i, c in enumerate(conds):
        y = 2.4 + i * 0.8
        d.rule(s, M, y, 7.3)
        H1(s, M, y + 0.14, 0.7, 0.5, f"{i + 1:02d}", size=22, color="brass")
        T(s, M + 0.85, y + 0.2, 6.4, 0.6, c, size=13, color="ink", leading=1.3)
    d.rule(s, M, 2.4 + 5 * 0.8, 7.3)
    qx = 8.6
    d.box(s, qx, 2.4, W - M - qx, 4.0, "ink")
    H1(s, qx + 0.4, 2.75, 3.3, 1.8, "No quibbles.\nWritten terms.\nPaid within 30 days.", size=24,
       color="bg", leading=1.15)
    d.rule(s, qx + 0.4, 4.6, 1.0, "brass")
    T(s, qx + 0.4, 4.78, 3.2, 1.5, "We can only make this promise because we run the property ourselves, to one "
      "standard. No manager who leaves the cleaning to you could offer it.", size=10.5, color="paper", leading=1.35)

    # 8 — Divider: what ₹1,11,111 buys (dark) --------------------------------------------------------
    s = d.slide(dark=True, chrome=False, transition=("fade", True))
    d.photo(s, 0, 0, W, H, "hotel_bed")
    d.veil(s, 35)
    d.mask(s, 0, 0, 8.5, H, "right")
    d.logo(s, W - M - 0.55, 0.32, 0.55)
    label(s, M, 1.4, 7, "What your ₹1,11,111 actually buys", color="brass_lt")
    H1(s, M, 1.8, 7, 2.6, "90+", size=180, leading=0.85)
    H1(s, M, 4.45, 6.6, 1.0, "pieces — sourced, delivered, installed and photographed.", size=28, leading=1.05)
    T(s, M, 5.65, 5.5, 0.5, "Plus the launch.", size=14, color="brass_lt", bold=True)

    # 9 / 10 — The kit ------------------------------------------------------------------------------
    kit = [
        [("lamp", "Lighting", ["Pendant lights, bedside lamps, a floor lamp and plug-in wall sconces",
                               "Warm 2700K bulbs in every fitting, dimmable where possible",
                               "Balcony or headboard string lights, hallway and bathroom night lights"]),
         ("linen", "Bedding & bath, hotel standard",
          ["Two full sets of white cotton linen per bedroom", "Mattress and pillow protectors; four pillows per bed",
           "Two sets of bath towels, hand towels and mats per bathroom", "Waffle robes, slippers, laundry bag, hamper"]),
         ("styling", "Soft styling", ["Six designer cushion covers with inserts, two textured throws",
                                      "Blackout curtains for every bedroom window", "Living and bedside rugs"]),
         ("chair", "Decor & finishing", ["A statement mirror, three framed prints in the Outpost style",
                                         "Four plants in ceramic or terracotta pots",
                                         "Coffee-table and bedside styling sets", "Scent diffuser and first refill"])],
        [("bath", "Guest kit", ["Tea and coffee station: kettle, jars, mugs, tray", "Welcome basket, first fill",
                                "Refillable toiletry dispensers", "Hair dryer, iron and board, umbrella stand"]),
         ("living", "Safety & technology", ["Smoke alarm, extinguisher, fire blanket, first-aid kit, emergency card",
                                            "Keyless smart lock; WiFi with a guest network",
                                            "Smart TV stick; bedside USB and charging points"]),
         (None, "Your branding — the Signature Kit", ["Engraved RED FLAG × [Your Property] brass door plate",
                                                       "Branded tissue boxes, coasters, key tags, labels, door hanger",
                                                       "Branded welcome card and printed house guide"]),
         ("mountain_bed", "The work, not just the things",
          ["Design plan, sourcing, delivery and a full installation day", "Deep clean and styling before the shoot",
           "Professional shoot, 25+ edited images", "Listings on Airbnb, Booking.com and your direct page",
           "AI pricing, creator launch, guidebook and house rules"])],
    ]
    for page, cols in enumerate(kit):
        s = d.slide()
        d.header(s, f"The kit  ·  {page + 1} of 2", "Everything a guest touches — chosen for you." if page == 0
                 else "Ready for guests on day one.", size=30)
        cw_, g = 2.72, 0.31
        for i, (img, t, items) in enumerate(cols):
            x = M + i * (cw_ + g)
            if img:
                d.photo(s, x, 2.2, cw_, 1.75, img)
            else:  # the engraved brass door plate, drawn natively
                d.box(s, x, 2.2, cw_, 1.75, "ink")
                plate = d.box(s, x + 0.3, 2.62, cw_ - 0.6, 0.92, "brass", line="brass_lt", round_=True)
                plate.name = "static-plate"
                for sx in (x + 0.45, x + cw_ - 0.55):
                    oval(s, sx, 3.03, 0.1, 0.1, "#7A5C2E", C=d.C).name = "static-screw"
                T(s, x + 0.3, 2.78, cw_ - 0.6, 0.3, "RED FLAG  ×", size=10, color="ink", bold=True,
                  align="center", spacing=300)
                T(s, x + 0.3, 3.08, cw_ - 0.6, 0.3, "YOUR PROPERTY", size=9, color="ink", bold=True,
                  align="center", spacing=300)
            H1(s, x, 4.12, cw_, 0.6, t, size=16.5, leading=1.0)
            d.bullets(s, x, 4.8, cw_, 1.9, items, size=9.5, gap=4)
        if page == 1:
            T(s, M, 6.72, 11.8, 0.3, "Optional Icon Kit: monogram curtains, branded linen, blankets, ceramics, robes and "
              "slippers. Studio ₹35,000 · Villa ₹75,000.", size=8.5, color="muted")

    # 11 — Every month ---------------------------------------------------------------------------
    s = d.slide()
    d.photo(s, 0, 0, 4.9, H, "bedroom_wood")
    x0 = 5.55
    label(s, x0, 1.2, 6, "What we do every month")
    H1(s, x0, 1.52, 7, 1.0, "You step back. We run it.", size=36)
    monthly = ["Dynamic AI pricing, adjusted daily to demand, season and local events",
               "24/7 guest messaging, from first enquiry to check-out",
               "Every clean, turnover and linen change arranged and inspected",
               "Minor maintenance up to ₹2,000 per incident, with all vendor coordination",
               "Review and reputation management",
               "Listings kept in sync across every channel",
               "Payouts twice a month with an itemised statement"]
    for i, m in enumerate(monthly):
        y = 2.55 + i * 0.58
        d.rule(s, x0, y, W - M - x0)
        oval(s, x0, y + 0.2, 0.16, 0.16, ACC["brass"], C=d.C)
        T(s, x0 + 0.38, y + 0.16, W - M - x0 - 0.4, 0.35, m, size=12.5, color="ink")
    d.rule(s, x0, 2.55 + 7 * 0.58, W - M - x0)

    # 12 — Not included ----------------------------------------------------------------------------
    s = d.slide()
    d.header(s, "What is not included — stated upfront", "No surprises on your first statement.", size=36)
    excl = [("Rent and security deposit (leased properties)", "You"),
            ("Electricity, water, internet, society charges, property tax", "You"),
            ("Repairs above ₹2,000 per incident", "You — and we ask before we spend"),
            ("Major appliance replacement (AC, geyser, fridge)", "You"),
            ("Furniture for an unfurnished property", "You — quoted separately, itemised, at cost"),
            ("Keys beyond the first two", "You — ₹45,000 per extra key"),
            ("Building and contents insurance", "You"),
            ("Cleaning, linen, consumables, minor repairs", "Deducted from booking revenue at cost")]
    label(s, M, 2.3, 4, "Item", color="muted")
    label(s, 7.9, 2.3, 4, "Who pays", color="muted")
    for i, (k, v) in enumerate(excl):
        y = 2.6 + i * 0.47
        last = i == len(excl) - 1
        if last:
            d.box(s, M, y, W - 2 * M, 0.47, "paper")
        d.rule(s, M, y, W - 2 * M)
        T(s, M + (0.15 if last else 0), y + 0.13, 6.8, 0.3, k, size=12.5, color="ink")
        T(s, 7.9, y + 0.13, 4.7, 0.3, v, size=12.5, color="brass" if last else "ink", bold=True)
    d.rule(s, M, 2.6 + 8 * 0.47, W - 2 * M)

    # 13 — Four stages -----------------------------------------------------------------------------
    s = d.slide()
    d.header(s, "How it works", "Four stages. You step back after the first.", size=36)
    stages = [("01", "The Hunt", "We check your property, or help you find one, against real demand data — before you commit a rupee."),
              ("02", "The Flip", "Design, furnishing, branding kit, smart lock and WiFi. Photograph-ready in under 45 days."),
              ("03", "The Drop", "Creators stay and post, a professional shoot, and listings go live on every channel."),
              ("04", "The Yield", "AI pricing, 24/7 guests, every clean and repair handled. Payouts on the 1st and 16th.")]
    sw = (W - 2 * M) / 4
    d.box(s, M, 2.55, sw - 0.2, 0.34, "paper")
    T(s, M + 0.15, 2.61, sw, 0.25, "You + us", size=9, color="ink", bold=True, spacing=150, caps=True)
    d.box(s, M + sw, 2.55, 3 * sw, 0.34, "ink")
    T(s, M + sw + 0.15, 2.61, 3 * sw, 0.25, "Us — you step back", size=9, color="white", bold=True, spacing=150,
      caps=True)
    d.rule(s, M, 3.55, W - 2 * M, "ink", 0.018)
    for i, (n, t, b) in enumerate(stages):
        x = M + i * sw
        oval(s, x, 3.46, 0.2, 0.2, ACC["red"] if i == 0 else ACC["brass"], C=d.C)
        H1(s, x, 3.95, sw - 0.3, 0.8, n, size=54, color="brass", leading=0.9)
        H1(s, x, 4.85, sw - 0.3, 0.5, t, size=24)
        T(s, x, 5.45, sw - 0.4, 1.2, b, size=11, color="body", leading=1.35)

    # 14 — How you get paid ------------------------------------------------------------------------
    s = d.slide()
    d.header(s, "How you get paid", "One statement. Twice a month.", size=36,
             sub="You never receive a separate bill.")
    parts = [("Gross\nbookings", "everything guests pay", "paper", "ink"), ("8%\n", "Red Flag commission", "paper", "ink"),
             ("At\ncost", "cleaning & maintenance", "paper", "ink"), ("Your\npayout", "straight to your account", "ink", "bg")]
    bw_, gap = 2.55, 0.5
    for i, (big, small, fill, fg) in enumerate(parts):
        x = M + i * (bw_ + gap)
        d.box(s, x, 2.8, bw_, 1.7, fill)
        H1(s, x + 0.25, 3.05, bw_ - 0.4, 0.9, big, size=24, color=fg, leading=1.0)
        T(s, x + 0.25, 4.1, bw_ - 0.4, 0.3, small, size=10.5, color="muted" if fg == "ink" else "paper")
        if i < 3:
            H1(s, x + bw_ + 0.08, 3.25, 0.35, 0.6, "−" if i < 2 else "=", size=30, color="brass", align="center")
    for i, dday in enumerate(("1st", "16th")):
        cx = M + i * 1.9
        oval(s, cx, 5.05, 1.4, 1.4, ACC["red"] if i == 0 else ACC["brass"], C=d.C)
        H1(s, cx, 5.43, 1.4, 0.6, dday, size=28, color="white", align="center")
    T(s, 4.8, 5.25, 7.6, 1.2, "Payouts land on the 1st and 16th of every month, with a statement showing every booking "
      "and every deduction. Gross bookings, minus 8%, minus cleaning and maintenance at cost, equals your payout.",
      size=12.5, color="body", leading=1.4)

    # 15 — Illustrative numbers --------------------------------------------------------------------
    s = d.slide()
    d.header(s, "Illustrative monthly numbers  ·  3-key villa", "What a month could look like.", size=34)
    occ = ["45%", "55%", "60%", "65%"]
    tbl = [("Gross booking revenue", ["₹1,21,500", "₹1,48,500", "₹1,62,000", "₹1,75,500"]),
           ("Less Red Flag 8%", ["₹9,720", "₹11,880", "₹12,960", "₹14,040"]),
           ("Less cleaning & maintenance", ["₹16,500", "₹18,500", "₹20,400", "₹22,300"]),
           ("= Net payout", ["₹95,280", "₹1,18,120", "₹1,28,640", "₹1,39,160"]),
           ("Less electricity & utilities", ["₹12,000"] * 4),
           ("Owner: more than a tenant pays", ["+₹8,280", "+₹31,120", "+₹41,640", "+₹52,160"]),
           ("Lessee: net after rent", ["₹8,280", "₹31,120", "₹41,640", "₹52,160"]),
           ("Months to earn back the fee", ["20.7", "5.5", "4.1", "3.3"])]
    lx, cx0, ccw = M, 3.85, 1.2
    ty = 2.35
    d.box(s, cx0 - 0.1, ty - 0.1, ccw, 0.44 + len(tbl) * 0.43 + 0.1, "paper")
    label(s, lx, ty + 0.05, 3, "Occupancy", color="muted")
    for j, o in enumerate(occ):
        H1(s, cx0 + j * ccw, ty - 0.02, ccw - 0.2, 0.4, o, size=18, color="red" if j == 0 else "ink", align="right")
    for i, (k, vals) in enumerate(tbl):
        y = ty + 0.44 + i * 0.43
        strong = k.startswith("=") or k.startswith("Months")
        d.rule(s, lx, y, cx0 + 4 * ccw - lx - 0.1, "ink" if strong else "hair")
        T(s, lx, y + 0.12, 3.1, 0.3, k, size=10.5, color="ink", bold=strong)
        for j, v in enumerate(vals):
            T(s, cx0 + j * ccw, y + 0.12, ccw - 0.2, 0.3, v, size=10.5 if not strong else 11.5,
              color="red" if (j == 0 and k.startswith("Months")) else "ink", bold=strong, align="right")
    nx = 9.15
    label(s, nx, 2.35, 3.5, "Read the first column honestly", color="red")
    T(s, nx, 2.72, W - M - nx, 2.6, "At 45% occupancy the fee takes longer than 12 months to earn back — which is exactly "
      "when the Earn-Back Promise pays you. We publish that column deliberately. A company that only shows you its best "
      "case is showing you the wrong thing.", size=11.5, color="ink", leading=1.4)
    d.rule(s, nx, 4.85, 1.0, "brass")
    T(s, nx, 5.0, W - M - nx, 1.5, "Owners: compare with what a tenant pays you today (₹75,000 assumed). Lessees: rent is "
      "₹75,000. Breakeven for a leased Outpost is around 41% occupancy.", size=9.5, color="muted", leading=1.35)
    T(s, M, 6.55, 11.8, 0.3, "Illustrative only — ₹9,000 average nightly rate, ₹75,000 market rent. You receive a "
      "projection for your own property during assessment.", size=8.5, color="muted")

    # 16 — Three ways -----------------------------------------------------------------------------
    s = d.slide()
    d.header(s, "Three ways to run a property", "Compare honestly.", size=34)
    opts = ["Do it yourself", "A co-host", "Outpost Classic"]
    cmp_rows = [("Commission", "None", "15–20%", "8%, min ₹4,999/month"),
                ("Joining fee", "None", "₹20,000–50,000 typical", "₹59,999 one time"),
                ("Setup and design", "Yours", "Yours", "₹1,11,111 · 90+ pieces, shoot, launch"),
                ("Cleaning", "You arrange and pay", "Usually you", "We arrange — at cost from bookings"),
                ("Repairs", "You", "Usually you", "Handled to ₹2,000 per incident"),
                ("Guests at 2am", "You", "Rarely", "Us, 24/7"),
                ("Pricing", "You guess", "Usually manual", "AI, daily"),
                ("Marketing", "Your problem", "Listing only", "Creator launch + direct bookings"),
                ("Your time", "10–15 hrs a month", "4–6 hrs a month", "Close to zero"),
                ("If it underperforms", "Your loss", "Your loss", "We pay you the difference")]
    c = [M, 3.4, 6.0, 8.75]
    oy, orh = 2.35, 0.41
    d.box(s, c[3] - 0.2, oy - 0.12, W - M - c[3] + 0.2, 0.42 + len(cmp_rows) * orh + 0.16, "ink")
    for j, o in enumerate(opts):
        T(s, c[j + 1], oy, 3.0, 0.3, o, size=9, color="white" if j == 2 else "muted", bold=True, spacing=150,
          caps=True)
    for i, row in enumerate(cmp_rows):
        y = oy + 0.42 + i * orh
        d.rule(s, M, y, c[3] - 0.2 - M)
        T(s, c[0], y + 0.11, 2.6, 0.3, row[0], size=10.5, color="ink", bold=True)
        T(s, c[1], y + 0.11, 2.5, 0.3, row[1], size=10.5, color="muted")
        T(s, c[2], y + 0.11, 2.6, 0.3, row[2], size=10.5, color="muted")
        T(s, c[3], y + 0.11, W - M - c[3] - 0.1, 0.3, row[3], size=10.5, color="white" if i < 9 else "brass_lt",
          bold=i == 9)

    # 17 — Your name on the door (dark) -------------------------------------------------------------
    s = d.slide(dark=True, transition=("fade", True))
    d.photo(s, 0, 0, 4.9, H, "oasis")
    x0 = 5.5
    label(s, x0, 1.15, 6, "Your name on the door", color="brass_lt")
    H1(s, x0, 1.48, 7.2, 1.0, "An identity, not a listing number.", size=32)
    perks = [("A higher nightly rate", "Design detail supports premium pricing. You stop being compared to the flat next door."),
             ("Free guest content", "Guests photograph and post branded details unprompted."),
             ("Instant credibility", "Reads as a hospitality brand, not someone’s spare flat."),
             ("The collection effect", "Guests who stay in one Red Flag property recognise and book others."),
             ("Exit value", "A branded, operating, revenue-producing property sells for more than a furnished flat."),
             ("Your name on the door", "An engraved RED FLAG × [Your Property] plate.")]
    pw_ = (W - M - x0 - 0.4) / 2
    for i, (t, b) in enumerate(perks):
        x = x0 + (i % 2) * (pw_ + 0.4)
        y = 2.6 + (i // 2) * 1.35
        d.rule(s, x, y, pw_, "hair")
        H1(s, x, y + 0.15, pw_, 0.4, t, size=16, color="ink")
        T(s, x, y + 0.55, pw_, 0.75, b, size=10.5, color="muted", leading=1.35)

    # 18 — Is this for you? ------------------------------------------------------------------------
    s = d.slide()
    d.header(s, "Is this for you?", "Honest fit, both ways.", size=36)
    yes = ["You own a flat or villa that is empty, under-used, or earning a modest rent",
           "You can put ₹1,71,110 to work without needing it back next month",
           "You want the income without the job",
           "Your property is in or near Goa, Mumbai, Bengaluru, Delhi NCR or Lucknow — or you live abroad and own one there"]
    no = ["You want a guaranteed monthly income regardless of bookings. Short stays don’t work that way — anyone promising it is not telling you the truth.",
          "You want to set your own nightly rates. Our pricing control is what makes the Earn-Back Promise possible.",
          "You need the property free for long stretches. Owner nights are agreed at signing.",
          "Your society or landlord prohibits short stays and won’t give written permission."]
    colw = (W - 2 * M - 0.5) / 2
    for j, (t, items, mark, mc, fill) in enumerate((("It works well if", yes, "✓", "brass", "paper"),
                                                     ("It is not for you if", no, "✕", "red", "bg"))):
        x = M + j * (colw + 0.5)
        d.box(s, x, 2.3, colw, 4.35, fill, line=None if j == 0 else "hair")
        H1(s, x + 0.35, 2.55, colw - 0.7, 0.5, t, size=22)
        d.bullets(s, x + 0.35, 3.25, colw - 0.7, 3.3, items, size=11.5, color="ink", gap=10, mark=mark,
                  mark_color=mc)

    # 19 — What we need from you ------------------------------------------------------------------
    s = d.slide()
    d.header(s, "What we need from you", "Six simple requirements.", size=36)
    need = [("Age & residency", "You are 18 or older, and an Indian citizen or valid resident."),
            ("Proof", "Proof of ownership, or a lease permitting short-stay use."),
            ("Permission", "Written permission from your society or landlord where required."),
            ("Clear title", "The property free of legal dispute or encumbrance."),
            ("KYC", "Basic KYC and bank details for payouts."),
            ("Access", "Access for cleaning, maintenance and guests — and approval of major repairs within 7 days.")]
    gw_ = (W - 2 * M - 2 * 0.3) / 3
    for i, (t, b) in enumerate(need):
        x = M + (i % 3) * (gw_ + 0.3)
        y = 2.35 + (i // 3) * 2.15
        d.box(s, x, y, gw_, 1.95, "paper")
        H1(s, x + 0.3, y + 0.25, 1, 0.5, f"{i + 1:02d}", size=24, color="brass")
        H1(s, x + 0.3, y + 0.8, gw_ - 0.6, 0.4, t, size=17)
        T(s, x + 0.3, y + 1.2, gw_ - 0.6, 0.7, b, size=10.5, color="body", leading=1.35)

    # 20 — Where we are (dark) ---------------------------------------------------------------------
    s = d.slide(dark=True, transition=("fade", True))
    d.photo(s, 6.2, 0, W - 6.2, H, "island")
    d.mask(s, 6.2, 0, 1.2, H, "right")
    label(s, M, 1.15, 6, "Where we are", color="brass_lt")
    for i, city in enumerate(["Goa", "Mumbai", "Bengaluru", "Delhi NCR", "Lucknow"]):
        H1(s, M, 1.55 + i * 0.62, 5, 0.6, city, size=34, leading=0.95)
    d.rule(s, M, 4.85, 5.0, "brass")
    H1(s, M, 5.0, 5.4, 0.5, "Living in Dubai, London or New York?", size=18, color="brass_lt", italic=True)
    T(s, M, 5.45, 5.3, 1.1, "Plant your flag at home — we run your Indian property end to end, fully remotely, and you "
      "join without flying back. Dubai Outposts open in 2027.", size=11.5, color="body", leading=1.35)

    # 21 / 22 — FAQ --------------------------------------------------------------------------------
    faq = [
        ("What is Outpost Classic?", "A programme where Red Flag designs, brands, launches and runs your short-stay property end to end. You pay a one-time programme fee and 8% of booking revenue."),
        ("What does it cost?", "₹59,999 franchise fee plus ₹1,11,111 setup, one time — ₹1,71,110 for up to two keys, plus GST. Extra keys ₹45,000 each. Commission 8%, min ₹4,999 a month."),
        ("When do I pay?", "After your property passes assessment and you sign. Nothing before that. The free assessment carries no obligation."),
        ("I already own a property. What do I pay?", "The programme fee, electricity and utilities. No rent, no deposit."),
        ("I don’t own a property. Can I still join?", "Yes. We help you find and assess a high-demand property before you sign a lease. You pay the programme fee, and rent and deposit to your landlord."),
        ("What if my property doesn’t qualify?", "We tell you why, in writing, and you pay nothing. We would rather decline a property than run one that will disappoint you."),
        ("What is the Earn-Back Promise?", "If net operating profit in the first 12 full months is less than the programme fee you paid, we pay the difference within 30 days of a valid claim. Full terms in the Programme Terms."),
        ("Is my return guaranteed?", "No. Earnings depend on occupancy, rates, season and location, and a property can earn less than its costs. The only guarantee is the Earn-Back Promise, exactly as written."),
        ("Who pays for cleaning and repairs?", "Cleaning, linen, consumables and repairs up to ₹2,000 per incident are arranged by us and deducted at cost, itemised. Anything above ₹2,000 needs your approval first."),
        ("Why is your commission only 8%?", "Our operations run on an AI system rather than a large payroll, so each extra property costs us little to run. The first 200 partners lock 8% for life."),
        ("Can I start with less than ₹1,71,110?", "Yes. With Setup Flex you pay ₹1,15,555 at signing and the remaining ₹55,556 comes from your first six payouts, at no extra charge."),
        ("How long until my property is live?", "Under 45 days from signing, if the property is ready to work on."),
        ("When and how am I paid?", "Twice a month, on the 1st and 16th, with a statement showing every booking and deduction."),
        ("Can I still use my property?", "Yes. Owner nights are agreed at signing and blocked in the calendar. Nights beyond that allowance are excluded from the Earn-Back Promise."),
        ("Can I add a second property later?", "Yes, and performing partners are prioritised. Each property has its own programme fee and its own Earn-Back Promise."),
        ("I live abroad. Can I join?", "Yes. NRIs can join fully remotely for a property in India. Payouts go to your NRO account and Indian tax is deducted at source as the law requires."),
        ("Can I leave the programme?", "Yes, with 90 days’ notice after an initial 12-month term. The fee is not refundable except under the Earn-Back Promise; everything installed stays yours."),
        ("What happens to the branding if I leave?", "We remove the Red Flag name and door plate. Furniture and fittings you paid for stay with you."),
    ]
    for page in range(2):
        s = d.slide()
        d.header(s, f"Frequently asked questions  ·  {page + 1} of 2",
                 "Straight answers." if page == 0 else "More straight answers.", size=30)
        fw = (W - 2 * M - 2 * 0.35) / 3
        for i, (q, a) in enumerate(faq[page * 9:(page + 1) * 9]):
            x = M + (i % 3) * (fw + 0.35)
            y = 2.05 + (i // 3) * 1.55
            d.rule(s, x, y, fw)
            T(s, x, y + 0.1, 0.5, 0.3, f"{page * 9 + i + 1:02d}", size=9, color="brass", bold=True)
            H1(s, x + 0.42, y + 0.07, fw - 0.42, 0.35, q, size=12.5, leading=1.0)
            T(s, x + 0.42, y + 0.42, fw - 0.42, 1.05, a, size=9, color="body", leading=1.3)

    # 23 — Next step -------------------------------------------------------------------------------
    s = d.slide()
    d.photo(s, 9.2, 0, W - 9.2, H, "infinity")
    label(s, M, 1.15, 6, "Your next step")
    H1(s, M, 1.48, 8, 1.0, "It takes 48 hours.", size=40)
    nxt = ["Apply for allocation at redflaghomes.in. Two minutes.",
           "We call you — usually the same day — and ask about your property and your city.",
           "We assess it against real demand data, free, with no obligation.",
           "You receive a projection within 48 hours — your rent, your city, your numbers.",
           "If the numbers work for both of us, you sign and pay. Live in under 45 days."]
    for i, t in enumerate(nxt):
        y = 2.55 + i * 0.73
        oval(s, M, y + 0.02, 0.42, 0.42, ACC["red"] if i == 0 else ACC["brass"], C=d.C)
        T(s, M, y + 0.1, 0.42, 0.3, str(i + 1), size=11, color="white", bold=True, align="center")
        if i < 4:
            d.box(s, M + 0.2, y + 0.46, 0.012, 0.3, "hair")
        T(s, M + 0.7, y + 0.1, 7.5, 0.4, t, size=13, color="ink")
    label(s, M, 6.35, 8, "Limited allocation  ·  Applications reviewed in the order received", color="red")

    # 24 — CTA (dark) --------------------------------------------------------------------------------
    s = d.slide(dark=True, chrome=False, transition=("morph", False))
    d.photo(s, 7.2, 0, W - 7.2, H, "resort_night")
    d.mask(s, 7.2, 0, 1.8, H, "right")
    d.logo(s, M - 0.08, 0.6, 1.3)
    label(s, M, 2.4, 6, "Apply for allocation", color="brass_lt")
    H1(s, M, 2.72, 7.5, 1.2, "redflaghomes.in", size=58, leading=0.9)
    T(s, M, 3.9, 6.3, 0.8, "Two minutes to apply. A projection for your own property within 48 hours — free, "
      "with no obligation.", size=14, color="body", leading=1.35)
    d.rule(s, M, 4.95, 6.0, "brass")
    T(s, M, 5.1, 7, 0.35, "hello@redflaghomes.in", size=14, color="ink", bold=True)
    T(s, M, 5.45, 7, 0.3, "Red Flag Homes Network · a company of Red Flag World Holdings Ltd, London", size=9.5,
      color="muted")
    T(s, M, 6.2, 6.4, 0.9, "Figures are illustrative, based on a 3-key villa, ₹9,000 average nightly rate and ₹75,000 "
      "market rent. Returns are not guaranteed and a property can earn less than its costs. The Earn-Back Promise is "
      "subject to the Programme Terms. All prices plus applicable GST.", size=7.5, color="muted", leading=1.3)

    assert len(d.prs.slides) == TOTAL, len(d.prs.slides)
    d.save()


if __name__ == "__main__":
    build()
