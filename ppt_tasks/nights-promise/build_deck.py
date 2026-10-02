"""Build The Nights Promise client deck — Red Flag Outpost Classic.

Concept: every promised night is a moon on a calendar. Build Mode with the pptx-designer public API;
charts and calendars are native shapes. Fonts: Playfair Display + DM Sans (see fonts/).
Art: make_art.py (own video frame + drawn moon + star field). Run: python make_art.py && python build_deck.py
"""
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.util import Inches, Pt

from pptx_designer import Presentation
from pptx_designer.tools.images import gradient_mask_image
from pptx_designer.tools.layout import add_slide, clean_save
from pptx_designer.tools.shapes import oval, rect, rrect, triangle
from pptx_designer.tools.text import text

import motion

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
ART = HERE / "art"
OUT = HERE / "output"
LOGO = str(ROOT / "logobg.png")

SERIF, SANS = "Playfair Display", "DM Sans"
W, H, M = 13.333, 7.5, 0.8
TOTAL = 18

NIGHT = {"bg": "#0E1424", "panel": "#182139", "ink": "#F1EBDD", "body": "#C9CCD6", "muted": "#8790A6",
         "hair": "#2B3653"}
DAY = {"bg": "#F1EBDD", "panel": "#E5DCCA", "ink": "#121A2E", "body": "#3D4458", "muted": "#7C8396",
       "hair": "#D2C7B2"}
ACC = {"gold": "#D9B56A", "red": "#E5483B", "white": "#FFFFFF", "navy": "#0E1424", "cream": "#F1EBDD"}


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.pal, self.plan, self.notes = NIGHT, {}, {}

    def col(self, k):
        return k if k.startswith("#") else (self.pal.get(k) or ACC[k])

    @property
    def C(self):
        return {"background": self.pal["bg"], "text_body": self.pal["body"], "text_dark": self.pal["ink"],
                "text_muted": self.pal["muted"], "font_body": SANS, "font_heading": SERIF}

    def T(self, s, x, y, w, h, txt, size=13, font=SANS, color="body", italic=False, bold=False, align="left",
          spacing=None, leading=None, caps=False):
        box = text(s, x, y, w, h, txt.upper() if caps else txt, font_size=size, color=self.col(color), bold=bold,
                   align=align, font_name=font, C=self.C)
        tf = box.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for p in tf.paragraphs:
            p.alignment = tf.paragraphs[0].alignment
            if leading:
                p.line_spacing = leading
            for r in p.runs:
                r.font.italic = italic
                if spacing is not None:
                    r.font._element.set("spc", str(spacing))
        return box

    def H1(self, s, x, y, w, h, txt, size=40, color="ink", italic=False, align="left", leading=1.0, bold=False):
        return self.T(s, x, y, w, h, txt, size=size, font=SERIF, color=color, italic=italic, align=align,
                      leading=leading, bold=bold)

    def label(self, s, x, y, w, txt, color="gold", size=8.5, align="left"):
        return self.T(s, x, y, w, 0.25, txt, size=size, color=color, bold=True, caps=True, spacing=300, align=align)

    def box(self, s, x, y, w, h, fill, line=None, round_=False, radius=0.1, line_w=None):
        f = (rrect if round_ else rect)(s, x, y, w, h, self.col(fill), line=self.col(line) if line else None, C=self.C)
        if round_:
            f.adjustments[0] = radius
        if line_w and line:
            f.line.width = Pt(line_w)
        return f

    def ring(self, s, x, y, d, color="hair", w=1.0, fill=None):
        o = oval(s, x, y, d, d, self.col(fill) if fill else "#000000", line=self.col(color), C=self.C)
        if not fill:
            o.fill.background()
        o.line.width = Pt(w)
        return o

    def moon(self, s, x, y, d, filled=True, num=None):
        if filled:
            oval(s, x, y, d, d, ACC["gold"], C=self.C)
        else:
            self.ring(s, x, y, d)
        if num is not None:
            self.T(s, x, y + d * .28, d, d * .5, str(num), size=max(6, d * 14), color="navy" if filled else "muted",
                   bold=filled, align="center")

    def rule(self, s, x, y, w, color="hair", t=0.012):
        return self.box(s, x, y, w, t, color)

    def picture(self, s, path, x, y, w=None, h=None, name="static-art"):
        pic = s.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
        pic.name = name
        return pic

    def logo(self, s, x, y, size):
        return self.picture(s, LOGO, x, y, w=size, h=size, name="!!logo")

    def slide(self, dark=True, sky=False, chrome=True, transition=("fade", False), notes=None):
        self.pal = NIGHT if dark else DAY
        s = add_slide(self.prs)
        n = len(self.prs.slides)
        self.plan[n] = transition
        if sky:
            self.picture(s, ART / "stars_soft.jpg", 0, 0, w=W, h=H, name="static-bg")
        else:
            rect(s, 0, 0, W, H, self.pal["bg"], C=self.C).name = "static-bg"
        if chrome:
            self.logo(s, W - M - 0.5, 0.32, 0.5)
            for shp in (self.label(s, M, 0.5, 6, "The Nights Promise", color="muted", size=7),
                        self.T(s, W - M - 1.5, H - 0.5, 1.5, 0.25, f"{n:02d} / {TOTAL:02d}", size=7, color="muted",
                               bold=True, spacing=200, align="right"),
                        self.label(s, M, H - 0.5, 5, "Red Flag Outpost Classic", color="muted", size=7)):
                shp.name = "chrome"
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s

    def save(self):
        for sl in self.prs.slides:
            tree = sl.shapes._spTree
            for shp in [x for x in sl.shapes if x.name == "!!logo"]:
                tree.remove(shp._element)
                tree.append(shp._element)
        for i, sl in enumerate(self.prs.slides, start=1):
            kind, black = self.plan.get(i, ("fade", False))
            motion.transition(sl, kind, through_black=black)
            motion.choreograph(sl, step_ms=110, budget_ms=2200)
        OUT.mkdir(exist_ok=True)
        path = OUT / "Nights_Promise.pptx"
        clean_save(self.prs, str(path))
        print(path)


PRESENTER = ("Before you present this:\n"
             "• Never say ₹185 without “with three properties” in the same sentence. A single-property owner who "
             "discovers it’s ₹555 after the call won’t sign.\n"
             "• Run their real nightly rate and cleaning cost before the meeting. Existing clients know their own numbers.\n"
             "• Quote the ₹4,999 minimum, not the 8%. At ₹2,000 a night and 18 nights, 8% is only about ₹2,880, so the "
             "minimum applies on most listings.\n"
             "• Never say “no questions asked”. Say what’s true and stronger: no claim form, no proof, no chasing — we count "
             "it and we pay.\n"
             "• The promise is nights, never income. Don’t let a conversation drift into promising rupees.\n\n"
             "Say on the call: Open with nights, not the programme name.")


def build():
    d = Deck()
    T, H1, label = d.T, d.H1, d.label

    # 1 — Cover -----------------------------------------------------------------------------------
    s = d.slide(chrome=False, transition=("fade", True), notes=PRESENTER)
    d.picture(s, ART / "night.jpg", 0, 0, w=W, h=H, name="static-photo")
    gradient_mask_image(s, 0, 0, 8.6, H, bg_color=NIGHT["bg"], direction="right", alpha_start=95,
                        alpha_end=0).name = "static-mask"
    d.picture(s, ART / "moon.png", 8.3, 0.35, w=4.4)
    d.logo(s, M - 0.06, 0.55, 1.15)
    label(s, M, 2.15, 7, "Red Flag Outpost Classic  ·  The Nights Promise")
    H1(s, M, 2.45, 8, 2.4, "We’ll put 18 nights\na month on your\ncalendar.", size=46, leading=1.02)
    H1(s, M, 5.15, 7.5, 0.6, "In writing. Or the money comes back to your bank.", size=19, italic=True, color="gold")
    d.box(s, M, 6.1, 7.15, 0.46, "panel", round_=True, radius=0.5)
    T(s, M, 6.21, 7.15, 0.3, "From ₹185 a night with three properties  ·  paid back automatically if we miss",
      size=10, color="ink", bold=True, align="center")

    # 2 — The problem (day) ------------------------------------------------------------------------
    s = d.slide(dark=False)
    label(s, M, 1.15, 6, "You already know the problem", color="red")
    H1(s, M, 1.5, 6, 1.6, "Most months, the\ncalendar sits half empty.", size=36, leading=1.05)
    for i, t in enumerate(["Most properties in India sit booked 10–14 nights a month",
                           "The rate never changes, so weekends go cheap and weekdays go empty",
                           "Enquiries come at midnight, the cleaner cancels on Sunday"]):
        y = 3.1 + i * 0.75
        d.rule(s, M, y, 5.6)
        T(s, M, y + 0.15, 5.6, 0.55, t, size=12, color="ink", leading=1.25)
    d.rule(s, M, 3.1 + 3 * 0.75, 5.6)
    H1(s, M, 5.55, 5.8, 0.9, "It isn’t the property. Nobody is working the calendar every single day.", size=16,
       italic=True, color="body", leading=1.25)
    cx, cy, cd, gap = 7.35, 1.55, 0.6, 0.1
    for j, wd in enumerate("MTWTFSS"):
        T(s, cx + j * (cd + gap), cy, cd, 0.25, wd, size=9, color="muted", bold=True, align="center")
    booked = {3, 4, 5, 10, 11, 12, 17, 18, 19, 24, 25, 26}            # 12 nights, mostly weekends
    for day in range(1, 31):
        k = day + 1                                                      # month starts on a Wednesday
        r, c = divmod(k, 7)
        d.moon(s, cx + c * (cd + gap), cy + 0.35 + r * (cd + gap), cd, filled=day in booked, num=day)
    T(s, cx, cy + 0.35 + 5 * (cd + gap) + 0.15, 4.8, 0.3, "●  12 nights booked        ○  18 nights empty", size=10,
      color="muted", bold=True)

    # 3 — So we put it in writing (sky) ------------------------------------------------------------------
    s = d.slide(sky=True, transition=("fade", True),
                notes="Say on the call: Pause here. This is the slide the conversation turns on.")
    label(s, M, 1.15, 6, "So we put it in writing")
    H1(s, M, 1.35, 6, 2.6, "108", size=190, color="gold", leading=0.9)
    T(s, M + 0.1, 4.15, 6, 0.4, "nights per property, in 6 months", size=18, color="ink", bold=True)
    T(s, M + 0.1, 4.7, 5.3, 1.0, "An average of 18 a month. Miss it, and your fee comes back — automatically.",
      size=14, color="body", leading=1.4)
    mw, md, mg = 1.62, 0.17, 0.055
    for m in range(6):
        x0 = 7.0 + (m % 3) * (mw + 0.3)
        y0 = 1.6 + (m // 3) * 2.35
        T(s, x0, y0, mw, 0.25, f"Month {m + 1}", size=9, color="muted", bold=True, spacing=150)
        for day in range(30):
            r, c = divmod(day, 7)
            filled = day in {1, 2, 4, 5, 6, 8, 9, 11, 12, 13, 15, 16, 18, 19, 20, 22, 25, 26}
            d.moon(s, x0 + c * (md + mg), y0 + 0.35 + r * (md + mg), md, filled=filled)
        T(s, x0, y0 + 0.35 + 5 * (md + mg) + 0.08, mw, 0.25, "18 nights", size=9, color="gold", bold=True)
    T(s, 7.0, 6.35, 5.5, 0.3, "6 months × 18 nights = 108 nights, in writing", size=10, color="muted", italic=True)

    # 4 — What it costs (day) ------------------------------------------------------------------------------
    s = d.slide(dark=False)
    label(s, M, 1.15, 6, "What it costs")
    H1(s, M, 1.5, 5.5, 1.5, "One fee.\nEverything else is 8%.", size=38, leading=1.05)
    H1(s, M, 3.4, 5, 1.0, "₹59,999", size=60, color="red")
    T(s, M, 4.45, 5, 0.3, "joining fee, one time", size=12, color="muted", bold=True)
    rows = [("Joining fee, one time", "₹59,999"), ("Setup fee", "None — your property, your furniture"),
            ("Commission", "8%, or ₹4,999 per listing a month, whichever is higher"),
            ("Rate", "Locked at 8% for life"), ("Covers", "Up to 3 properties, onboarded together")]
    for i, (k, v) in enumerate(rows):
        y = 1.6 + i * 0.95
        d.rule(s, 6.6, y, W - M - 6.6)
        label(s, 6.6, y + 0.18, 3, k, color="muted", size=8)
        H1(s, 6.6, y + 0.42, W - M - 6.6, 0.5, v, size=17 if len(v) < 40 else 14.5, leading=1.1)
    d.rule(s, 6.6, 1.6 + 5 * 0.95, W - M - 6.6)

    # 5 — Divide it by the nights (sky) ---------------------------------------------------------------------
    s = d.slide(sky=True, transition=("fade", True),
                notes="Say on the call: This is the closing slide. Show ₹185 first, then ₹555 — the gap is what makes "
                      "them find a third property.")
    label(s, M, 1.15, 6, "Now divide it by the nights")
    H1(s, M, 1.45, 6, 1.6, "₹185", size=120, color="gold", leading=0.9)
    T(s, M + 0.05, 3.15, 6, 0.4, "a night, with three properties", size=18, color="ink", bold=True)
    T(s, M + 0.05, 3.65, 5.4, 0.3, "One time. Never again.", size=14, color="gold", italic=True)
    T(s, M + 0.05, 4.3, 5.2, 1.2, "One fee covers three properties. The third costs you nothing extra and cuts your "
      "per-night cost to a third.", size=13, color="body", leading=1.45)
    bars = [("1 property", "108 nights", 555, "₹555"), ("2 properties", "216 nights", 278, "₹278"),
            ("3 properties", "324 nights", 185, "₹185")]
    bx, bmax = 7.0, 4.9
    for i, (k, n, v, t) in enumerate(bars):
        y = 1.75 + i * 1.45
        hi = i == 2
        T(s, bx, y, 3, 0.3, k, size=13, color="ink", bold=True)
        T(s, bx + 1.55, y + 0.02, 2, 0.3, n, size=10.5, color="muted")
        d.box(s, bx, y + 0.42, bmax * v / 555, 0.5, "gold" if hi else "hair", round_=True, radius=0.5)
        H1(s, bx + bmax * v / 555 + 0.15, y + 0.4, 1.5, 0.5, t, size=24, color="gold" if hi else "body")
        T(s, bx, y + 1.0, 5.5, 0.25, "per promised night", size=8.5, color="muted")

    # 6 — What you keep per night (day) --------------------------------------------------------------------
    s = d.slide(dark=False, notes="Say on the call: Say “around” — we’ll run your real rate and your real cleaning cost "
                                  "on the assessment.")
    label(s, M, 1.15, 6, "What you keep per night")
    H1(s, M, 1.5, 5, 1.6, "You pay ₹185\nfor the night.", size=38, leading=1.05)
    T(s, M, 2.9, 5, 0.3, "with three properties", size=11, color="muted", bold=True, caps=True, spacing=200)
    H1(s, M, 3.6, 5.2, 1.3, "You keep around\n₹1,270 of it.", size=32, color="red", leading=1.1)
    T(s, M, 5.15, 4.9, 0.8, "Illustrative, on a typical ₹2,000 night. We’ll run your real rate and cleaning cost on "
      "the assessment.", size=10.5, color="muted", leading=1.4)
    base, scale, bw, x0 = 6.3, 4.4 / 2000, 1.15, 6.4
    steps = [("Typical nightly rate", 2000, 0, "ink", "₹2,000"), ("Our commission", 280, 1720, "red", "−₹280"),
             ("Cleaning, linen, consumables", 450, 1270, "muted", "−₹450"), ("Lands with you", 1270, 0, "gold", "₹1,270")]
    for i, (k, v, start, c, t) in enumerate(steps):
        x = x0 + i * (bw + 0.45)
        top = base - (start + v) * scale
        d.box(s, x, top, bw, v * scale, c)
        H1(s, x - 0.2, top - 0.5, bw + 0.4, 0.4, t, size=18, color="red" if c == "red" else "ink", align="center")
        T(s, x - 0.2, base + 0.12, bw + 0.4, 0.5, k, size=9, color="body", align="center", leading=1.2)
        if 0 < i:
            d.rule(s, x - 0.45, base - (start + v if i < 3 else v) * scale, 0.45, "muted", 0.01)
    d.rule(s, x0 - 0.1, base, 4 * bw + 3 * 0.45 + 0.2, "ink", 0.015)

    # 7 — Three properties, six months (sky) -------------------------------------------------------------------
    s = d.slide(sky=True, notes="Say on the call: The point isn’t the total. It’s that a one-time fee buys a permanent income.")
    label(s, 0, 1.2, W, "Across three properties, six months", align="center")
    H1(s, 0, 1.75, W, 1.5, "₹4,11,000", size=96, color="gold", align="center")
    T(s, 0, 3.45, W, 0.35, "324 nights at about ₹1,270 to you", size=15, color="ink", bold=True, align="center")
    d.rule(s, W / 2 - 2.6, 4.0, 5.2, "hair")
    H1(s, 0, 4.2, W, 0.6, "−₹59,999", size=30, color="body", align="center")
    T(s, 0, 4.85, W, 0.3, "your joining fee", size=11, color="muted", align="center")
    d.rule(s, W / 2 - 2.6, 5.4, 5.2, "hair")
    H1(s, 0, 5.6, W, 0.6, "And the fee never repeats.", size=24, italic=True, color="ink", align="center")
    T(s, 0, 6.35, W, 0.3, "Illustrative — your nights, rate and costs are calculated on the assessment.", size=9,
      color="muted", align="center")

    # 8 — If we miss (dark) --------------------------------------------------------------------------------------
    s = d.slide(transition=("fade", True), notes="Say on the call: This is the trust slide. Most guarantees are designed "
                "to be hard to claim. Ours doesn’t need claiming at all.")
    label(s, M, 1.15, 8, "If we miss, you don’t have to chase us", color="red")
    for i, t in enumerate(["No claim form.", "No proof.", "No chasing."]):
        H1(s, M, 1.55 + i * 0.85, 6, 0.8, t, size=44, italic=i == 2, color="gold" if i == 2 else "ink")
    pts = [("01", "We count the nights from our own booking records — not yours"),
           ("02", "We check at month six, whether you ask or not"),
           ("03", "If we’re short, the refund goes to your bank within 30 days"),
           ("04", "If you think our count is wrong, you have 30 days to tell us")]
    for i, (n, t) in enumerate(pts):
        y = 1.65 + i * 1.12
        d.box(s, 7.2, y, W - M - 7.2, 0.95, "panel", round_=True, radius=0.12)
        T(s, 7.45, y + 0.3, 0.7, 0.4, n, size=16, color="gold", bold=True)
        T(s, 8.25, y + 0.24, W - M - 8.5, 0.6, t, size=12.5, color="ink", leading=1.3)
    T(s, M, 4.6, 5.6, 1.2, "We count it, and we pay. The refund is pro-rata across the properties that fell short — "
      "three enrolled and one short means ₹20,000 back.", size=12.5, color="body", leading=1.45)

    # 9 — What you don't pay (day) -----------------------------------------------------------------------------
    s = d.slide(dark=False)
    label(s, M, 1.15, 6, "What you don’t pay")
    H1(s, M, 1.5, 11, 0.9, "Four things other managers charge for.", size=34)
    zero = [("Setup fee", "Your property is already furnished"), ("Per-listing onboarding", "No charge per listing"),
            ("Monthly retainer", "None"), ("Photography, listing writing, pricing", "All included")]
    cw = (W - 2 * M - 3 * 0.3) / 4
    for i, (k, v) in enumerate(zero):
        x = M + i * (cw + 0.3)
        d.box(s, x, 2.75, cw, 3.4, "panel", round_=True, radius=0.06)
        T(s, x + 0.3, 3.05, cw - 0.6, 1.1, "₹0", size=54, color="red", bold=True)
        d.rule(s, x + 0.3, 4.35, cw - 0.6, "hair")
        H1(s, x + 0.3, 4.5, cw - 0.6, 0.8, k, size=15, leading=1.1)
        T(s, x + 0.3, 5.55, cw - 0.6, 0.6, v, size=10.5, color="muted", leading=1.3)

    # 10 — What we do, every day (sky) ---------------------------------------------------------------------------
    s = d.slide(sky=True)
    label(s, M, 1.15, 6, "What we do, every day")
    H1(s, M, 1.5, 11, 0.9, "We work the calendar. Every single day.", size=34)
    jobs = [("Reprice", "Your calendar daily — weekends, events, season, local demand"),
            ("Rank", "Rewrite and rank your listing on Airbnb, Booking.com and direct"),
            ("Answer", "Every guest, 24/7, enquiry to checkout"),
            ("Clean", "Every clean, turnover and linen change arranged"),
            ("Fix", "Minor repairs and vendor coordination handled"),
            ("Pay", "You, twice a month — 1st and 16th, itemised")]
    gw = (W - 2 * M - 2 * 0.35) / 3
    for i, (k, v) in enumerate(jobs):
        x = M + (i % 3) * (gw + 0.35)
        y = 2.7 + (i // 3) * 1.85
        d.moon(s, x, y + 0.05, 0.42)
        H1(s, x + 0.62, y, gw - 0.6, 0.5, k, size=22, color="ink")
        T(s, x + 0.62, y + 0.58, gw - 0.7, 0.9, v, size=11.5, color="body", leading=1.4)

    # 11 — What you do, once (day) ----------------------------------------------------------------------------
    s = d.slide(dark=False)
    label(s, M, 1.15, 6, "What you do, once")
    H1(s, M, 1.5, 8, 0.9, "One setup spec. Then you step back.", size=34)
    T(s, M, 2.35, 9, 0.6, "We assess your property and send a written setup spec — what to fix, add or replace before "
      "we list it.", size=13, color="body", leading=1.4)
    once = [("Within 30 days", "You complete it, at your cost"), ("Sign-off", "Our team inspects and signs off in writing"),
            ("Day one", "The six months start at sign-off, not at payment")]
    sw = (W - 2 * M) / 3
    d.rule(s, M, 3.62, W - 2 * M, "ink", 0.015)
    for i, (k, v) in enumerate(once):
        x = M + i * sw
        oval(s, x, 3.52, 0.22, 0.22, ACC["gold"] if i < 2 else ACC["red"], C=d.C)
        T(s, x, 4.0, sw - 0.3, 0.5, k, size=19, color="ink", bold=True)
        T(s, x, 4.55, sw - 0.4, 0.6, v, size=12, color="body", leading=1.35)
    d.box(s, M, 5.55, W - 2 * M, 0.85, "panel", round_=True, radius=0.1)
    d.box(s, M, 5.55, 0.07, 0.85, "red")
    T(s, M + 0.35, 5.7, W - 2 * M - 0.6, 0.6, "Skip items on the spec and the promise doesn’t apply — it’s the only way we "
      "can guarantee nights on a property we didn’t furnish.", size=12, color="ink", leading=1.35)

    # 12 — Three properties, one fee (dark) -----------------------------------------------------------------------
    s = d.slide(notes="Say on the call: Say this plainly now. A partner who assumes it follows every future property "
                      "feels cheated in month eight.")
    label(s, M, 1.15, 6, "Three properties, one fee")
    H1(s, M, 1.5, 8, 0.9, "₹59,999 covers up to three.", size=34)

    def house(x, y, filled):
        roof = triangle(s, x, y, 1.6, 0.8, ACC["gold"] if filled else NIGHT["bg"],
                        line=None if filled else ACC["gold"], C=d.C)
        body = rect(s, x + 0.2, y + 0.8, 1.2, 1.0, ACC["gold"] if filled else NIGHT["bg"],
                    line=None if filled else ACC["gold"], C=d.C)
        for shp in (roof, body):
            if not filled:
                shp.line.dash_style = MSO_LINE.DASH
                shp.line.width = Pt(1.25)
        rect(s, x + 0.62, y + 1.2, 0.36, 0.6, NIGHT["bg"] if filled else NIGHT["bg"], C=d.C)

    for i in range(4):
        x = M + 0.1 + i * 2.1
        house(x, 2.75, filled=i < 3)
        T(s, x - 0.2, 4.75, 2.0, 0.3, f"Property {i + 1}" if i < 3 else "Added later", size=11,
          color="ink" if i < 3 else "muted", bold=True, align="center")
    T(s, M, 5.25, 6.1, 0.9, "Handed over together on day one — covered by the nights promise.", size=12, color="gold",
      bold=True, leading=1.35)
    x0 = 9.3
    d.box(s, x0, 2.6, W - M - x0, 3.6, "panel", round_=True, radius=0.06)
    H1(s, x0 + 0.3, 2.85, W - M - x0 - 0.6, 0.5, "Adding more later?", size=20)
    T(s, x0 + 0.3, 3.45, W - M - x0 - 0.6, 2.6, "You keep 8% for life on those too.\n\nBut the nights promise covers "
      "only the properties we assessed at onboarding — up to three.", size=12, color="body", leading=1.4)

    # 13 — The conditions, plainly (day) ------------------------------------------------------------------------
    s = d.slide(dark=False, notes="Say on the call: Owners trust the floor-rate line more than anything else in the deck.")
    label(s, M, 1.15, 6, "The conditions, plainly")
    H1(s, M, 1.5, 11, 0.9, "Four conditions. No small print.", size=34)
    conds = [("Available", "At least 27 nights a month, with up to 12 owner nights across the six months", False),
             ("Control", "We control pricing, availability and listing content", False),
             ("Floor rate", "Nights count at or above a floor rate we agree together — so we can’t hit the number by "
                            "dumping your price", True),
             ("Repairs", "Approved within 7 days", False)]
    for i, (k, v, hi) in enumerate(conds):
        y = 2.65 + i * 0.95
        if hi:
            d.box(s, M - 0.15, y + 0.02, W - 2 * M + 0.3, 0.92, "panel", round_=True, radius=0.1)
        d.rule(s, M, y, W - 2 * M)
        H1(s, M, y + 0.25, 3, 0.5, k, size=22, color="red" if hi else "ink")
        T(s, 4.0, y + 0.28, W - M - 4.0, 0.6, v, size=13, color="ink", leading=1.35, bold=hi)
    d.rule(s, M, 2.65 + 4 * 0.95, W - 2 * M)

    # 14 — How it starts (sky) -----------------------------------------------------------------------------------
    s = d.slide(sky=True)
    label(s, M, 1.15, 6, "How it starts")
    H1(s, M, 1.5, 11, 0.9, "From first call to go-live in about five weeks.", size=32)
    tl = [("Week 1", "Free assessment. We check real demand and agree your floor rate."),
          ("Week 1", "You join. Setup spec issued."),
          ("Weeks 2–5", "You complete the spec. We inspect and sign off."),
          ("Go-live", "Listed, priced, live. The six months begin.")]
    sw = (W - 2 * M) / 4
    d.rule(s, M, 3.7, W - 2 * M, "hair", 0.015)
    for i, (k, v) in enumerate(tl):
        x = M + i * sw
        d.moon(s, x, 3.52, 0.36, filled=i == 3)
        if i < 3:
            d.ring(s, x, 3.52, 0.36, color="gold", w=1.25, fill=NIGHT["bg"])
        T(s, x, 4.15, sw - 0.3, 0.5, k, size=20, color="gold" if i == 3 else "ink", bold=True)
        T(s, x, 4.7, sw - 0.45, 1.2, v, size=12, color="body", leading=1.4)
    T(s, M, 6.2, 11, 0.3, "The 108 nights are counted from go-live, after sign-off.", size=11, color="muted", italic=True)

    # 15 — Next step (dark) -----------------------------------------------------------------------------------------
    s = d.slide(transition=("fade", True), notes="Say on the call: Ending on “if they can’t, we’ll say so” earns more "
                                                  "trust than any promise earlier in the deck.")
    d.picture(s, ART / "moon.png", 8.6, 1.0, w=4.4)
    label(s, M, 1.4, 6, "Next step")
    H1(s, M, 1.8, 7.6, 2.4, "Send us your listing links.", size=46, leading=1.05)
    T(s, M, 3.55, 7.0, 1.0, "We’ll assess them free and tell you whether they can hit the number — before you pay "
      "anything.", size=16, color="body", leading=1.45)
    d.rule(s, M, 4.85, 1.2, "gold", 0.02)
    H1(s, M, 5.05, 7.6, 0.8, "If they can’t, we’ll say so.", size=30, italic=True, color="gold")

    # 16 — Close -------------------------------------------------------------------------------------------------
    s = d.slide(chrome=False, transition=("fade", True))
    d.picture(s, ART / "night.jpg", 0, 0, w=W, h=H, name="static-photo")
    gradient_mask_image(s, 0, 0, W, H, bg_color=NIGHT["bg"], direction="top", alpha_start=20,
                        alpha_end=85).name = "static-mask"
    d.logo(s, W / 2 - 0.65, 0.8, 1.3)
    label(s, 0, 2.4, W, "Red Flag Outpost Classic", align="center")
    H1(s, 0, 2.8, W, 0.9, "From ₹185 a night, with three properties.", size=38, align="center")
    H1(s, 0, 3.65, W, 0.7, "8% for life. Nights in writing.", size=28, italic=True, color="gold", align="center")
    T(s, 0, 4.7, W, 0.35, "redflaghomes.in   ·   hello@redflaghomes.in", size=15, color="ink", bold=True,
      align="center")
    T(s, 1.5, 6.55, W - 3, 0.4, "Figures are illustrative. The Nights Promise is subject to its terms. Returns are not "
      "guaranteed. Prices plus GST.", size=8.5, color="muted", align="center")

    # 17–18 — FAQ ------------------------------------------------------------------------------------------------
    faq = [
        ("What exactly are you promising?", "A number of nights, not an amount of money. 108 qualifying nights per property across six months — an average of 18 a month. The number for your property is written into your agreement."),
        ("What happens if you miss it?", "We refund your joining fee, pro-rata across the properties that fell short. Three properties enrolled and one short means ₹20,000 back."),
        ("Do I have to claim it?", "No. We count the nights from our own booking records and check at month six whether you ask or not. If we’re short, the money goes to your bank within 30 days."),
        ("What if I think your count is wrong?", "You have 30 days from the statement to tell us. We review and respond within 15 days."),
        ("Why ₹185 a night for some and ₹555 for others?", "The fee is the same ₹59,999 whether you enrol one property or three. One property buys 108 promised nights, three buys 324 — so the cost per night falls from ₹555 to ₹185."),
        ("Can I add a fourth property?", "Yes, and you keep 8% for life on it. But the nights promise covers only the properties we assessed at onboarding, up to three."),
        ("Why no setup fee?", "Your property is already furnished. We give you a written spec of what to fix, you do it at your cost, and we sign it off before listing."),
        ("What if I don’t want to do the setup work?", "Then the promise doesn’t apply — we can’t guarantee nights on a property that isn’t ready. We’ll still manage it at 8%, and you can do the work later."),
        ("Can you just drop my price to hit the number?", "No. We agree a floor rate with you at onboarding and nights sold below it don’t count toward the promise. It protects you and it stops us gaming our own guarantee."),
        ("Can I still use my property myself?", "Yes — up to 12 owner nights across the six months, blocked in advance. Beyond that the promise no longer applies, because we can’t sell nights that aren’t available."),
        ("What do I actually earn per night?", "On a ₹2,000 night, roughly ₹1,270 lands with you after our commission and cleaning. Your real number depends on your rate and your city — we’ll calculate it on the assessment."),
        ("When do I get paid?", "Twice a month, on the 1st and the 16th, with a statement showing every booking and every deduction."),
        ("What if my city is seasonal?", "We set your number from your market, not a template. A Goa property might be promised fewer nights than a Bengaluru one. We’d rather promise a number we can hit."),
        ("What if you decline my property?", "We tell you why, in writing, and you pay nothing. We’d rather turn down a property than take a fee for nights we can’t deliver."),
        ("Is my income guaranteed?", "No. We guarantee nights, not rupees. What those nights earn depends on your rate, your season and your market."),
    ]
    for page, chunk in enumerate((faq[:8], faq[8:])):
        s = d.slide(dark=False)
        label(s, M, 1.0, 8, f"Frequently asked questions  ·  {page + 1} of 2")
        H1(s, M, 1.3, 10, 0.7, "Straight answers." if page == 0 else "More straight answers.", size=28)
        fw = (W - 2 * M - 0.5) / 2
        for i, (q, a) in enumerate(chunk):
            x = M + (i % 2) * (fw + 0.5)
            y = 2.05 + (i // 2) * 1.2
            d.rule(s, x, y, fw)
            H1(s, x, y + 0.1, fw, 0.35, q, size=13.5, bold=True)
            T(s, x, y + 0.45, fw, 0.72, a, size=9, color="body", leading=1.28)

    assert len(d.prs.slides) == TOTAL, len(d.prs.slides)
    d.save()


if __name__ == "__main__":
    build()
