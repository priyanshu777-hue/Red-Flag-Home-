"""Build The Nights Promise · United States owner deck — Red Flag Managed.

US product-brand look: Bricolage Grotesque + Inter, aurora gradients, bento tiles and native, editable UI mockups
(calendar, promise card, payout receipt, refund notice, listing card). Build Mode, pptx-designer public API.
Compliance (from the brief): no franchise vocabulary in selling copy, nights never income, the listing stays in
the owner's name, an "Illustrative" line under every performance number.
Run: python make_art.py && python build_deck.py
"""
from pathlib import Path

from lxml import etree
from pptx.enum.dml import MSO_LINE
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
ART = HERE / "art"
OUT = HERE / "output"
LOGO = str(ROOT / "logobg.png")

DISPLAY, SANS = "Bricolage Grotesque", "Inter"
W, H, M = 13.333, 7.5, 0.75
TOTAL = 19
ILLUS = "Illustrative. Your rate, market and costs will change this."

DARK = {"bg": "#090C1C", "card": "#151A33", "ink": "#FFFFFF", "body": "#C5C9DA", "muted": "#8A90AA", "hair": "#2A3152"}
LIGHT = {"bg": "#F7F5F1", "card": "#FFFFFF", "ink": "#0B1020", "body": "#41465A", "muted": "#7D8296", "hair": "#E3E0D8"}
ACC = {"coral": "#FF5C4D", "indigo": "#5846E5", "violet": "#8B5CF6", "amber": "#FFB547", "green": "#1DB954",
       "white": "#FFFFFF", "night": "#090C1C", "mist": "#EEF0FA", "soft": "#F1EFEA"}


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.pal, self.plan = DARK, {}

    def col(self, k):
        return k if k.startswith("#") else (self.pal.get(k) or ACC[k])

    @property
    def C(self):
        return {"background": self.pal["bg"], "text_body": self.pal["body"], "text_dark": self.pal["ink"],
                "text_muted": self.pal["muted"], "font_body": SANS, "font_heading": DISPLAY}

    # ---- primitives ---------------------------------------------------------------------------
    def T(self, s, x, y, w, h, txt, size=13, font=SANS, color="body", bold=False, italic=False, align="left",
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

    def H1(self, s, x, y, w, h, txt, size=44, color="ink", align="left", leading=0.95, spacing=-100):
        return self.T(s, x, y, w, h, txt, size=size, font=DISPLAY, color=color, bold=True, align=align,
                      leading=leading, spacing=spacing)

    def eyebrow(self, s, x, y, txt, color="coral", w=8, align="left"):
        return self.T(s, x, y, w, 0.25, txt, size=9, color=color, bold=True, caps=True, spacing=200, align=align)

    def shadow(self, shp, blur=28, dist=8, alpha=18):
        sp = shp._element.spPr
        eff = etree.SubElement(sp, qn("a:effectLst"))
        sh = etree.SubElement(eff, qn("a:outerShdw"), blurRad=str(blur * 12700), dist=str(dist * 12700),
                              dir="5400000", algn="t", rotWithShape="0")
        c = etree.SubElement(sh, qn("a:srgbClr"), val="0B1020")
        etree.SubElement(c, qn("a:alpha"), val=str(alpha * 1000))
        return shp

    def alpha(self, shp, pct):
        clr = shp.fill._xPr.find(qn("a:solidFill"))[0]
        clr.append(clr.makeelement(qn("a:alpha"), {"val": str(pct * 1000)}))
        return shp

    def card(self, s, x, y, w, h, fill="card", radius=0.08, shadow=True, line=None, glass=False):
        f = rrect(s, x, y, w, h, self.col(fill), line=self.col(line) if line else None, C=self.C)
        f.adjustments[0] = radius
        if line:
            f.line.width = Pt(0.75)
        if glass:
            self.alpha(f, 10)
            f.line.color.rgb = __import__("pptx").dml.color.RGBColor.from_string("FFFFFF")
            f.line.width = Pt(0.75)
            ln = f.line._get_or_add_ln().find(qn("a:solidFill"))[0]
            ln.append(ln.makeelement(qn("a:alpha"), {"val": "25000"}))
        elif shadow:
            self.shadow(f)
        return f

    def pill(self, s, x, y, w, h, fill, txt, color="white", size=9, bold=True):
        p = rrect(s, x, y, w, h, self.col(fill), C=self.C)
        p.adjustments[0] = 0.5
        self.T(s, x, y + (h - size / 72 * 1.25) / 2, w, h, txt, size=size, color=color, bold=bold, align="center")
        return p

    def dot(self, s, x, y, d, fill):
        return oval(s, x, y, d, d, self.col(fill), C=self.C)

    def rule(self, s, x, y, w, color="hair", t=0.012):
        return rect(s, x, y, w, t, self.col(color), C=self.C)

    def picture(self, s, path, x, y, w=None, h=None, name="static-art"):
        pic = s.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
        pic.name = name
        return pic

    def logo(self, s, x, y, size):
        return self.picture(s, LOGO, x, y, w=size, h=size, name="!!logo")

    def slide(self, dark=True, bg=None, chrome=True, transition=("fade", False), notes=None):
        self.pal = DARK if dark else LIGHT
        s = add_slide(self.prs)
        n = len(self.prs.slides)
        self.plan[n] = transition
        if bg:
            self.picture(s, ART / f"{bg}.jpg", 0, 0, w=W, h=H, name="static-bg")
        else:
            rect(s, 0, 0, W, H, self.pal["bg"], C=self.C).name = "static-bg"
        if chrome:
            self.logo(s, W - M - 0.5, 0.3, 0.5)
            for shp in (self.T(s, M, 0.48, 5, 0.25, "Red Flag Managed  ·  The Nights Promise", size=8, color="muted",
                               bold=True),
                        self.T(s, W - M - 1.5, H - 0.48, 1.5, 0.25, f"{n:02d}", size=8, color="muted", bold=True,
                               align="right")):
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
            motion.choreograph(sl, step_ms=90, budget_ms=2000)
        OUT.mkdir(exist_ok=True)
        path = OUT / "Nights_Promise_USA.pptx"
        clean_save(self.prs, str(path))
        print(path)

    # ---- UI mockups (native, editable) -----------------------------------------------------------
    def calendar_card(self, s, x, y, w, booked, title, sub, month="October", accent="coral"):
        h = w * 0.95
        self.card(s, x, y, w, h, fill="#FFFFFF", radius=0.06)
        self.T(s, x + 0.3, y + 0.28, w - 0.6, 0.3, month, size=13, color="#0B1020", bold=True)
        self.pill(s, x + w - 1.95, y + 0.24, 1.65, 0.34, "#E9F8EE" if accent == "coral" else "#FFF1EF", sub,
                  color="#13843D" if accent == "coral" else "#D8402F", size=8.5)
        cw = (w - 0.6) / 7
        for j, wd in enumerate("SMTWTFS"):
            self.T(s, x + 0.3 + j * cw, y + 0.75, cw, 0.2, wd, size=8, color="#9AA0B4", bold=True, align="center")
        for day in range(1, 32):
            k = day + 2
            r, c = divmod(k, 7)
            cx, cy = x + 0.3 + c * cw, y + 1.02 + r * cw * 0.8
            if day in booked:
                b = rrect(s, cx + 0.04, cy, cw - 0.08, cw * 0.66, ACC[accent], C=self.C)
                b.adjustments[0] = 0.3
                self.T(s, cx, cy + cw * 0.17, cw, 0.2, str(day), size=8, color="white", bold=True, align="center")
            else:
                self.T(s, cx, cy + cw * 0.17, cw, 0.2, str(day), size=8, color="#B4B8C8", align="center")
        self.rule(s, x + 0.3, y + h - 0.62, w - 0.6, "#ECEAE4")
        self.T(s, x + 0.3, y + h - 0.45, w - 0.6, 0.3, title, size=10.5, color="#0B1020", bold=True)


def build():
    d = Deck()
    T, H1, eb = d.T, d.H1, d.eyebrow

    PRE = ("Before you present:\n"
           "• Never say $2.16 without “with three properties” in the same sentence.\n"
           "• Check the city ordinance before the call. Don’t pitch an owner in a city that bans short-term rentals — "
           "New York City is the obvious one.\n"
           "• Confirm our licensing position in that state before quoting. Several states require a real estate broker "
           "license to manage property for compensation.\n"
           "• Run their real nightly rate before the meeting. US owners know their numbers.\n"
           "• Never say “no questions asked.” Say: no claim form, no proof, no chasing — we count it and we pay.\n"
           "• Never say franchise. Not in a deck, not on a call, not in a DM.\n\n"
           "Language rules: never use franchise, franchisee, territory or royalty — this is a management agreement with "
           "an onboarding fee. Never promise income, revenue, profit or return — we promise nights. The listing stays "
           "in the owner’s name; we do not brand their property. Every performance number needs the illustrative line.\n\n"
           "Say on the call: Open with nights. Never with the programme name.")

    # 1 — Cover ------------------------------------------------------------------------------------
    s = d.slide(bg="aurora", chrome=False, transition=("fade", True), notes=PRE)
    d.logo(s, M - 0.05, 0.6, 0.95)
    eb(s, M, 1.85, "Red Flag Managed  ·  The Nights Promise", color="amber")
    H1(s, M, 2.15, 7.3, 2.9, "We’ll put 18 nights a month on your calendar.", size=52, leading=0.95)
    T(s, M, 4.95, 6.6, 0.8, "In writing. Or your onboarding fee comes back — automatically.", size=18,
      color="white", leading=1.3)
    d.pill(s, M, 6.0, 6.4, 0.5, "#FFFFFF", "From $2.16 a night with three properties  ·  Refunded automatically if we miss",
           color="#0B1020", size=10)
    booked = {1, 2, 3, 4, 8, 9, 10, 11, 15, 16, 17, 18, 22, 23, 24, 25, 29, 30}
    c = d.card(s, 8.62, 1.22, 4.36, 4.16, fill="#FFFFFF", glass=True)
    d.calendar_card(s, 8.85, 1.45, 3.9, booked, "18 nights booked this month", "✓  On track")
    T(s, 8.85, 5.55, 3.9, 0.3, "Illustrative calendar.", size=8.5, color="body")

    # 2 — The problem (light) ----------------------------------------------------------------------------
    s = d.slide(dark=False)
    eb(s, M, 1.1, "You already know the problem")
    H1(s, M, 1.45, 6.2, 1.5, "Most rentals sit\nhalf empty.", size=50)
    big = d.card(s, M, 3.2, 2.7, 2.1, fill="night", radius=0.08)
    H1(s, M + 0.3, 3.4, 2.3, 1.0, "12–15", size=48, color="amber")
    T(s, M + 0.3, 4.4, 2.2, 0.7, "nights a month — where most short-term rentals sit", size=10, color="white", leading=1.3)
    pts = ["The rate barely moves, so weekends sell cheap and weekdays sit empty",
           "Guests message at midnight and the cleaner cancels on Sunday"]
    for i, t in enumerate(pts):
        y = 3.2 + i * 1.08
        d.card(s, 3.7, y, 3.3, 0.95, radius=0.1)
        T(s, 3.95, y + 0.18, 2.9, 0.7, t, size=10.5, color="ink", leading=1.3)
    T(s, M, 5.65, 6.4, 0.8, "It isn’t the property. Nobody is working the calendar every single day.", size=15,
      color="ink", bold=True, leading=1.3)
    d.calendar_card(s, 7.75, 1.25, 4.8, {6, 7, 13, 14, 20, 21, 27, 28, 3, 4, 10, 17, 24}, "13 nights booked · 18 empty",
                    "Half empty", accent="coral")
    s.shapes[-1]  # noqa: B018 — keep order
    T(s, 7.75, 6.0, 4.8, 0.3, "Illustrative calendar.", size=8.5, color="muted")

    # 3 — So we put it in writing (dark) -------------------------------------------------------------------
    s = d.slide(bg="aurora2", transition=("fade", True), notes="Say on the call: Pause here. This is the slide everything turns on.")
    eb(s, M, 1.1, "So we put it in writing", color="amber")
    H1(s, M, 1.35, 6.5, 2.6, "108", size=210, color="white", leading=0.85, spacing=-400)
    T(s, M + 0.1, 4.2, 6, 0.4, "nights per property, in six months", size=20, color="white", bold=True)
    T(s, M + 0.1, 4.75, 5.5, 0.8, "An average of 18 a month. Miss it, and your onboarding fee comes back.", size=14,
      color="body", leading=1.4)
    cx, cy, cw, ch = 7.6, 1.3, 5.0, 4.9                                     # the promise, as a signed document
    d.card(s, cx, cy, cw, ch, fill="#FFFFFF", radius=0.05)
    T(s, cx + 0.4, cy + 0.35, 3, 0.25, "THE NIGHTS PROMISE", size=8.5, color="#FF5C4D", bold=True, spacing=200)
    H1(s, cx + 0.4, cy + 0.65, cw - 0.8, 0.6, "Written into your agreement", size=20, color="#0B1020", spacing=0)
    rows = [("Nights per property", "108"), ("Period", "6 months"), ("Monthly average", "18"),
            ("If we miss", "Onboarding fee refunded"), ("Claim needed", "None")]
    for i, (k, v) in enumerate(rows):
        y = cy + 1.45 + i * 0.47
        d.rule(s, cx + 0.4, y, cw - 0.8, "#ECEAE4")
        T(s, cx + 0.4, y + 0.13, 2.4, 0.3, k, size=10, color="#7D8296")
        T(s, cx + 2.4, y + 0.13, cw - 2.8, 0.3, v, size=10, color="#0B1020", bold=True, align="right")
    d.rule(s, cx + 0.4, cy + ch - 0.95, 2.2, "#0B1020", 0.01)
    T(s, cx + 0.4, cy + ch - 0.82, 2.4, 0.25, "Red Flag Managed", size=8.5, color="#7D8296")
    sx, sy, sd = cx + cw - 1.35, cy + ch - 1.05, 0.85
    oval(s, sx, sy, sd, sd, ACC["coral"], C=d.C)
    T(s, sx, sy + 0.25, sd, 0.2, "IN", size=7.5, color="white", bold=True, align="center", spacing=100)
    T(s, sx, sy + 0.43, sd, 0.2, "WRITING", size=7.5, color="white", bold=True, align="center", spacing=100)

    # 4 — What it costs (light bento) ------------------------------------------------------------------------
    s = d.slide(dark=False, notes="Say on the call: Most full-service managers charge 15–25%. Let them react to 8% before "
                                  "you move on.")
    eb(s, M, 1.1, "What it costs")
    H1(s, M, 1.45, 11, 0.9, "One onboarding fee. Then 8%.", size=40)
    gx, gy, g = M, 2.55, 0.25
    bw = (W - 2 * M - 2 * g) / 3
    tiles = [(gx, gy, bw, 3.95, "night", [("$699", 72, "amber"), ("Onboarding fee, one time", 12, "white")]),
             (gx + bw + g, gy, bw, 1.85, "card", [("8%", 44, "coral"), ("of booking revenue — or $99 per listing a month, whichever is higher", 10, "body")]),
             (gx + 2 * (bw + g), gy, bw, 1.85, "card", [("Locked for life", 24, "ink"), ("Your 8% rate never rises", 10, "body")]),
             (gx + bw + g, gy + 2.1, bw, 1.85, "card", [("$0 setup", 30, "ink"), ("Your property, your furniture", 10, "body")]),
             (gx + 2 * (bw + g), gy + 2.1, bw, 1.85, "coral", [("Up to 3", 34, "white"), ("properties, onboarded together", 10, "white")])]
    for (x, y, w, h, fill, lines) in tiles:
        d.card(s, x, y, w, h, fill=fill, radius=0.06)
        big, small = lines
        if h > 3:
            H1(s, x + 0.35, y + 1.2, w - 0.7, 1.2, big[0], size=big[1], color=big[2])
            T(s, x + 0.4, y + 2.55, w - 0.7, 0.4, small[0], size=small[1] + 2, color=small[2], bold=True)
        else:
            H1(s, x + 0.35, y + 0.3, w - 0.7, 0.8, big[0], size=big[1], color=big[2])
            T(s, x + 0.35, y + 1.15, w - 0.7, 0.6, small[0], size=small[1], color=small[2], leading=1.3)

    # 5 — Divide by the nights (dark) ------------------------------------------------------------------------
    s = d.slide(bg="aurora3", transition=("fade", True),
                notes="Say on the call: Show $2.16 first, then $6.47. The gap is what makes a two-property owner find a "
                      "third. Never say $2.16 without “with three properties” in the same sentence.")
    gradient_mask_image(s, 0, 0, 7.2, H, bg_color=DARK["bg"], direction="right", alpha_start=85,
                        alpha_end=0).name = "static-mask"
    eb(s, M, 1.1, "Now divide it by the nights", color="amber")
    H1(s, M, 1.4, 6.5, 1.8, "$2.16", size=130, color="white", spacing=-300)
    T(s, M + 0.05, 3.3, 6, 0.4, "a night, with three properties", size=20, color="white", bold=True)
    T(s, M + 0.05, 3.8, 6, 0.4, "One time. Never again.", size=15, color="amber", bold=True)
    T(s, M + 0.05, 4.45, 5.4, 0.9, "One fee covers three properties. The third costs nothing extra and cuts your "
      "per-night cost to a third.", size=13, color="body", leading=1.45)
    T(s, M + 0.05, 5.6, 5.5, 0.3, "Illustrative: $699 ÷ promised nights.", size=9, color="muted")
    rows = [("1 property", "108 nights", "$6.47", 6.47), ("2 properties", "216 nights", "$3.24", 3.24),
            ("3 properties", "324 nights", "$2.16", 2.16)]
    for i, (k, n, v, val) in enumerate(rows):
        y = 1.55 + i * 1.5
        hi = i == 2
        d.card(s, 7.55, y, 5.05, 1.25, fill="#FFFFFF" if hi else "#FFFFFF", glass=not hi, radius=0.12)
        T(s, 7.9, y + 0.25, 2.5, 0.3, k, size=13, color="#0B1020" if hi else "white", bold=True)
        T(s, 7.9, y + 0.6, 2.5, 0.3, n, size=10, color="#7D8296" if hi else "body")
        H1(s, 10.2, y + 0.25, 2.1, 0.7, v, size=34, color="#FF5C4D" if hi else "white", align="right", spacing=-150)
        if hi:
            d.pill(s, 7.9, y + 0.92, 1.4, 0.24, "#0B1020", "BEST VALUE", size=7)

    # 6 — What you keep per night (light, receipt) ----------------------------------------------------------------
    s = d.slide(dark=False, bg="dawn")
    eb(s, M, 1.1, "What you keep per night")
    H1(s, M, 1.45, 6, 2.0, "You pay $2.16\nfor the night.", size=46)
    T(s, M, 3.1, 6, 0.3, "with three properties", size=12, color="muted", bold=True, caps=True, spacing=150)
    H1(s, M, 3.7, 6, 1.4, "About $138 of it\nlands with you.", size=38, color="coral")
    T(s, M, 5.3, 5.6, 0.9, "Illustrative. Your rate, your market, your own costs — utilities, supplies, taxes and "
      "insurance — will change this. We’ll run your real numbers at assessment.", size=10, color="muted", leading=1.4)
    rx, ry, rw = 7.6, 1.15, 5.0
    d.card(s, rx, ry, rw, 5.4, fill="#FFFFFF", radius=0.05)
    T(s, rx + 0.45, ry + 0.4, 3, 0.25, "PER NIGHT · EXAMPLE", size=8.5, color="#7D8296", bold=True, spacing=200)
    lines = [("Typical nightly rate", "$150", "#0B1020"), ("Management fee at 8%", "−$12", "#FF5C4D"),
             ("Cleaning", "Paid by the guest", "#13843D")]
    for i, (k, v, cc) in enumerate(lines):
        y = ry + 0.95 + i * 0.75
        T(s, rx + 0.45, y, 2.8, 0.35, k, size=13, color="#41465A")
        T(s, rx + 2.6, y, rw - 3.05, 0.35, v, size=14 if i < 2 else 11.5, color=cc, bold=True, align="right")
        d.rule(s, rx + 0.45, y + 0.52, rw - 0.9, "#ECEAE4")
    for k in range(18):                                                   # receipt perforation
        d.dot(s, rx + 0.45 + k * (rw - 0.9) / 17.6, ry + 3.35, 0.06, "#DDD9CF")
    T(s, rx + 0.45, ry + 3.7, 2.5, 0.35, "Lands with you", size=14, color="#0B1020", bold=True)
    H1(s, rx + 2.4, ry + 3.5, rw - 2.85, 0.9, "$138", size=54, color="#0B1020", align="right", spacing=-200)
    T(s, rx + 0.45, ry + 4.75, rw - 0.9, 0.3, "Illustrative, before your own operating costs and taxes.", size=8.5,
      color="#7D8296")

    # 7 — Three properties, six months (dark) ---------------------------------------------------------------------
    s = d.slide(bg="aurora", notes="Say on the call: The point isn’t the total — it’s that a one-time fee buys a permanent "
                                   "service. Always add: before your own operating costs and taxes.")
    gradient_mask_image(s, 0, 0, W, H, bg_color=DARK["bg"], direction="right", alpha_start=80,
                        alpha_end=20).name = "static-mask"
    eb(s, 0, 1.1, "Across three properties, six months", color="amber", w=W, align="center")
    H1(s, 0, 1.55, W, 1.8, "$44,700", size=128, color="white", align="center", spacing=-400)
    T(s, 0, 3.45, W, 0.35, "324 nights at about $138 to you", size=16, color="white", bold=True, align="center")
    eq = [("−$699", "your onboarding fee"), ("Never", "And the fee never repeats")]
    for i, (v, k) in enumerate(eq):
        x = W / 2 - 4.6 + i * 4.8
        d.card(s, x, 4.25, 4.4, 1.2, glass=True, radius=0.12)
        H1(s, x + 0.35, 4.4, 1.9, 0.7, v, size=32, color="amber" if i else "white", spacing=-100)
        T(s, x + 2.2, 4.62, 2.0, 0.7, k, size=10.5, color="body", leading=1.3)
    T(s, 0, 5.85, W, 0.3, "Illustrative, before your own operating costs and taxes. We guarantee nights, not income.",
      size=10, color="muted", align="center")

    # 8 — If we miss (dark, notification) ----------------------------------------------------------------------------
    s = d.slide(transition=("fade", True), notes="Say on the call: Most guarantees are built to be hard to claim. This one "
                                                  "doesn’t need claiming at all. That’s the trust slide.")
    eb(s, M, 1.1, "If we miss, you don’t chase us")
    for i, t in enumerate(["No claim form.", "No proof.", "No chasing."]):
        H1(s, M, 1.45 + i * 0.85, 6.5, 0.8, t, size=50, color="coral" if i == 2 else "white")
    pts = ["We count the nights from our own booking records, not yours", "We check at month six whether you ask or not",
           "If we’re short, the refund goes to your account within 30 days",
           "If you think our count is wrong, you have 30 days to tell us"]
    for i, t in enumerate(pts):
        y = 4.35 + i * 0.52
        d.dot(s, M, y + 0.07, 0.14, "coral")
        T(s, M + 0.35, y, 6.2, 0.4, t, size=12, color="body")
    px, py, pw, ph = 8.4, 0.95, 3.6, 5.9                                     # phone
    d.card(s, px, py, pw, ph, fill="#1B2140", radius=0.12, line="#3A4270")
    d.card(s, px + 0.15, py + 0.15, pw - 0.3, ph - 0.3, fill="#0E1330", radius=0.1, shadow=False)
    d.pill(s, px + pw / 2 - 0.55, py + 0.3, 1.1, 0.22, "#000000", "", size=6)
    T(s, px, py + 0.75, pw, 0.5, "9:41", size=30, color="white", bold=True, align="center")
    T(s, px, py + 1.35, pw, 0.3, "Month six", size=10, color="body", align="center")
    n1 = d.card(s, px + 0.3, py + 2.0, pw - 0.6, 1.25, fill="#FFFFFF", radius=0.12)
    d.dot(s, px + 0.5, py + 2.2, 0.34, "coral")
    T(s, px + 0.95, py + 2.2, 2.2, 0.25, "Red Flag Managed", size=8.5, color="#7D8296", bold=True)
    T(s, px + 0.95, py + 2.45, 2.3, 0.3, "Refund sent: $233", size=12, color="#0B1020", bold=True)
    T(s, px + 0.5, py + 2.8, pw - 1.0, 0.4, "1 of 3 properties fell short. No claim needed.", size=8.5, color="#41465A")
    d.card(s, px + 0.3, py + 3.45, pw - 0.6, 0.9, glass=True, radius=0.14)
    T(s, px + 0.5, py + 3.62, pw - 1.0, 0.6, "Nights counted from our own booking records", size=9, color="body",
      leading=1.3)
    T(s, px, py + ph - 0.65, pw, 0.3, "Example: $699 ÷ 3 properties", size=8.5, color="muted", align="center")

    # 9 — What you don't pay (light) --------------------------------------------------------------------------------
    s = d.slide(dark=False)
    eb(s, M, 1.1, "What you don’t pay")
    H1(s, M, 1.45, 11, 0.9, "No setup. No retainer. No extras.", size=40)
    zero = [("Setup fee", "Your property is already furnished"), ("Per-listing onboarding", "No charge per listing"),
            ("Monthly retainer", "None"), ("Photos, rewrites, pricing", "All included")]
    cw = (W - 2 * M - 3 * 0.25) / 4
    for i, (k, v) in enumerate(zero):
        x = M + i * (cw + 0.25)
        d.card(s, x, 2.65, cw, 3.6, radius=0.06)
        H1(s, x + 0.35, 2.95, cw - 0.7, 1.1, "$0", size=72, color="coral", spacing=-200)
        d.rule(s, x + 0.35, 4.35, cw - 0.7)
        T(s, x + 0.35, 4.55, cw - 0.7, 0.6, k, size=15, color="ink", bold=True, leading=1.2)
        T(s, x + 0.35, 5.35, cw - 0.7, 0.6, v, size=10.5, color="muted", leading=1.3)

    # 10 — What we do, every day (dark bento) --------------------------------------------------------------------------
    s = d.slide(bg="aurora2")
    eb(s, M, 1.1, "What we do, every day", color="amber")
    H1(s, M, 1.45, 11, 0.9, "We work the calendar. Every single day.", size=40)
    jobs = [("Reprice", "Your calendar daily — weekends, events, season, local demand", "coral"),
            ("Rank", "Rewrite and rank your listing on Airbnb, Vrbo, Booking.com and direct", "violet"),
            ("Answer", "Every guest, 24/7, inquiry to checkout", "amber"),
            ("Clean", "Schedule every clean, turnover and linen change", "green"),
            ("Maintain", "Coordinate maintenance and vendors", "indigo"),
            ("Pay", "You, twice a month, with an itemized statement", "coral")]
    gw = (W - 2 * M - 2 * 0.25) / 3
    for i, (k, v, c) in enumerate(jobs):
        x = M + (i % 3) * (gw + 0.25)
        y = 2.6 + (i // 3) * 1.95
        d.card(s, x, y, gw, 1.75, glass=True, radius=0.08)
        d.dot(s, x + 0.35, y + 0.35, 0.36, c)
        T(s, x + 0.35, y + 0.42, 0.36, 0.25, f"{i + 1}", size=9, color="white", bold=True, align="center")
        H1(s, x + 0.9, y + 0.3, gw - 1.1, 0.5, k, size=22, color="white", spacing=-50)
        T(s, x + 0.35, y + 0.95, gw - 0.7, 0.7, v, size=11, color="body", leading=1.35)

    # 11 — Your property, your listing (light, listing card) -------------------------------------------------------------
    s = d.slide(dark=False, bg="dawn", notes="Say on the call: American owners care about keeping their account and reviews. "
                                             "Lead with that — it removes the biggest fear of switching.")
    eb(s, M, 1.1, "It stays yours")
    H1(s, M, 1.45, 6.3, 1.6, "Your property.\nYour listing.", size=50)
    own = ["The listing stays in your name, with your reviews and your history",
           "We don’t brand your property or put our name on your door",
           "You keep full ownership and control of the asset", "Cancel on 30 days’ notice, any time"]
    for i, t in enumerate(own):
        y = 3.35 + i * 0.68
        d.pill(s, M, y + 0.02, 0.34, 0.34, "green", "✓", size=9)
        T(s, M + 0.5, y + 0.03, 5.6, 0.5, t, size=13, color="ink", leading=1.3)
    lx, ly, lw = 7.75, 1.2, 4.8
    d.card(s, lx, ly, lw, 5.2, fill="#FFFFFF", radius=0.05)
    d.picture(s, ART / "listing.jpg", lx + 0.3, ly + 0.3, w=lw - 0.6, h=2.4)
    hx, hy = lx + lw / 2 - 0.75, ly + 0.85                               # a simple home mark
    roof = __import__("pptx_designer.tools.shapes", fromlist=["triangle"]).triangle(s, hx, hy, 1.5, 0.7, "#FFFFFF", C=d.C)
    rect(s, hx + 0.2, hy + 0.7, 1.1, 0.85, "#FFFFFF", C=d.C)
    rrect(s, hx + 0.58, hy + 1.0, 0.34, 0.55, ACC["coral"], C=d.C)
    oval(s, hx + 0.62, hy + 0.3, 0.26, 0.26, ACC["amber"], C=d.C)
    d.pill(s, lx + 0.5, ly + 0.5, 1.5, 0.3, "#FFFFFF", "Your listing", color="#0B1020", size=8.5)
    T(s, lx + 0.3, ly + 2.9, lw - 0.6, 0.35, "Your home, your title", size=15, color="#0B1020", bold=True)
    T(s, lx + 0.3, ly + 3.3, lw - 0.6, 0.3, "★ Your rating  ·  Your reviews  ·  Your history", size=10.5, color="#41465A")
    d.rule(s, lx + 0.3, ly + 3.85, lw - 0.6, "#ECEAE4")
    d.dot(s, lx + 0.3, ly + 4.1, 0.55, "#0B1020")
    T(s, lx + 0.3, ly + 4.27, 0.55, 0.25, "YOU", size=7.5, color="white", bold=True, align="center")
    T(s, lx + 1.0, ly + 4.1, 3.4, 0.3, "Hosted by you", size=12, color="#0B1020", bold=True)
    T(s, lx + 1.0, ly + 4.42, 3.4, 0.3, "Managed by Red Flag, behind the scenes", size=9.5, color="#7D8296")

    # 12 — What you do, once (light) ---------------------------------------------------------------------------------
    s = d.slide(dark=False)
    eb(s, M, 1.1, "What you do, once")
    H1(s, M, 1.45, 11, 0.9, "One setup spec. Then you step back.", size=40)
    T(s, M, 2.4, 10, 0.6, "We assess your property and send a written setup spec — what to fix, add or replace before "
      "we relist it.", size=14, color="body", leading=1.4)
    steps = [("Within 30 days", "You complete the spec, at your cost"), ("Sign-off", "Our team inspects and signs off in writing"),
             ("Day one", "The six months start at sign-off, not at payment")]
    sw = (W - 2 * M - 2 * 0.25) / 3
    for i, (k, v) in enumerate(steps):
        x = M + i * (sw + 0.25)
        d.card(s, x, 3.35, sw, 1.75, radius=0.08)
        bar = rrect(s, x + 0.35, 3.65, sw - 0.7, 0.12, ACC["mist"], C=d.C)
        bar.adjustments[0] = 0.5
        fill = rrect(s, x + 0.35, 3.65, (sw - 0.7) * (i + 1) / 3, 0.12, ACC["coral"] if i == 2 else ACC["indigo"], C=d.C)
        fill.adjustments[0] = 0.5
        T(s, x + 0.35, 3.95, sw - 0.7, 0.4, k, size=17, color="ink", bold=True)
        T(s, x + 0.35, 4.4, sw - 0.7, 0.6, v, size=11.5, color="body", leading=1.35)
    warn = d.card(s, M, 5.45, W - 2 * M, 0.85, fill="#FFF1EF", radius=0.12, shadow=False)
    T(s, M + 0.35, 5.62, W - 2 * M - 0.7, 0.6, "Skip items on the spec and the promise doesn’t apply — it’s the only way "
      "we can guarantee nights on a property we didn’t furnish.", size=12, color="#B0342A", bold=True, leading=1.35)

    # 13 — Three properties, one fee (dark) ------------------------------------------------------------------------------
    s = d.slide(bg="aurora3")
    gradient_mask_image(s, 0, 0, W, H, bg_color=DARK["bg"], direction="bottom", alpha_start=90,
                        alpha_end=30).name = "static-mask"
    eb(s, M, 1.1, "Three properties, one fee", color="amber")
    H1(s, M, 1.45, 11, 0.9, "$699 covers up to three.", size=44)
    pw_ = (W - 2 * M - 3 * 0.25) / 4
    for i in range(4):
        x = M + i * (pw_ + 0.25)
        if i < 3:
            d.card(s, x, 2.7, pw_, 2.3, fill="#FFFFFF", radius=0.08)
            d.dot(s, x + 0.3, 2.95, 0.5, "coral")
            T(s, x + 0.3, 3.08, 0.5, 0.25, str(i + 1), size=11, color="white", bold=True, align="center")
            T(s, x + 0.3, 3.65, pw_ - 0.6, 0.35, f"Property {i + 1}", size=15, color="#0B1020", bold=True)
            T(s, x + 0.3, 4.05, pw_ - 0.6, 0.8, "Enrolled on day one · 108 nights promised", size=10, color="#41465A",
              leading=1.35)
        else:
            f = d.card(s, x, 2.7, pw_, 2.3, glass=True, radius=0.08)
            f.line.dash_style = MSO_LINE.DASH
            T(s, x + 0.3, 3.65, pw_ - 0.6, 0.35, "Added later", size=15, color="white", bold=True)
            T(s, x + 0.3, 4.05, pw_ - 0.6, 0.8, "8% for life — but no nights promise", size=10, color="body", leading=1.35)
    T(s, M, 5.45, 11.5, 0.9, "The promise applies to the properties enrolled on day one. Add more later: you keep 8% for "
      "life on those too, but the nights promise covers only what we assessed at onboarding.", size=12.5, color="body",
      leading=1.45)

    # 14 — The conditions, plainly (light) ------------------------------------------------------------------------------
    s = d.slide(dark=False, notes="Say on the call: Owners trust the floor-rate line more than anything else in the deck.")
    eb(s, M, 1.1, "The conditions, plainly")
    H1(s, M, 1.45, 11, 0.9, "Five conditions. No small print.", size=40)
    conds = [("Available", "At least 27 nights a month, with up to 12 owner nights across the six months", False),
             ("Control", "We control pricing, availability and listing content", False),
             ("Floor rate", "Nights count at or above a floor rate we agree together — so we can’t hit the number by "
                            "dumping your rate", True),
             ("Repairs", "Approved within 7 days", False),
             ("Permitted", "Your property stays permitted and insured for short-term rental", False)]
    for i, (k, v, hi) in enumerate(conds):
        y = 2.55 + i * 0.82
        d.card(s, M, y, W - 2 * M, 0.7, fill="night" if hi else "card", radius=0.15, shadow=not hi)
        T(s, M + 0.35, y + 0.2, 2.4, 0.35, k, size=14, color="amber" if hi else "ink", bold=True)
        T(s, M + 2.9, y + 0.21, W - 2 * M - 3.2, 0.4, v, size=12, color="white" if hi else "body")

    # 15 — How it starts (dark) -------------------------------------------------------------------------------------
    s = d.slide(bg="aurora2")
    eb(s, M, 1.1, "How it starts", color="amber")
    H1(s, M, 1.45, 11, 0.9, "From first call to live in about five weeks.", size=40)
    tl = [("Week 1", "Free assessment. We check real demand in your market and agree your floor rate."),
          ("Week 1", "You enroll. Setup spec issued. W-9 and insurance confirmed."),
          ("Weeks 2–5", "You complete the spec. We inspect and sign off."),
          ("Go-live", "Relisted, repriced, live. The six months begin.")]
    sw = (W - 2 * M - 3 * 0.25) / 4
    for i, (k, v) in enumerate(tl):
        x = M + i * (sw + 0.25)
        d.card(s, x, 2.75, sw, 2.9, fill="#FFFFFF" if i == 3 else "card", glass=i < 3, radius=0.08)
        d.pill(s, x + 0.3, 3.05, 1.35, 0.32, "coral" if i == 3 else "#FFFFFF", k, color="white" if i == 3 else "#0B1020",
               size=9.5)
        T(s, x + 0.3, 3.65, sw - 0.6, 1.8, v, size=12.5, color="#0B1020" if i == 3 else "body", leading=1.45)

    # 16 — Close (aurora) -----------------------------------------------------------------------------------------------
    s = d.slide(bg="aurora", chrome=False, transition=("fade", True))
    gradient_mask_image(s, 0, 0, W, H, bg_color=DARK["bg"], direction="right", alpha_start=75,
                        alpha_end=10).name = "static-mask"
    d.logo(s, M - 0.05, 0.6, 0.95)
    eb(s, M, 1.85, "Red Flag Managed", color="amber")
    H1(s, M, 2.2, 8.5, 1.6, "From $2.16 a night,\nwith three properties.", size=52)
    T(s, M, 3.95, 8, 0.4, "8% for life.  Nights in writing.", size=20, color="amber", bold=True)
    T(s, M, 4.65, 7.2, 1.0, "Send us your listing links. We’ll assess them free and tell you whether they can hit the "
      "number — before you pay anything. If they can’t, we’ll say so.", size=14, color="white", leading=1.45)
    d.pill(s, M, 5.85, 4.6, 0.5, "#FFFFFF", "redflaghomes.in  ·  hello@redflaghomes.in", color="#0B1020", size=11)
    T(s, M, 6.65, 11.8, 0.5, "Figures are illustrative and depend on rate, season, market and your own operating costs. We "
      "guarantee a number of nights, not income. The Nights Promise is subject to its terms. This is a property "
      "management agreement, not a franchise.", size=8.5, color="body", leading=1.3)

    # 17–19 — FAQ ---------------------------------------------------------------------------------------------------------
    faq = [
        ("What exactly are you guaranteeing?", "A number of nights, not an amount of money. 108 qualifying nights per property over six months — an average of 18 a month. The number for your property is written into your agreement."),
        ("What if you miss it?", "We refund your onboarding fee, pro-rated across the properties that fell short. Three properties enrolled and one short means $233 back."),
        ("Do I have to file a claim?", "No. We count the nights from our own booking records and check at month six whether you ask or not. If we’re short, the money goes back to your account within 30 days."),
        ("Why $2.16 a night for some owners and $6.47 for others?", "The fee is the same $699 whether you enroll one property or three. One property buys 108 promised nights; three buys 324 — so it’s $6.47 a night with one property and $2.16 a night with three properties."),
        ("Is this a franchise?", "No. This is a property management agreement. We don’t license you our brand, we don’t put our name on your property, and you aren’t buying a business. You’re hiring a manager."),
        ("Who holds the listing?", "You do. It stays in your name with your reviews, your history and your account. We manage it for you. If you leave, nothing transfers."),
        ("Why is your fee so much lower than everyone else’s?", "Most full-service managers run on payroll. Our operations run on an AI system, so each additional property costs us very little. That’s the whole difference."),
        ("What about cleaning?", "Cleaning is charged to the guest as a cleaning fee, the way it works on every major platform. We schedule it and manage the cleaners."),
        ("What do I still pay for?", "Utilities, supplies, insurance, your mortgage or rent, major repairs, and your own taxes. Also anything on the setup spec."),
        ("What about permits and HOA rules?", "Your property must be legally permitted for short-term rental in your city, and your HOA or condo rules must allow it. We’ll ask you to confirm both in writing before we start — and we’ll tell you if we think your city makes it unworkable."),
        ("What about taxes?", "Lodging and occupancy taxes are collected by the platforms in many markets, but not all. Income tax is yours. We’ll need a W-9 before your first payout and may issue a 1099. Talk to your CPA."),
        ("Do I need special insurance?", "Yes. A standard homeowner policy usually excludes paying guests. You’ll need a short-term rental endorsement or policy, and we ask to be named as an additional insured where your carrier allows."),
        ("Can you just drop my rate to hit the number?", "No. We agree a floor rate with you at onboarding, and nights sold below it don’t count toward the promise. It protects you and it stops us gaming our own guarantee."),
        ("Can I still use my place?", "Yes — up to 12 owner nights across the six months, blocked in advance. Beyond that the promise no longer applies, because we can’t sell nights that aren’t available."),
        ("How do I get paid?", "Twice a month, with a statement showing every booking and every deduction."),
        ("Can I cancel?", "Yes, on 30 days’ notice, honoring any confirmed bookings. If you cancel inside the first six months, the promise lapses."),
        ("What if my market is seasonal?", "We set your number from your market, not a template. A beach property might be promised fewer nights than an urban one. We’d rather promise a number we can hit."),
        ("What if you turn my property down?", "We’ll tell you why, in writing, and you pay nothing. We’d rather decline than take a fee for nights we can’t deliver."),
        ("Is my income guaranteed?", "No. We guarantee nights, not dollars. What those nights earn depends on your rate, your season and your market."),
    ]
    for page, chunk in enumerate((faq[:7], faq[7:13], faq[13:])):
        s = d.slide(dark=False)
        eb(s, M, 1.0, f"Questions owners ask  ·  {page + 1} of 3")
        H1(s, M, 1.3, 10, 0.7, ["The guarantee.", "Running costs, taxes and rules.", "Using, leaving and earning."][page],
           size=30)
        fw = (W - 2 * M - 0.3) / 2
        for i, (q, a) in enumerate(chunk):
            x = M + (i % 2) * (fw + 0.3)
            y = 2.1 + (i // 2) * 1.22
            d.card(s, x, y, fw, 1.12, radius=0.08)
            T(s, x + 0.25, y + 0.14, fw - 0.5, 0.3, q, size=11.5, color="ink", bold=True)
            T(s, x + 0.25, y + 0.44, fw - 0.5, 0.62, a, size=8.8, color="body", leading=1.28)

    assert len(d.prs.slides) == TOTAL, len(d.prs.slides)
    d.save()


if __name__ == "__main__":
    build()
