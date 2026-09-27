"""Build 'The Visibility Booster' deck — Red Flag Homes Network.

Build Mode, pptx-designer public API. Theme lock: theme-lock.yaml v1 ("Private Dossier").
Run from anywhere:  python build_deck.py
"""
from pathlib import Path

from PIL import Image
from pptx.oxml.ns import qn
from pptx.util import Pt

from pptx_designer import Presentation
from pptx_designer.tools.images import cover_image, gradient_mask_image
from pptx_designer.tools.layout import add_slide, clean_save
from pptx_designer.tools.shapes import rect
from pptx_designer.tools.text import text

import motion

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
ASSETS = HERE / "assets"
OUT = HERE / "output"

# ---- Theme lock v1 tokens -------------------------------------------------
C = {
    "background": "#0C0B0A", "surface": "#161412", "ink": "#F3EBDD", "muted": "#A39B8E",
    "hairline": "#3A342D", "champagne": "#C8B48A", "red": "#FF3B30",
    "text_body": "#F3EBDD", "text_dark": "#F3EBDD", "text_muted": "#A39B8E",
    "font_body": "Inter Tight", "font_heading": "Instrument Serif",
}
SERIF, SANS = "Instrument Serif", "Inter Tight"
W, H = 13.333, 7.5
M = 0.8  # outer margin
TOTAL = 14


def prep_assets():
    ASSETS.mkdir(exist_ok=True)
    out = {}
    for key, name in {"villa": "handpicked.jpeg", "guest": "guest.JPEG", "clean": "cleaning.JPEG",
                      "street": "how-we-host.jpg",
                      "lounge": "franchise-asset/scenes/interior-details.webp",
                      "dunes": "franchise-asset/footer/footer-bg.webp",
                      "cabin": "franchise-asset/scenes/sequence/desktop/204.webp"}.items():
        dst = ASSETS / f"{key}.jpg"
        if not dst.exists():
            im = Image.open(ROOT / name).convert("RGB")
            im.thumbnail((2400, 2400))
            im.save(dst, quality=88)
        out[key] = str(dst)
    out["logo"] = str(ROOT / "logobg.png")
    return out


def T(slide, x, y, w, h, txt, size=12, font=SANS, color="ink", italic=False, bold=False,
      spacing=None, align="left", anchor="top", leading=None, caps=False):
    """Native text box via pptx_designer.text, then typographic refinements."""
    box = text(slide, x, y, w, h, txt.upper() if caps else txt, font_size=size, color=C[color],
               bold=bold, align=align, font_name=font, C=C, anchor=anchor)
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


def rule(slide, x, y, w, color="hairline", thick=0.012):
    return rect(slide, x, y, w, thick, C[color], C=C)


def veil(slide, x, y, w, h, alpha_pct):
    """Flat background-colour overlay at the given opacity (keeps photos quiet under type)."""
    shp = rect(slide, x, y, w, h, C["background"], C=C)
    shp.name = "static-veil"
    clr = shp.fill._xPr.find(qn("a:solidFill"))[0]
    clr.append(clr.makeelement(qn("a:alpha"), {"val": str(alpha_pct * 1000)}))
    return shp


def vrule(slide, x, y, h, color="hairline"):
    return rect(slide, x, y, 0.012, h, C[color], C=C)


def label(slide, x, y, w, txt, color="champagne", size=9, align="left"):
    return T(slide, x, y, w, 0.3, txt, size=size, color=color, spacing=300, caps=True, align=align)


def logo(slide, A, x, y, size):
    # "!!" prefix pairs the logo across slides so Morph glides it between positions.
    pic = slide.shapes.add_picture(A["logo"], _in(x), _in(y), _in(size), _in(size))
    pic.name = "!!logo"
    return pic


def photo(slide, x, y, w, h, path):
    pic = cover_image(slide, x, y, w, h, path)
    pic.name = "static-photo"
    return pic


def mask(slide, x, y, w, h, direction, alpha_start=100, alpha_end=0):
    shp = gradient_mask_image(slide, x, y, w, h, bg_color=C["background"], direction=direction,
                              alpha_start=alpha_start, alpha_end=alpha_end)
    shp.name = "static-mask"
    return shp


def chrome(shape):
    shape.name = "chrome"
    return shape


def _in(v):
    from pptx.util import Inches
    return Inches(v)


def base(prs, A, mark=True):
    s = add_slide(prs)
    n = len(prs.slides)
    rect(s, 0, 0, W, H, C["background"], C=C).name = "static-bg"
    if mark:
        logo(s, A, W - M - 0.62, 0.36, 0.62)
        chrome(label(s, M, 0.55, 5, "The Visibility Booster", color="muted", size=8))
        chrome(T(s, W - M - 1.5, H - 0.55, 1.5, 0.25, f"{n:02d} / {TOTAL:02d}", size=8, color="muted",
                 spacing=200, align="right"))
        chrome(label(s, M, H - 0.55, 6, "Red Flag Homes Network", color="muted", size=8))
    return s


def divider(prs, A, numeral, kicker, title, sub, img, veil_pct=30):
    s = base(prs, A, mark=False)
    photo(s, 0, 0, W, H, img)
    veil(s, 0, 0, W, H, veil_pct)
    mask(s, 0, 0, 9.0, H, "right")
    logo(s, A, W - M - 0.62, 0.36, 0.62)
    T(s, M - 0.05, 0.35, 7, 3.9, numeral, size=250, font=SERIF, color="champagne", leading=0.8)
    label(s, M, 4.35, 8, kicker)
    T(s, M, 4.7, 8.5, 1.0, title, size=54, font=SERIF, leading=0.9)
    rule(s, M, 5.85, 3.2, "champagne")
    T(s, M, 6.0, 6.5, 0.8, sub, size=14, color="muted", leading=1.3)
    return s


def lever(s, x, y, w, num, title, body, title_size=22, body_size=11, num_size=40):
    T(s, x, y, 1.2, 0.7, num, size=num_size, font=SERIF, color="red")
    rule(s, x, y + 0.82, w, "hairline")
    T(s, x, y + 0.98, w, 0.9, title, size=title_size, font=SERIF, leading=0.95)
    T(s, x, y + 2.0, w, 2.0, body, size=body_size, color="muted", leading=1.3)


def build():
    A = prep_assets()
    prs = Presentation()
    prs.slide_width, prs.slide_height = _in(W), _in(H)

    # 1 — Cover ----------------------------------------------------------------
    s = base(prs, A, mark=False)
    photo(s, 5.6, 0, W - 5.6, H, A["villa"])
    mask(s, 5.6, 0, 2.2, H, "right", 100, 0)
    logo(s, A, M - 0.1, 0.55, 1.55)
    label(s, M, 2.55, 5, "Red Flag Homes Network  ·  Host Dossier")
    T(s, M, 2.85, 5.4, 2.4, "The Visibility\nBooster", size=74, font=SERIF, leading=0.88)
    T(s, M, 5.3, 4.6, 0.9, "14 things that get your Airbnb seen —\nand booked", size=24,
      font=SERIF, italic=True, color="champagne", leading=1.0)
    rule(s, M, H - 0.85, 4.3, "hairline")
    label(s, M, H - 0.7, 5, "redflaghomes.in", color="muted", size=8)

    # 2 — Thesis ---------------------------------------------------------------
    s = base(prs, A)
    label(s, M, 1.7, 6, "The premise")
    T(s, M, 2.1, 8.4, 2.6, "Most empty listings\naren’t bad.", size=66, font=SERIF, leading=0.9)
    T(s, M, 4.0, 8.4, 1.3, "They’re invisible.", size=66, font=SERIF, italic=True, color="red")
    vrule(s, 9.55, 2.2, 3.4)
    T(s, 9.9, 2.2, 2.65, 2.4,
      "Airbnb doesn’t show guests the best property. It shows the property most likely to get "
      "booked and reviewed well.", size=13, leading=1.3)
    T(s, 9.9, 3.75, 2.65, 2.0,
      "Those are different things — and the gap between them is where your occupancy goes.",
      size=13, color="muted", leading=1.3)
    rule(s, M, 5.95, W - 2 * M)
    T(s, M, 6.1, 11, 0.4, "Everything that follows is a lever you control. Work down the list in order — "
      "the first five move the most.", size=11, color="champagne", italic=False)

    # 3 — Divider I -----------------------------------------------------------------
    divider(prs, A, "I", "Part one  ·  Levers 01–05", "The five that matter most.",
            "They cost nothing — and they move occupancy faster than any amount of new furniture.",
            A["lounge"])

    # 3 — Levers 01–03 ---------------------------------------------------------
    s = base(prs, A)
    label(s, M, 1.25, 6, "I  ·  The five that matter most")
    T(s, M, 1.55, 9, 0.8, "Start where the algorithm looks first.", size=34, font=SERIF)
    cw, g = 3.55, 0.37
    items = [
        ("01", "Reply within an hour, every time",
         "Response rate and speed feed directly into ranking — and guests book whoever answers first. "
         "Set saved replies and phone notifications. Falling below a 90% response rate quietly buries a listing."),
        ("02", "Turn on Instant Book",
         "Listings that need approval rank lower and lose impulse bookings — most weekend bookings. "
         "Keep control by requiring verified ID and setting clear house rules instead."),
        ("03", "Update your calendar every few days",
         "Airbnb favours listings that look actively managed. Even opening the app and confirming "
         "availability signals the listing is alive. A calendar untouched for weeks is treated as stale."),
    ]
    for i, (n, t, b) in enumerate(items):
        lever(s, M + i * (cw + g), 2.7, cw, n, t, b, title_size=24, body_size=12.5, num_size=54)

    # 4 — Levers 04–05 ---------------------------------------------------------
    s = base(prs, A)
    photo(s, 0, 0, 4.7, H, A["guest"])
    mask(s, 2.9, 0, 1.8, H, "left", 100, 0)
    x1, x2, cw = 5.15, 9.25, 3.3
    label(s, x1, 1.25, 6, "I  ·  The five that matter most")
    T(s, x1, 1.75, 1, 0.6, "04", size=26, font=SERIF, color="red")
    T(s, x1, 2.35, cw, 1.4, "80%", size=88, font=SERIF, leading=0.85)
    T(s, x1, 3.6, cw, 0.4, "of the decision is the first image", size=11, color="champagne")
    rule(s, x1, 4.1, cw)
    T(s, x1, 4.3, cw, 0.5, "Fix the first photograph", size=22, font=SERIF)
    T(s, x1, 4.85, cw, 1.6,
      "On a phone it is roughly 80% of the decision. Lead with your single most impressive view — "
      "not the front door, not the lobby. Test a new first image for two weeks and compare views.",
      size=11, color="muted", leading=1.25)
    vrule(s, 8.83, 1.75, 4.7)
    T(s, x2, 1.75, 1, 0.6, "05", size=26, font=SERIF, color="red")
    T(s, x2, 2.35, cw, 1.4, "−25–30%", size=70, font=SERIF, leading=0.85)
    T(s, x2, 3.6, cw, 0.4, "below target for your first ten bookings", size=11, color="champagne")
    rule(s, x2, 4.1, cw)
    T(s, x2, 4.3, cw, 0.5, "Price low, deliberately", size=22, font=SERIF)
    T(s, x2, 4.85, cw, 1.6,
      "A listing with no reviews cannot win at full price. Collect ten reviews fast, then raise in steps. "
      "Wasting the new-listing boost at a price nobody books is the costliest mistake here.",
      size=11, color="muted", leading=1.25)

    # 5 — Divider II ----------------------------------------------------------------
    divider(prs, A, "II", "Part two  ·  Levers 06–14", "Nine more that compound.",
            "Each is small on its own. Together they decide who sees you — and how often.",
            A["dunes"], veil_pct=50)

    # 5 — Levers 06–09 ---------------------------------------------------------
    s = base(prs, A)
    label(s, M, 1.25, 11, "II  ·  Nine more that compound  —  the listing & the calendar")
    T(s, M, 1.55, 11, 0.8, "Be findable in every search you qualify for.", size=34, font=SERIF)
    grid = [
        ("06", "Write a title that says what, why and where",
         "“Sea-view 3BR villa with private pool · 5 min to Ashwem” beats “Beautiful villa in Goa”. "
         "Guests scan titles; specifics stop the scroll."),
        ("07", "Tick every amenity you genuinely have",
         "Guests filter by amenity — a missing tick removes you from entire searches: WiFi speed, parking, "
         "workspace, kitchen, AC, pet-friendly. Never tick one you don’t have."),
        ("08", "Lower your minimum nights",
         "A two-night minimum removes you from every one-night search — in city markets, up to half the "
         "demand. Try one night midweek, keep two at weekends."),
        ("09", "Open your calendar further ahead",
         "Many guests book three to six months out, especially for peak season and weddings. "
         "A calendar open only 60 days ahead is invisible to them."),
    ]
    gw, gh = 5.6, 1.95
    for i, (n, t, b) in enumerate(grid):
        cx = M + (i % 2) * (gw + 0.53)
        cy = 2.65 + (i // 2) * (gh + 0.1)
        rule(s, cx, cy, gw)
        T(s, cx, cy + 0.2, 0.8, 0.6, n, size=30, font=SERIF, color="red")
        T(s, cx + 0.85, cy + 0.22, gw - 0.85, 0.5, t, size=19, font=SERIF)
        T(s, cx + 0.85, cy + 0.72, gw - 0.85, 1.2, b, size=11, color="muted", leading=1.25)

    # 6 — Levers 10–14 ---------------------------------------------------------
    s = base(prs, A)
    label(s, M, 1.25, 11, "II  ·  Nine more that compound  —  reputation & reach")
    rows = [
        ("10", "Add weekly and monthly discounts",
         "Longer stays cut cleaning and vacancy risk, and Airbnb surfaces them to long-stay searchers. Even 10% weekly, 20% monthly."),
        ("11", "Never cancel a booking",
         "Host cancellations suppress ranking for months. Move the guest, refund generously — but don’t cancel."),
        ("12", "Chase reviews in the first 48 hours",
         "Ask in person at check-out and once by message. Early on, review velocity matters more than count."),
        ("13", "List on a second platform",
         "Booking.com fills midweek gaps and reaches a different traveller. Sync calendars to avoid double-booking."),
        ("14", "Build a direct route",
         "A simple page, a WhatsApp number, a link in your Instagram bio. No commission — and yours forever."),
    ]
    rx, rw, rh, ry = M, 8.3, 0.98, 1.7
    for i, (n, t, b) in enumerate(rows):
        y = ry + i * rh
        rule(s, rx, y, rw)
        T(s, rx, y + 0.2, 0.7, 0.6, n, size=26, font=SERIF, color="red")
        T(s, rx + 0.75, y + 0.24, 3.1, 0.7, t, size=17, font=SERIF, leading=0.95)
        T(s, rx + 4.0, y + 0.22, rw - 4.0, 0.75, b, size=10.5, color="muted", leading=1.25)
    rule(s, rx, ry + 5 * rh, rw)
    rect(s, 9.75, 1.7, W - 9.75, 4.9, C["surface"], C=C)
    T(s, 10.15, 2.1, 0.8, 0.8, "“", size=72, font=SERIF, color="red")
    T(s, 10.15, 2.95, 2.5, 2.2, "Ten reviews at 4.9 beats forty at 4.6.", size=30, font=SERIF,
      italic=True, leading=0.95)
    rule(s, 10.15, 4.85, 1.2, "champagne")
    T(s, 10.15, 5.0, 2.4, 1.2, "The only visibility an algorithm can’t change is a direct booking.",
      size=10.5, color="muted", leading=1.25)

    # 7 — This week --------------------------------------------------------------
    s = base(prs, A)
    photo(s, 8.4, 0, W - 8.4, H, A["street"])
    mask(s, 8.4, 0, 1.6, H, "right", 100, 0)
    label(s, M, 1.25, 6, "Your first week")
    T(s, M, 1.55, 7.5, 0.9, "Start with these five,\nthis week.", size=40, font=SERIF, leading=0.92)
    todo = [
        "Turn on Instant Book and set enquiry notifications on your phone.",
        "Swap your first photograph for your single most impressive view.",
        "Tick every honest amenity and drop minimum nights midweek.",
        "Open your calendar six months ahead; add weekly and monthly discounts.",
        "Ask every departing guest for a review — in person, and once by message.",
    ]
    for i, t in enumerate(todo):
        y = 3.05 + i * 0.58
        rect(s, M, y + 0.06, 0.2, 0.2, C["background"], line=C["champagne"], C=C)
        T(s, M + 0.45, y, 6.8, 0.4, t, size=13)
    rule(s, M, 6.1, 7.1)
    T(s, M, 6.22, 7.1, 0.5, "Then change one thing at a time — and compare views and bookings "
      "against the previous fortnight.", size=10.5, color="champagne", italic=False)

    # 8 — Transition -------------------------------------------------------------
    s = base(prs, A, mark=False)
    photo(s, 0, 0, W, H, A["clean"])
    veil(s, 0, 0, W, H, 45)
    mask(s, 0, 0, 8.2, H, "right", 100, 0)
    logo(s, A, W - M - 0.62, 0.36, 0.62)
    label(s, M, 1.9, 6, "Red Flag Outpost Classic", color="red")
    T(s, M, 2.3, 7.4, 2.2, "Or let someone do\nall fourteen for you.", size=58, font=SERIF, leading=0.9)
    T(s, M, 4.55, 5.4, 1.4,
      "Everything in this guide is work — daily pricing, hourly replies, calendars, reviews, cleaning, "
      "repairs. Most owners manage it for three months, then stop. Occupancy follows.",
      size=12.5, color="muted", leading=1.3)
    rule(s, M, 5.95, 5.4, "champagne")
    T(s, M, 6.1, 5.6, 0.8, "We design, brand, launch and run your property completely. "
      "You pay rent and electricity. That’s the list.", size=12.5, leading=1.3)

    # 9 — Comparison -------------------------------------------------------------
    s = base(prs, A)
    label(s, M, 1.25, 6, "The difference")
    T(s, M, 1.55, 9, 0.8, "Your month, before and after.", size=34, font=SERIF)
    c0, c1, c2 = M, 4.1, 8.1
    colw = 4.0
    top = 2.6
    rect(s, c2 - 0.25, top - 0.1, W - M - c2 + 0.25, 4.25, C["surface"], C=C)
    rect(s, c2 - 0.25, top - 0.1, W - M - c2 + 0.25, 0.03, C["red"], C=C)
    label(s, c1, top + 0.08, colw, "Doing it yourself", color="muted")
    label(s, c2, top + 0.08, colw, "Red Flag Outpost Classic", color="red")
    table = [
        ("Guest replies", "You, at 2am", "Us, 24/7"),
        ("Pricing", "You guess", "AI, adjusted daily"),
        ("Cleaning & linen", "You arrange and pay", "We arrange — deducted from bookings at cost"),
        ("Repairs", "You chase them", "Handled up to ₹2,000 per incident"),
        ("Marketing", "None", "Creator launch campaign + direct booking page"),
        ("Your time each month", "10–15 hours", "Close to zero"),
    ]
    for i, (a, b, c) in enumerate(table):
        y = top + 0.5 + i * 0.6
        rule(s, c0, y, W - 2 * M)
        T(s, c0, y + 0.15, 3.2, 0.4, a, size=16, font=SERIF)
        T(s, c1, y + 0.18, 3.7, 0.4, b, size=12, color="muted")
        T(s, c2, y + 0.18, 4.3, 0.4, c, size=12)

    # 10 — What it costs ------------------------------------------------------------
    s = base(prs, A)
    label(s, M, 1.25, 6, "What it costs")
    T(s, M, 1.55, 9, 0.8, "Transparent, from day one.", size=34, font=SERIF)
    tiers = [
        ("Franchise fee", "₹59,999", "one time"),
        ("Setup & procuring", "₹1,11,111", "up to 2 keys"),
        ("Commission", "8%", "of bookings, or ₹4,999 per listing / month — whichever is higher"),
    ]
    tw = 3.6
    for i, (k, v, d) in enumerate(tiers):
        x = M + i * (tw + 0.27)
        rule(s, x, 2.75, tw, "champagne")
        label(s, x, 2.92, tw, k, color="muted")
        T(s, x, 3.3, tw, 1.1, v, size=54, font=SERIF, leading=0.9)
        T(s, x, 4.35, tw - 0.2, 0.8, d, size=11, color="muted", leading=1.25)
    rect(s, M, 5.4, W - 2 * M, 1.05, C["surface"], C=C)
    rect(s, M, 5.4, 0.04, 1.05, C["red"], C=C)
    T(s, M + 0.35, 5.52, 5.2, 0.45, "Own the property already?", size=20, font=SERIF)
    T(s, M + 0.35, 5.98, 5.4, 0.4, "That’s your entire investment — no rent, no deposit.", size=11,
      color="muted")
    T(s, 6.9, 5.55, 5.3, 0.4, "200 founding allocations across India", size=11, color="champagne",
      spacing=100)
    T(s, 6.9, 5.87, 5.4, 0.5, "The first 200 partners keep 8% for life.", size=18, font=SERIF,
      italic=True)
    T(s, M, 6.55, 6, 0.3, "Prices plus applicable GST.", size=8, color="muted")

    # 11 — Earn-Back Promise ----------------------------------------------------------
    s = base(prs, A)
    label(s, M, 1.7, 6, "The Earn-Back Promise", color="red")
    T(s, M, 2.1, 11.5, 1.2, "Earn it back in 12 months.", size=64, font=SERIF, leading=0.9)
    T(s, M, 3.15, 11.5, 1.2, "Or we pay the difference.", size=64, font=SERIF, italic=True,
      color="champagne", leading=0.9)
    rule(s, M, 4.6, W - 2 * M)
    T(s, M, 4.85, 6.3, 1.5,
      "If your Outpost’s net operating profit over its first 12 months after going live is less than "
      "the programme fee you paid, we pay you the difference.", size=13, leading=1.3)
    specs = [("No quibbles", ""), ("Written terms", ""), ("Paid within 30 days", "")]
    for i, (k, _) in enumerate(specs):
        T(s, 7.9, 4.85 + i * 0.45, 4.6, 0.4, "—  " + k, size=15, font=SERIF, color="champagne")
    T(s, M, 6.15, 11, 0.4, "We can promise it because we run the property ourselves, to one standard. "
      "Full conditions in the Programme Terms.", size=9.5, color="muted")

    # 12 — CTA ------------------------------------------------------------------------
    s = base(prs, A, mark=False)
    photo(s, 8.9, 0, W - 8.9, H, A["cabin"])
    mask(s, 8.9, 0, 1.6, H, "right", 100, 0)
    logo(s, A, M - 0.1, 0.6, 1.45)
    label(s, M, 2.45, 6, "Apply for allocation")
    T(s, M, 2.8, 8, 1.1, "redflaghomes.in", size=64, font=SERIF, leading=0.9)
    T(s, M, 3.95, 7.4, 0.9, "We assess your property and send you a projection for it — free, "
      "before you commit anything.", size=14, color="muted", leading=1.3)
    rule(s, M, 5.0, 7.3, "champagne")
    T(s, M, 5.15, 7.4, 0.35, "hello@redflaghomes.in", size=13, color="ink")
    T(s, M, 5.48, 7.4, 0.35, "Red Flag Homes Network, a company of Red Flag World Holdings Ltd, London",
      size=9.5, color="muted")
    T(s, M, 6.2, 7.5, 1.0,
      "General information, not legal, tax or investment advice. Airbnb ranking factors change and are not "
      "published in full; the guidance here reflects widely observed practice. All figures are illustrative "
      "and depend on location, season, occupancy and costs; a property can earn less than its costs. "
      "Prices plus applicable GST. The Earn-Back Promise is subject to the Programme Terms.",
      size=7, color="muted", leading=1.2)

    # Motion (theme lock: motion dial 4/10 — slow fades, no fly-ins, content settles in 2s) -----
    plan = {1: ("fade", True), 2: ("morph", False), 3: ("fade", True), 6: ("fade", True),
            10: ("fade", True), 13: ("fade", True), 14: ("morph", False)}
    for i, slide in enumerate(prs.slides, start=1):
        kind, black = plan.get(i, ("fade", False))
        motion.transition(slide, kind, through_black=black)
        motion.choreograph(slide)

    OUT.mkdir(exist_ok=True)
    path = OUT / "The_Visibility_Booster.pptx"
    clean_save(prs, str(path))
    print(path)


if __name__ == "__main__":
    build()
