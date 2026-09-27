# Visual review (PPTX -> PDF via LibreOffice 24.2 -> PNG via PyMuPDF, 110 dpi)

Round 1 findings -> fixes:
- S1 subtitle collided with title -> title up, subtitle to 5.3in.
- S2 "They're invisible." crowded line 2 -> moved to 4.0in.
- S3 lever bodies misaligned (estimated title height) -> fixed body baseline across columns.
- S5/S6 kicker label wrapped -> widened.
- S8 body text over busy photo -> 45% veil + left gradient.
- S10 "Own the property already?" overlapped sub-line -> explicit positions.

Round 2: all 12 slides re-rendered and inspected; no overflow, overlap or clipping.

| ID | Status | Evidence |
|---|---|---|
| R1 luxury aesthetic | PASS | Black/ivory/champagne system, Instrument Serif display, photography full-bleed |
| R2 logo | PASS | Large on S1 + S12, mark top-right on S2–S11 |
| R3 14 levers | PASS | S3 (01–03), S4 (04–05), S5 (06–09), S6 (10–14) |
| R4 offer/figures | PASS | S8–S12; figures match source doc |
| R5 no defects | PASS | Round 2 PNGs |
| R6 varied architecture | PASS | 9 distinct page architectures |
| R7 editable | PASS | All copy is native text boxes |

Note: fonts Instrument Serif + Inter Tight (Google Fonts) must be installed on the presenting machine;
the ₹ glyph falls back to a system font.
