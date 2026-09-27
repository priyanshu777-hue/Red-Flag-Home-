# Visual review — 24 slides (PPTX -> PDF LibreOffice 24.2 -> PNG PyMuPDF)

Round 1 -> fixes: cover subtitle collided with title; S2 headline overran body; logo hidden under photos on S2/S6/S23
(also breaks Morph) -> logo moved to top of z-order; S4 shared values now span both columns; kit bullets misaligned;
S14 "Gross bookings" crowded caption; CTA photo too murky -> resort-at-dusk photo on the right.
Round 2: all 24 slides re-inspected, no overflow/overlap/clipping. File size 44 MB -> 2.4 MB (JPEG crops instead of
lossless PNG). Structure: schema order, unique timing ids and animation targets OK on every slide; LibreOffice
re-import sees 24 transitions and all entrance builds.

| ID | Status | Evidence |
|---|---|---|
| R1 first impression / fresh images | PASS | 20 images, none used in the Visibility Booster deck |
| R2 fonts | PASS | Fraunces + Manrope, ₹ drawn in both faces |
| R3 logo | PASS | Cover, CTA, mark on every other slide |
| R4 facts | PASS | Cross-checked fees, Setup Flex, example (₹51,110), table, FAQ 1–18 |
| R5 honesty signals | PASS | 45% column highlighted, "not for you if", disclaimers on S4/S15/S24 |
| R6 defects | PASS | Round 2 PNGs |
| R7 motion | PASS (not played in PowerPoint here) | motion.py timing + transitions |
