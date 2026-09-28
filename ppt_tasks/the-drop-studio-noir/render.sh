#!/bin/sh
# PPTX -> PDF -> PNG (LibreOffice + PyMuPDF). Run from the task directory.
set -e
cd "$(dirname "$0")/output"
timeout 300 soffice --headless --norestore --convert-to pdf TheDrop_Studio_Noir.pptx >/dev/null 2>&1
python3 - <<'PY'
import os, pymupdf
d = pymupdf.open("TheDrop_Studio_Noir.pdf"); os.makedirs("png", exist_ok=True)
for i, p in enumerate(d): p.get_pixmap(dpi=110).save(f"png/slide{i+1:02d}.png")
print(len(d), "pages")
PY
