#!/usr/bin/env python3
"""F06 independent RTL reference PDFs (fpdf2 + uharfbuzz text shaping)."""
import hashlib, json, sys, importlib.metadata
from pathlib import Path
from fpdf import FPDF

font = Path(sys.argv[1]); out = Path(sys.argv[2])
font_sha = hashlib.sha256(font.read_bytes()).hexdigest()
assert font_sha == "3362132d725d94e23e6995749627c11fdb197a6a6a95faa068e9b757877a8212"
cases = [
    # name, source, direction, width (None = single-line cell)
    ("arabic-ltr-cell", "abc مرحبا لا def", "ltr", None),
    ("arabic-rtl-cell", "مرحبا بالعالم لا", "rtl", None),
    ("hebrew-rtl-wrap", "שלום עולם abc מקום טוב", "rtl", 30),
    ("hebrew-digits-cell", "abc שלום 123 def", "ltr", None),
    ("hebrew-niqqud-cell", "abc שָׁלוֹם def", "ltr", None),
]
rows = []
for name, source, direction, width in cases:
    pdf = FPDF(unit="mm", format=(100, 40))
    pdf.set_creation_date(__import__("datetime").datetime(2026, 11, 6, tzinfo=__import__("datetime").timezone.utc))
    pdf.add_page()
    pdf.add_font("FixtureDejaVu", fname=str(font))
    pdf.set_font("FixtureDejaVu", size=4 * 72 / 25.4)
    pdf.set_text_shaping(True, direction=direction)
    pdf.c_margin = 0
    pdf.set_xy(10, 8)
    if width is None:
        pdf.cell(w=80, h=6, text=source, align="R" if direction == "rtl" else "L")
    else:
        pdf.multi_cell(w=width, h=6, text=source, align="R")
    path = out / f"fpdf2-{name}.pdf"
    pdf.output(path)
    rows.append({"name": name, "source": source, "direction": direction, "widthMillimeters": width,
                 "sha256": hashlib.sha256(path.read_bytes()).hexdigest(), "bytes": path.stat().st_size})
print(json.dumps({"fontSha256": font_sha,
                  "versions": {n: importlib.metadata.version(n) for n in ("fpdf2", "uharfbuzz", "fonttools")},
                  "cases": rows}, ensure_ascii=False, indent=2))
