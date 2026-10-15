#!/usr/bin/env python3
"""Original MIT XML-boundary glyph fixture; requires developer-only fonttools.

Preserve the existing regular font and append distinct visible rectangle glyphs.
The fixture is a test-symbol font, not a reading face or a product font service.
"""
from copy import deepcopy
from pathlib import Path
from fontTools.ttLib import TTFont
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable
root = Path(__file__).resolve().parents[1]
source = root / 'e2e/Ofdrw.Net.SkiaSharp.E2E/testdata/fonts/regular.ttf'
out = root / 'e2e/Ofdrw.Net.Pdf.Vector.E2E/testdata/xml-boundary.ttf'
font = TTFont(source, recalcTimestamp=False)
base = font.getBestCmap()[ord('A')]
order = list(font.getGlyphOrder())
extra = {}
for code in [9, 10, 13, 31, 0xD7FF, 0xE000, 0xFDD0, 0xFFFD, 0xFFFE, 0xFFFF]:
    name = f'xml{code:04X}'
    font['glyf'][name] = deepcopy(font['glyf'][base])
    font['hmtx'][name] = font['hmtx'][base]
    order.append(name)
    extra[code] = name
font.setGlyphOrder(order)
for table in font['cmap'].tables:
    if table.isUnicode(): table.cmap.update(extra)
wide = CmapSubtable.newSubtable(12)
wide.platformID = 3; wide.platEncID = 10; wide.language = 0
wide.cmap = dict(font.getBestCmap())
wide.cmap.update({0x10000: base, 0x10FFFF: base})
font['cmap'].tables.append(wide)
font['head'].created = font['head'].modified = 3_866_889_600
font.save(out)
print(out)
