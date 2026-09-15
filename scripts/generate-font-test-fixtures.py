#!/usr/bin/env python3
"""Generate original, deterministic test fonts (fonttools is a developer-only dependency).

Both files deliberately share their family/PostScript names. Their glyph widths
and outlines differ, allowing resource/caching isolation to be tested without
redistributing third-party fonts. The shapes are test symbols, not a reading font.
"""
from pathlib import Path
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen

root = Path(__file__).resolve().parents[1]
output = root / 'e2e' / 'Ofdrw.Net.Converter.Pdf.E2E' / 'testdata' / 'fonts'
output.mkdir(parents=True, exist_ok=True)
characters = list(range(32, 127)) + [ord(c) for c in '中文字体测试甲乙']
order = ['.notdef'] + [f'uni{code:04X}' for code in characters]
for variant, advance in [('narrow', 500), ('wide', 850), ('budget', 1000)]:
    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder(order)
    builder.setupCharacterMap({code: f'uni{code:04X}' for code in characters})
    glyphs = {}
    for name in order:
        pen = TTGlyphPen(None)
        if name != 'uni0020':
            pen.moveTo((40, 0))
            pen.lineTo((advance - 40, 0))
            if variant == 'narrow':
                pen.lineTo((advance - 40, 700))
                pen.lineTo((40, 700))
            else:
                pen.lineTo((advance // 2, 700))
            pen.closePath()
        glyphs[name] = pen.glyph()
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics({name: (advance, 40 if name != 'uni0020' else 0) for name in order})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.setupNameTable({
        'familyName': 'Ofdrw Test Face', 'styleName': 'Regular',
        'uniqueFontIdentifier': f'Ofdrw original fixture {variant}',
        'fullName': 'Ofdrw Test Face Regular', 'psName': 'OfdrwTestFace-Regular',
        'version': 'Version 1.0',
        'copyright': 'Original Ofdrw.Net test fixture; distributed under the repository MIT license.'
    })
    builder.setupOS2(sTypoAscender=800, sTypoDescender=-200, usWinAscent=800, usWinDescent=200, fsSelection=0x40)
    builder.setupPost()
    builder.setupMaxp()
    builder.font['head'].created = builder.font['head'].modified = 3_866_889_600
    builder.font.recalcTimestamp = False
    builder.save(output / f'{variant}.ttf')
    print(output / f'{variant}.ttf')
