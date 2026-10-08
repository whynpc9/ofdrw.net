#!/usr/bin/env python3
"""Generate original MIT synthetic font-style boundaries from the existing narrow fixture.

These shapes are test symbols, not redistributions of a third-party reading face.
Requires developer-only fonttools==4.59.2; no runtime font service is introduced.
"""
from pathlib import Path
from fontTools.ttLib import TTFont

root = Path(__file__).resolve().parents[1]
source = root / 'e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf'
output = root / 'e2e/Ofdrw.Net.SkiaSharp.E2E/testdata/fonts'
output.mkdir(parents=True, exist_ok=True)
variants = {
    'regular': (400, 0x00c0, 0, False, 0),
    'semibold': (600, 0x00c0, 0, False, 10),
    'oblique': (400, 0x0280, 0, True, 0),
    'bold': (700, 0x00a0, 1, False, 20),
    'italic': (400, 0x0081, 2, True, 0),
    'mismatch': (400, 0x00c0, 1, False, 0),
    'missing-os2': (400, 0x00c0, 0, False, 0),
}
for name, (weight, selection, mac, shear, spread) in variants.items():
    font = TTFont(source, recalcTimestamp=False)
    font['OS/2'].version = 4
    font['OS/2'].usWeightClass = weight
    font['OS/2'].fsSelection = selection
    font['head'].macStyle = mac
    if shear or spread:
        for glyphName in font.getGlyphOrder():
            glyph = font['glyf'][glyphName]
            if not hasattr(glyph, 'coordinates'):
                continue
            for i, (x, y) in enumerate(glyph.coordinates):
                glyph.coordinates[i] = (int(x + (0.2 * y if shear else 0) + (spread if x > 250 else -spread)), y)
            glyph.recalcBounds(font['glyf'])
    for record in font['name'].names:
        if record.nameID in (1, 2, 3, 4, 6):
            text = (f'Ofdrw Synthetic {name}' if record.nameID != 6 else f'OfdrwSynthetic-{name}')
            record.string = text.encode(record.getEncoding())
    if name == 'missing-os2':
        del font['OS/2']
    font['head'].created = font['head'].modified = 3_866_889_600
    font.save(output / (name + '.ttf'))
    print(name, hex(selection), mac)

# A deliberately short OS/2 table tests bounded native reads. Keep other
# directory/data offsets intact; checksum mismatch is intentional malformed input.
import struct
from io import BytesIO
short = TTFont(output / 'regular.ttf', recalcTimestamp=False)
for record in short['name'].names:
    if record.nameID in (1, 2, 3, 4, 6):
        record.string = 'OfdrwSynthetic-ShortOS2'.encode(record.getEncoding())
stream = BytesIO(); short.save(stream)
data = bytearray(stream.getvalue())
for index in range(struct.unpack_from('>H', data, 4)[0]):
    entry = 12 + index * 16
    if data[entry:entry+4] == b'OS/2':
        struct.pack_into('>I', data, entry + 12, 63)
        break
(output / 'short-os2.ttf').write_bytes(data)
