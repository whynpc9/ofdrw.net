#!/usr/bin/env python3
"""Diagnostic only: character CID semantics behind unchanged visible OFD outlines.

E/F/G/H vary word-group order (logical/visual) and character stream order
(logical/visual). No ActualText, direction controls, extra text or font-byte
changes are introduced. Inputs are existing product PDFs and corresponding OFDs.
"""
import argparse
import hashlib
import importlib.util
import json
import math
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import fontTools
from fontTools.ttLib import TTFont
import pypdf
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, ByteStringObject, ContentStream,
                          DecodedStreamObject, FloatObject, NameObject, NumberObject)

spec = importlib.util.spec_from_file_location('word_matrix', Path(__file__).with_name('make-matrix.py'))
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)


def local(node):
    return node.tag.rsplit('}', 1)[-1]


def ofd_page(path, page_index):
    with ZipFile(path) as archive:
        names = sorted(name for name in archive.namelist() if name.endswith('/Content.xml'))
        xml = ET.fromstring(archive.read(names[page_index]))
        fonts = [name for name in archive.namelist() if name.endswith('.ttf')]
        assert len(fonts) == 1
        source_font = TTFont(BytesIO(archive.read(fonts[0])))
        rows = []
        for element in xml.iter():
            if local(element) != 'TextObject':
                continue
            codes = [node for node in element.iter() if local(node) == 'TextCode']
            assert len(codes) == 1
            code = codes[0]
            glyphs = next((node for node in element.iter() if local(node) == 'Glyphs'), None)
            rows.append({'text': code.text or '', 'size': float(element.attrib['Size']),
                         'boundary': [float(value) for value in element.attrib['Boundary'].split()],
                         'gids': [int(value) for value in glyphs.text.split()] if glyphs is not None else [],
                         'delta': [float(value) for value in code.attrib.get('DeltaX', '').split()]})
    return rows, source_font


def font_bytes(page):
    result = {}
    for key, value in page['/Resources']['/Font'].items():
        descendant = value.get_object()['/DescendantFonts'][0].get_object()
        data = descendant['/FontDescriptor'].get_object()['/FontFile2'].get_object().get_data()
        result[str(key)] = hashlib.sha256(data).hexdigest()
    return result


def change_font(writer, page, name, word, source_font):
    face = page['/Resources']['/Font'][name].get_object()
    descendant = face['/DescendantFonts'][0].get_object()
    data = descendant['/FontDescriptor'].get_object()['/FontFile2'].get_object().get_data()
    embedded = TTFont(BytesIO(data))
    order = embedded.getGlyphOrder()
    source_order = source_font.getGlyphOrder()
    maps, metrics = {}, []
    for gid, character in zip(word['gids'], reversed(word['text'])):
        assert source_font.getBestCmap()[ord(character)] == source_order[gid]
        outline = embedded['glyf'][order[gid]].getCoordinates(embedded['glyf'])
        original = source_font['glyf'][source_order[gid]].getCoordinates(source_font['glyf'])
        assert all(list(left) == list(right) for left, right in zip(outline, original))
        assert embedded['hmtx'][order[gid]] == source_font['hmtx'][source_order[gid]]
        maps[gid] = character
        width = math.floor(embedded['hmtx'][order[gid]][0] / embedded['head'].unitsPerEm * 1000 + 0.5)
        metrics.append({'gid': gid, 'unicode': f'U+{ord(character):04X}', 'width': width})
    assert len(maps) == len(word['text'])
    lines = [b'/CIDInit /ProcSet findresource begin', b'12 dict begin', b'begincmap',
             b'/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
             b'/CMapName /Adobe-Identity-UCS def /CMapType 2 def',
             b'1 begincodespacerange', b'<0000><FFFF>', b'endcodespacerange',
             f'{len(maps)} beginbfchar'.encode()]
    lines += [f'<{gid:04X}><{ord(character):04X}>'.encode() for gid, character in sorted(maps.items())]
    lines += [b'endbfchar', b'endcmap CMapName currentdict /CMap defineresource pop end end']
    cmap = DecodedStreamObject()
    cmap.set_data(b'\n'.join(lines) + b'\n')
    face[NameObject('/ToUnicode')] = writer._add_object(cmap)
    widths = ArrayObject()
    for metric in sorted(metrics, key=lambda value: value['gid']):
        widths.extend([NumberObject(metric['gid']), ArrayObject([NumberObject(metric['width'])])])
    descendant[NameObject('/W')] = widths
    descendant[NameObject('/CIDToGIDMap')] = NameObject('/Identity')
    return metrics


def change_group(group, word, logical_characters):
    assert len(word['gids']) == len(word['text']) and len(word['delta']) == len(word['text']) - 1
    assert sum(operator == b'Tj' for _, operator in group) == 1
    pen = 0
    positions = []
    for index, gid in enumerate(word['gids']):
        positions.append((gid, pen * 72 / 25.4))
        if index < len(word['delta']):
            pen += word['delta'][index]
    if logical_characters:
        positions.reverse()
    result, origin = [], None
    for operands, operator in group:
        if operator == b'Td':
            origin = [float(value) for value in operands]
        if operator == b'Tj':
            assert origin is not None
            for gid, x in positions:
                result.extend([
                    ([NumberObject(1), NumberObject(0), NumberObject(0), NumberObject(1),
                      FloatObject(origin[0] + x), FloatObject(origin[1])], b'Tm'),
                    ([ByteStringObject(gid.to_bytes(2, 'big'))], b'Tj')])
        else:
            result.append((operands, operator))
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('word_matrix', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--isolated-ofd', required=True, type=Path)
    parser.add_argument('--isolated-pdf', required=True, type=Path)
    args = parser.parse_args()
    assert pypdf.__version__ == '6.21.2'
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    previous = json.loads(args.word_matrix.read_text())
    cases = []
    for row in previous['cases']:
        if row['variant'] != 'A':
            continue
        product = Path(row['sourcePath'])
        product_receipt = json.loads((product.parent / 'r22-mixed-output-receipt.json').read_text())
        item = next(case for case in product_receipt['cases'] if case['name'] == product.stem)
        cases.append((product.stem, product, product.with_suffix('.ofd'), 0, item['lines'], row))
    isolated = PdfReader(args.isolated_pdf)
    assert len(isolated.pages) == 2
    isolated_hash = base.sha(args.isolated_pdf.read_bytes())
    for index, name in enumerate(['isolated-shalom', 'isolated-makom']):
        words, source_font = ofd_page(args.isolated_ofd, index)
        assert len(words) == 1
        word = words[0]
        last_gid = word['gids'][-1]
        width = sum(word['delta']) + source_font['hmtx'][source_font.getGlyphOrder()[last_gid]][0] / source_font['head'].unitsPerEm * word['size']
        left = word['boundary'][0] * 72 / 25.4
        line = {'runs': [{'SourceStart': 0, 'Text': word['text'], 'XMillimeters': 0}]}
        row = {'sourceSha256': isolated_hash, 'expected': word['text'], 'expectedSelection': [word['text']],
               'expectedWords': [{'text': word['text'], 'count': 1}],
               'expectedGeometry': [{'text': word['text'], 'sourceStart': 0, 'left': left,
                                     'right': left + width * 72 / 25.4, 'line': 0}]}
        cases.append((name, args.isolated_pdf.resolve(), args.isolated_ofd.resolve(), index, [line], row))
    rows = []
    for name, pdf, ofd, page_index, lines, previous_row in cases:
        original_bytes, ofd_bytes = pdf.read_bytes(), ofd.read_bytes()
        assert base.sha(original_bytes) == previous_row['sourceSha256']
        source = PdfReader(pdf)
        page = source.pages[page_index]
        original_fonts = font_bytes(page)
        operations = ContentStream(page.get_contents(), source).operations
        units = base.groups(operations)
        words, source_font = ofd_page(ofd, page_index)
        runs = [run for line in lines for run in line['runs']]
        maps = {str(key): base.cmap(value) for key, value in page['/Resources']['/Font'].items()}
        assert len(units) == len(words) == len(runs)
        assert [base.source_text(unit, maps) for unit in units] == [word['text'] for word in words] == [run['Text'] for run in runs]
        baseline = output / 'baseline-pdf' / (name + '.pdf')
        baseline.parent.mkdir(exist_ok=True)
        split = PdfWriter()
        split.add_page(page)
        with baseline.open('wb') as handle:
            split.write(handle)
        pixel_hash = base.raster(baseline, output / 'baseline-raster' / name)
        for label, visual_words, logical_chars in [('E', False, True), ('F', False, False),
                                                   ('G', True, True), ('H', True, False)]:
            writer = PdfWriter()
            writer.add_page(page)
            target_page = writer.pages[0]
            modified, font_receipts = {}, {}
            for run, unit, word in zip(runs, units, words):
                if all('\u05d0' <= char <= '\u05ea' for char in word['text']):
                    font_names = [str(operands[0]) for operands, operator in unit if operator == b'Tf']
                    assert len(font_names) == 1
                    font_name = font_names[0]
                    if font_name not in font_receipts:
                        font_receipts[font_name] = change_font(writer, target_page, font_name, word, source_font)
                    modified[run['SourceStart']] = change_group(unit, word, logical_chars)
                else:
                    modified[run['SourceStart']] = unit
            rewritten = list(operations[:2])
            for line in lines:
                ordered = sorted(line['runs'], key=lambda run: run['XMillimeters']) if visual_words else line['runs']
                for run in ordered:
                    rewritten.extend(modified[run['SourceStart']])
            rewritten.extend(operations[-2:])
            stream = ContentStream(target_page.get_contents(), writer)
            stream.operations = rewritten
            target_page[NameObject('/Contents')] = writer._add_object(stream)
            target = output / (name + '-' + label + '.pdf')
            with target.open('wb') as handle:
                writer.write(handle)
            reopened = PdfReader(target)
            assert font_bytes(reopened.pages[0]) == original_fonts
            assert base.raster(target, target.with_suffix('')) == pixel_hash, target.name
            text = reopened.pages[0].extract_text()
            rows.append({key: previous_row[key] for key in ['expected', 'expectedSelection', 'expectedWords', 'expectedGeometry']} |
                        {'name': name + '-' + label, 'variant': label, 'path': str(target), 'sha256': base.sha(target.read_bytes()),
                         'sourcePath': str(pdf), 'sourceSha256': base.sha(original_bytes), 'sourcePage': page_index + 1,
                         'ofdSha256': base.sha(ofd_bytes), 'streamOrder': 'visual' if visual_words else 'logical',
                         'characterStreamOrder': 'logical' if logical_chars else 'visual',
                         'embeddedFontSha256': original_fonts, 'verifiedOriginalGidMetrics': font_receipts,
                         'rasterSha256': pixel_hash, 'pypdfText': text, 'pypdfExactSource': text == previous_row['expected'],
                         'pypdfCodepoints': [f'U+{ord(char):04X}' for char in text], 'preview': 'BLOCKED'})
        assert pdf.read_bytes() == original_bytes and ofd.read_bytes() == ofd_bytes
    (output / 'matrix.json').write_text(json.dumps({'diagnosticOnly': True, 'pypdfVersion': pypdf.__version__,
        'fontToolsVersion': fontTools.__version__, 'sourceSha': previous['sourceSha'], 'face': previous['face'], 'cases': rows},
        ensure_ascii=False, indent=2) + '\n')
    print(f'{len(rows)} character-semantic diagnostics; original font bytes, outlines, hmtx and pixels preserved.')


if __name__ == '__main__':
    main()
