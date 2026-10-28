#!/usr/bin/env python3
"""Build diagnostic A/B/C/D PDFs from a recorded product output, never edit it.

Requires pypdf==6.21.2 and pdftoppm. This parser deliberately accepts only the
bounded E2E producer's PDF grammar; it is not a general PDF rewriter.
"""
import argparse
import hashlib
import json
import re
import subprocess
from pathlib import Path

import pypdf
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ByteStringObject, ContentStream, DictionaryObject, NameObject


def sha(data):
    return hashlib.sha256(data).hexdigest()


def cmap(font):
    font = font.get_object()
    assert font['/Encoding'] == '/Identity-H'
    data = font['/ToUnicode'].get_object().get_data()
    result = {}
    for section in re.findall(rb'beginbfchar(.*?)endbfchar', data, re.S):
        for key, value in re.findall(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', section):
            result[int(key, 16)] = bytes.fromhex(value.decode()).decode('utf-16-be')
    for section in re.findall(rb'beginbfrange(.*?)endbfrange', data, re.S):
        for first, last, value in re.findall(rb'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', section):
            first, last = int(first, 16), int(last, 16)
            raw = bytes.fromhex(value.decode())
            assert last - first < 256
            for code in range(first, last + 1):
                result[code] = (int.from_bytes(raw, 'big') + code - first).to_bytes(len(raw), 'big').decode('utf-16-be')
    assert result
    return result


def source_text(group, maps):
    font = None
    chunks = []
    for operands, operator in group:
        if operator == b'Tf':
            font = str(operands[0])
        elif operator == b'Tj':
            value = operands[0]
            raw = bytes(value) if isinstance(value, bytes) else value.original_bytes
            assert len(raw) % 2 == 0
            chunks.extend(maps[font][int.from_bytes(raw[i:i + 2], 'big')] for i in range(0, len(raw), 2))
        elif operator in (b'TJ', b"'", b'"', b'BDC', b'BMC', b'EMC'):
            raise ValueError('Unexpected source text/marked-content operator')
    return ''.join(chunks)


def groups(operations):
    assert [op for _, op in operations[:2]] == [b'q', b'q']
    assert [op for _, op in operations[-2:]] == [b'Q', b'Q']
    result, start, depth = [], 2, 0
    for index in range(2, len(operations) - 2):
        operator = operations[index][1]
        if depth == 0:
            assert operator == b'q', (index, operator)
            start = index
        depth += (operator == b'q') - (operator == b'Q')
        assert depth >= 0
        if depth == 0:
            result.append(operations[start:index + 1])
    assert depth == 0
    return result


def raster(path, prefix):
    prefix.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(['pdftoppm', '-singlefile', '-r', '144', '-png', str(path), str(prefix)], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    return sha(prefix.with_suffix('.png').read_bytes())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('product_directory', type=Path)
    parser.add_argument('output_directory', type=Path)
    args = parser.parse_args()
    assert pypdf.__version__ == '6.21.2', pypdf.__version__
    product = args.product_directory.resolve()
    output = args.output_directory.resolve()
    output.mkdir(parents=True, exist_ok=True)
    assert product != output
    receipt = json.loads((product / 'r22-mixed-output-receipt.json').read_text())
    hashes = {item['name']: item['sha256'] for item in receipt['outputs']}
    for name, digest in hashes.items():
        assert sha((product / name).read_bytes()) == digest, 'Product input changed: ' + name
    rows = []
    for case in receipt['cases']:
        original = product / (case['name'] + '.pdf')
        original_bytes = original.read_bytes()
        assert sha(original_bytes) == hashes[original.name]
        reader = PdfReader(original)
        assert len(reader.pages) == 1
        page = reader.pages[0]
        operations = ContentStream(page.get_contents(), reader).operations
        units = groups(operations)
        runs = [run for line in case['lines'] for run in line['runs']]
        assert len(units) == len(runs), (case['name'], len(units), len(runs))
        maps = {str(key): cmap(value) for key, value in page['/Resources']['/Font'].items()}
        assert [source_text(unit, maps) for unit in units] == [run['Text'] for run in runs]
        unit_by_start = {run['SourceStart']: unit for run, unit in zip(runs, units)}
        pixel_hash = raster(original, output / 'baseline' / case['name'])
        expected = case['source']
        line_texts = [expected[line['SourceStart']:line['SourceStart'] + line['SourceLength']]
                      for line in case['lines']]
        expected_selection = list(dict.fromkeys([expected, '\n'.join(line.rstrip(' ') for line in line_texts)]))
        hebrew_words = list(dict.fromkeys(run['Text'] for run in runs if any('\u05d0' <= c <= '\u05ea' for c in run['Text'])))
        for label, visual, actual_text in [('A', False, False), ('B', False, True),
                                          ('C', True, False), ('D', True, True)]:
            writer = PdfWriter()
            writer.append_pages_from_reader(reader)
            target_page = writer.pages[0]
            rewritten = list(operations[:2])
            spans = []
            for line in case['lines']:
                ordered = sorted(line['runs'], key=lambda run: run['XMillimeters']) if visual else line['runs']
                for run in ordered:
                    unit = unit_by_start[run['SourceStart']]
                    marked = actual_text and run['Text'] in hebrew_words
                    if marked:
                        raw = b'\xfe\xff' + run['Text'].encode('utf-16-be')
                        rewritten.append(([NameObject('/Span'), DictionaryObject({NameObject('/ActualText'): ByteStringObject(raw)})], b'BDC'))
                        spans.append({'sourceStart': run['SourceStart'], 'text': run['Text'], 'hex': raw.hex()})
                    rewritten.extend(unit)
                    if marked:
                        rewritten.append(([], b'EMC'))
            rewritten.extend(operations[-2:])
            stream = ContentStream(target_page.get_contents(), writer)
            stream.operations = rewritten
            target_page[NameObject('/Contents')] = writer._add_object(stream)
            target = output / (case['name'] + '-' + label + '.pdf')
            with target.open('wb') as handle:
                writer.write(handle)
            reopened = PdfReader(target)
            reread = ContentStream(reopened.pages[0].get_contents(), reopened).operations
            # Check balanced replacement scopes on the actual serialized file.
            depth, found = 0, []
            for operands, operator in reread:
                if operator == b'BDC':
                    depth += 1
                    assert depth == 1 and operands[0] == '/Span'
                    value = operands[1]['/ActualText']
                    raw = bytes(value) if isinstance(value, bytes) else value.original_bytes
                    found.append(raw.hex())
                elif operator == b'EMC':
                    assert depth == 1
                    depth -= 1
            assert depth == 0 and found == [span['hex'] for span in spans]
            # Moving complete balanced units changes no font, glyph or drawing operation.
            assert raster(target, target.with_suffix('')) == pixel_hash, 'Diagnostic changed raster: ' + target.name
            extracted = reopened.pages[0].extract_text()
            rows.append({'name': case['name'] + '-' + label, 'variant': label,
                         'sourcePath': str(original), 'sourceSha256': sha(original_bytes),
                         'path': str(target), 'sha256': sha(target.read_bytes()), 'expected': expected,
                         'expectedSelection': expected_selection,
                         'expectedWords': [{'text': word, 'count': sum(run['Text'] == word for run in runs)} for word in hebrew_words],
                         'expectedGeometry': [{'text': run['Text'], 'sourceStart': run['SourceStart'],
                                               'left': (10 + run['XMillimeters']) * 72 / 25.4,
                                               'right': (10 + run['XMillimeters'] + run['widthMillimeters']) * 72 / 25.4,
                                               'line': index}
                                              for index, line in enumerate(case['lines']) for run in line['runs'] if run['Text'] in hebrew_words],
                         'streamOrder': 'visual' if visual else 'logical', 'actualTextScopes': spans,
                         'pypdfText': extracted, 'pypdfCodepoints': [f'U+{ord(c):04X}' for c in extracted],
                         'pypdfExactSource': extracted == expected, 'rasterSha256': pixel_hash,
                         'preview': 'BLOCKED: GUI recovery and exact-file validation pending'})
        assert original.read_bytes() == original_bytes
    for name, digest in hashes.items():
        assert sha((product / name).read_bytes()) == digest, 'Matrix altered product input: ' + name
    result = {'diagnosticOnly': True, 'sourceSha': receipt['sourceSha'], 'face': receipt['face'],
              'pypdfVersion': pypdf.__version__, 'cases': rows}
    (output / 'matrix.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(f'{len(rows)} diagnostic PDFs; all rasters equal their own product baseline.')


if __name__ == '__main__':
    main()
