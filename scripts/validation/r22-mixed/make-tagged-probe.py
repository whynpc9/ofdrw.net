#!/usr/bin/env python3
"""Diagnostic only: add a logical tag tree to existing bounded mixed PDFs.

Retains each product word/space group, font, original GID and position. T-L
keeps logical page stream order; T-V writes groups in visual order while the
tag tree remains logical. Poppler and target readers judge the result.
"""
import argparse
import hashlib
import importlib.util
import json
import subprocess
from pathlib import Path

import pypdf
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, BooleanObject, ContentStream,
                          DictionaryObject, NameObject, NumberObject,
                          TextStringObject)

spec = importlib.util.spec_from_file_location('word_matrix', Path(__file__).with_name('make-matrix.py'))
base = importlib.util.module_from_spec(spec)
spec.loader.exec_module(base)


def embedded_font_hashes(page):
    return {str(name): hashlib.sha256(font.get_object()['/DescendantFonts'][0].get_object()
            ['/FontDescriptor'].get_object()['/FontFile2'].get_object().get_data()).hexdigest()
            for name, font in page['/Resources']['/Font'].items()}


def tag(writer, page, runs_by_line):
    page_reference = page.indirect_reference
    assert page_reference is not None
    page[NameObject('/StructParents')] = NumberObject(0)
    page[NameObject('/Tabs')] = NameObject('/S')
    root = DictionaryObject({NameObject('/Type'): NameObject('/StructTreeRoot')})
    root_ref = writer._add_object(root)
    document = DictionaryObject({NameObject('/Type'): NameObject('/StructElem'),
                                 NameObject('/S'): NameObject('/Document'),
                                 NameObject('/P'): root_ref})
    document_ref = writer._add_object(document)
    root[NameObject('/K')] = ArrayObject([document_ref])
    paragraphs, parents = ArrayObject(), ArrayObject()
    for line in runs_by_line:
        paragraph = DictionaryObject({NameObject('/Type'): NameObject('/StructElem'),
             NameObject('/S'): NameObject('/P'), NameObject('/P'): document_ref,
             NameObject('/Pg'): page_reference, NameObject('/K'): ArrayObject(),
             NameObject('/Lang'): TextStringObject('en-US'),
             NameObject('/A'): DictionaryObject({NameObject('/O'): NameObject('/Layout'),
                                                NameObject('/WritingMode'): NameObject('/LrTb')})})
        paragraph_ref = writer._add_object(paragraph)
        paragraphs.append(paragraph_ref)
        for run in line['runs']:
            marked_id = len(parents)
            span = DictionaryObject({NameObject('/Type'): NameObject('/StructElem'),
                NameObject('/S'): NameObject('/Span'), NameObject('/P'): paragraph_ref,
                NameObject('/Pg'): page_reference, NameObject('/K'): NumberObject(marked_id),
                NameObject('/Lang'): TextStringObject('he-IL' if any('\u05d0' <= character <= '\u05ea'
                                                          for character in run['Text']) else 'en-US')})
            span_ref = writer._add_object(span)
            paragraph['/K'].append(span_ref)
            parents.append(span_ref)
            run['mcid'] = marked_id
    document[NameObject('/K')] = paragraphs
    parent_tree = DictionaryObject({NameObject('/Nums'): ArrayObject([NumberObject(0), parents])})
    root[NameObject('/ParentTree')] = writer._add_object(parent_tree)
    root[NameObject('/ParentTreeNextKey')] = NumberObject(1)
    catalog = writer._root_object
    catalog[NameObject('/StructTreeRoot')] = root_ref
    catalog[NameObject('/MarkInfo')] = DictionaryObject({NameObject('/Marked'): BooleanObject(True)})
    catalog[NameObject('/Lang')] = TextStringObject('en-US')
    return len(parents)


def verify_tree(reader, expected):
    catalog = reader.trailer['/Root']
    assert catalog['/MarkInfo']['/Marked'] and catalog['/Lang'] == 'en-US'
    root = catalog['/StructTreeRoot']
    document_ref = root['/K'][0]
    document = document_ref.get_object()
    assert document['/S'] == '/Document' and document.raw_get('/P') == catalog.raw_get('/StructTreeRoot')
    parent_refs = root['/ParentTree']['/Nums'][1]
    page = reader.pages[0]
    assert page['/StructParents'] == 0 and len(parent_refs) == sum(len(line['runs']) for line in expected)
    assert len(document['/K']) == len(expected)
    observed = []
    for line, paragraph_ref in zip(expected, document['/K']):
        paragraph = paragraph_ref.get_object()
        assert paragraph['/S'] == '/P' and paragraph.raw_get('/P') == document_ref
        assert paragraph['/A']['/WritingMode'] == '/LrTb'
        assert len(paragraph['/K']) == len(line['runs'])
        for run, span_ref in zip(line['runs'], paragraph['/K']):
            span = span_ref.get_object()
            assert span['/S'] == '/Span' and span.raw_get('/P') == paragraph_ref
            marked_id = int(span['/K'])
            assert parent_refs[marked_id] == span_ref
            observed.append((marked_id, run['Text']))
    assert [item[0] for item in observed] == list(range(len(observed)))
    return observed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('word_matrix', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    assert pypdf.__version__ == '6.21.2'
    prior = json.loads(args.word_matrix.read_text())
    by_name = {item['name']: item for item in prior['cases']}
    names = ['r22-mixed-wide', 'r22-mixed-two-word', 'r22-mixed-repeated-word',
             'r22-mixed-leading-rtl', 'r22-mixed-trailing-rtl',
             'r22-mixed-between-rtl-wrap', 'r22-mixed-after-rtl-wrap']
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    rows = []
    for name in names:
        control = by_name[name + '-A']
        product_path = Path(control['sourcePath'])
        product_bytes = product_path.read_bytes()
        assert base.sha(product_bytes) == control['sourceSha256']
        receipt = json.loads((product_path.parent / 'r22-mixed-output-receipt.json').read_text())
        case = next(item for item in receipt['cases'] if item['name'] == name)
        source = PdfReader(product_path)
        assert len(source.pages) == 1
        page = source.pages[0]
        original_fonts = embedded_font_hashes(page)
        operations = ContentStream(page.get_contents(), source).operations
        groups = base.groups(operations)
        runs_by_line = json.loads(json.dumps(case['lines']))
        runs = [run for line in runs_by_line for run in line['runs']]
        maps = {str(key): base.cmap(value) for key, value in page['/Resources']['/Font'].items()}
        assert len(groups) == len(runs)
        assert [base.source_text(group, maps) for group in groups] == [run['Text'] for run in runs]
        for label, visual in [('T-L', False), ('T-V', True)]:
            writer = PdfWriter()
            writer.add_page(page)
            writer.pdf_header = '%PDF-1.7'
            target_page = writer.pages[0]
            count = tag(writer, target_page, runs_by_line)
            by_source = {run['SourceStart']: group for run, group in zip(runs, groups)}
            rewritten = list(operations[:2])
            for line in runs_by_line:
                ordered = sorted(line['runs'], key=lambda run: run['XMillimeters']) if visual else line['runs']
                for run in ordered:
                    props = DictionaryObject({NameObject('/MCID'): NumberObject(run['mcid'])})
                    rewritten.append(([NameObject('/Span'), props], b'BDC'))
                    rewritten.extend(by_source[run['SourceStart']])
                    rewritten.append(([], b'EMC'))
            rewritten.extend(operations[-2:])
            stream = ContentStream(target_page.get_contents(), writer)
            stream.operations = rewritten
            target_page[NameObject('/Contents')] = writer._add_object(stream)
            target = output / f'{name}-{label}.pdf'
            with target.open('wb') as handle:
                writer.write(handle)
            reread = PdfReader(target)
            assert reread.pdf_header == '%PDF-1.7' and reread.trailer['/Root'].get('/Version') is None
            assert embedded_font_hashes(reread.pages[0]) == original_fonts
            expected_tags = verify_tree(reread, runs_by_line)
            assert len(expected_tags) == count
            found = [int(operands[1]['/MCID']) for operands, operator in
                     ContentStream(reread.pages[0].get_contents(), reread).operations if operator == b'BDC']
            assert sorted(found) == list(range(count))
            pixel_hash = base.raster(target, target.with_suffix(''))
            assert pixel_hash == control['rasterSha256']
            structure = subprocess.run(['pdfinfo', '-struct-text', str(target)], text=True,
                                       capture_output=True, check=True).stdout
            structure_path = target.with_suffix('.structure.txt')
            structure_path.write_text(structure)
            rows.append({key: control[key] for key in ('expected','expectedSelection','expectedWords','expectedGeometry')} |
                        {'name': f'{name}-{label}', 'variant': label, 'path': str(target),
                         'sha256': base.sha(target.read_bytes()), 'sourcePath': str(product_path),
                         'sourceSha256': base.sha(product_bytes), 'pdfHeader': reread.pdf_header,
                         'embeddedFontSha256': original_fonts, 'mcids': expected_tags,
                         'popplerStructurePath': str(structure_path), 'popplerStructureRaw': structure,
                         'rasterSha256': pixel_hash,
                         'pypdfText': reread.pages[0].extract_text(),
                         'pypdfExactSource': reread.pages[0].extract_text() == case['source'],
                         'preview': 'BLOCKED'})
        assert product_path.read_bytes() == product_bytes
    (output / 'matrix.json').write_text(json.dumps({'diagnosticOnly': True,
        'pypdfVersion': pypdf.__version__, 'cases': rows}, ensure_ascii=False, indent=2) + '\n')
    print(f'{len(rows)} tagged diagnostic PDFs with structural and pixel checks; no product acceptance.')


if __name__ == '__main__':
    main()
