#!/usr/bin/env python3
"""Join unmodified consumer receipts; a diagnostic/API pass is not Preview acceptance."""
import argparse
import json
from pathlib import Path


def geometry_matches(expected, actual):
    if len(actual) != len(expected):
        return False
    expected = sorted(expected, key=lambda value: (value['line'], value['left']))
    actual = sorted(actual, key=lambda value: (-value['top'], value['left']))
    rectangles = {(value['left'], value['bottom'], value['right'], value['top']) for value in actual}
    return len(rectangles) == len(actual) and all(
        value['right'] > value['left'] and value['top'] > value['bottom'] and
        abs(value['left'] - target['left']) < 0.1 and abs(value['right'] - target['right']) < 0.1
        for target, value in zip(expected, actual))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('matrix', type=Path)
    parser.add_argument('pdfpig', type=Path)
    parser.add_argument('pdfkit', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    matrix = json.loads(args.matrix.read_text())['cases']
    pigs = {(item['Name'], item['UseActualText']): item for item in json.loads(args.pdfpig.read_text())['Rows']}
    kits = {item['name']: item for item in json.loads(args.pdfkit.read_text())['rows']}
    rows = []
    for case in matrix:
        pig = pigs[(case['name'], True)]
        kit = kits[case['name']]
        assert pig['Sha256'] == kit['sha256'] == case['sha256']
        pig_geometry, kit_geometry = True, True
        for word in case['expectedWords']:
            expected = [value for value in case['expectedGeometry'] if value['text'] == word['text']]
            pig_boxes = [{key.lower(): value[key] for key in ('Left', 'Bottom', 'Right', 'Top')}
                         for value in pig['Words'] if value['Text'] == word['text']]
            pig_geometry &= geometry_matches(expected, pig_boxes)
            find = next(value for value in kit['finds'] if value['query'] == word['text'])
            kit_boxes = [box for match in find['matches'] for box in match['bounds']]
            kit_geometry &= find['count'] == word['count'] and all(box['page'] == 1 for box in kit_boxes)
            kit_geometry &= geometry_matches(expected, kit_boxes)
        reverse_absent = all(item['count'] == 0 for item in kit['finds'] if not item['logical'])
        rows.append({'name': case['name'], 'variant': case['variant'], 'sha256': case['sha256'],
                     'pypdfExact': case['pypdfExactSource'],
                     'pdfpigDefaultExact': pigs[(case['name'], False)]['ExactSource'],
                     'pdfpigActualTextExact': pig['ExactSource'], 'pdfpigWordGeometry': pig_geometry,
                     'pdfkitExact': kit['exactSource'], 'pdfkitDeclaredLineBreaks': kit['matchesDeclaredLineBreaks'],
                     'pdfkitWordGeometry': kit_geometry, 'pdfkitReverseAbsent': reverse_absent,
                     'automatedCandidate': pig['ExactSource'] and pig_geometry and
                         kit['matchesDeclaredLineBreaks'] and kit_geometry and reverse_absent,
                     'preview': 'BLOCKED', 'productAcceptance': False})
    args.output.write_text(json.dumps({'geometryTolerancePoints': 0.1,
        'geometryScope': 'Count, distinct positive rectangles, expected horizontal extent and line order; Preview highlight still required',
        'rows': rows}, ensure_ascii=False, indent=2) + '\n')
    for row in rows:
        print(row['name'], 'Pig:', row['pdfpigActualTextExact'], row['pdfpigWordGeometry'],
              'Kit:', row['pdfkitDeclaredLineBreaks'], row['pdfkitWordGeometry'],
              'candidate:', row['automatedCandidate'])


if __name__ == '__main__':
    main()
