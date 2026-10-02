#!/usr/bin/env python3
"""Render every final PDF/SVG page and record native-object, Unicode and file-hash evidence."""
import collections
import hashlib
import json
from pathlib import Path
import platform
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

directory, root = map(lambda value: Path(value).resolve(), sys.argv[1:3])
pages = directory / 'pages'
pages.mkdir(exist_ok=True)
local = lambda node: node.tag.rsplit('}', 1)[-1]
checks = {}
for name in ('graphics', 'graphics-roundtrip', 'baseline-native', 'baseline-default'):
    with zipfile.ZipFile(directory / (name + '.ofd')) as archive:
        assert archive.testzip() is None
        objects, text = [], []
        document = ET.fromstring(archive.read('Doc_0/Document.xml'))
        tree = next(node for node in document if local(node) == 'Pages')
        assert len(tree) == 2
        for page in tree:
            xml = ET.fromstring(archive.read('Doc_0/' + page.attrib['BaseLoc']))
            objects.extend(local(node) for node in xml.iter() if local(node).endswith('Object'))
            text.extend(node.text or '' for node in xml.iter() if local(node) == 'TextCode')
        if name.startswith('graphics'):
            assert set(objects) == {'PathObject', 'TextObject'}
            assert sum(local(node) == 'Font' for node in ET.fromstring(archive.read('Doc_0/PublicRes.xml')).iter()) == 1
        info = subprocess.check_output(['pdfinfo', str(directory / (name + '.pdf'))], text=True)
        assert int(next(line.split(':')[1] for line in info.splitlines() if line.startswith('Pages:'))) == 2
        subprocess.run(['pdftotext', '-raw', str(directory / (name + '.pdf')), str(directory / (name + '.pdf.txt'))], check=True)
        extracted = (directory / (name + '.pdf.txt')).read_text()
        if name.startswith('graphics'):
            # Local bold/italic paint must not duplicate the searchable text layer.
            for phrase in ('No. 2026-004', '192.00', 'EvenOdd', 'NonZero', 'Restored:'):
                assert extracted.count(phrase) == 1, (name, phrase, extracted.count(phrase))
        subprocess.run(['pdftoppm', '-r', '110', '-png', str(directory / (name + '.pdf')), str(pages / name)], check=True)
        for page in range(1, 3):
            subprocess.run(['rsvg-convert', '--background-color', 'white', '--width', '720', '--output', str(pages / f'{name}-svg-{page}.png'), str(directory / f'{name}-{page}.svg')], check=True)
        checks[name] = {'pages': 2, 'objects': dict(collections.Counter(objects)), 'unicode_characters': len(''.join(text)), 'ofd_bytes': (directory / (name + '.ofd')).stat().st_size}
assert (directory / 'graphics.txt').read_text() == (directory / 'graphics-roundtrip.txt').read_text()
def git(*args):
    return subprocess.check_output(['git', '-C', str(root), *args], text=True).strip()
manifest = {'baseline': 'df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9', 'source_head': git('rev-parse', 'HEAD'),
            'source_diff_sha256': hashlib.sha256(subprocess.check_output(['git', '-C', str(root), 'diff', 'HEAD'])).hexdigest(),
            'environment': platform.platform(), 'checks': checks,
            'files': {str(path.relative_to(directory)): {'bytes': path.stat().st_size, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}
                      for path in sorted(directory.rglob('*')) if path.is_file() and path.name != 'manifest.json'}}
(directory / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n')
print(json.dumps(checks, ensure_ascii=False, indent=2))
