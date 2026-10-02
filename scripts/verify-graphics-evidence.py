#!/usr/bin/env python3
"""Render every final PDF/SVG page and record native-object, Unicode and file-hash evidence."""
import collections
import hashlib
import json
from pathlib import Path
import platform
import os
import re
import tempfile
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile

directory, root = map(lambda value: Path(value).resolve(), sys.argv[1:3])
pages = directory / 'pages'
pages.mkdir(exist_ok=True)
local = lambda node: node.tag.rsplit('}', 1)[-1]
checks = {}
fontconfig = directory / 'fontconfig.xml'
config = ET.Element('fontconfig')
ET.SubElement(config, 'dir').text = str(directory / 'fonts')
ET.SubElement(config, 'cachedir').text = str(directory.parent / 'font-cache')
fontconfig.write_text(ET.tostring(config, encoding='unicode'))
svg_environment = dict(os.environ, FONTCONFIG_FILE=str(fontconfig))
for name in ('graphics', 'graphics-roundtrip', 'graphics-name-only', 'baseline-native', 'baseline-default'):
    with zipfile.ZipFile(directory / (name + '.ofd')) as archive:
        assert archive.testzip() is None
        objects, text = [], []
        document = ET.fromstring(archive.read('Doc_0/Document.xml'))
        tree = next(node for node in document if local(node) == 'Pages')
        count = 1 if name == 'graphics-name-only' else 2
        assert len(tree) == count
        for page in tree:
            xml = ET.fromstring(archive.read('Doc_0/' + page.attrib['BaseLoc']))
            objects.extend(local(node) for node in xml.iter() if local(node).endswith('Object'))
            text.extend(node.text or '' for node in xml.iter() if local(node) == 'TextCode')
        if name.startswith('graphics'):
            assert set(objects) == {'PathObject', 'TextObject'}
            assert sum(local(node) == 'Font' for node in ET.fromstring(archive.read('Doc_0/PublicRes.xml')).iter()) == 1
        info = subprocess.check_output(['pdfinfo', str(directory / (name + '.pdf'))], text=True)
        assert int(next(line.split(':')[1] for line in info.splitlines() if line.startswith('Pages:'))) == count
        subprocess.run(['pdftotext', '-raw', str(directory / (name + '.pdf')), str(directory / (name + '.pdf.txt'))], check=True)
        extracted = (directory / (name + '.pdf.txt')).read_text()
        if name in ('graphics', 'graphics-roundtrip'):
            # Local bold/italic paint must not duplicate the searchable text layer.
            for phrase in ('No. 2026-004', '192.00', 'EvenOdd', 'NonZero', 'Restored:'):
                assert extracted.count(phrase) == 1, (name, phrase, extracted.count(phrase))
        subprocess.run(['pdftoppm', '-r', '110', '-png', str(directory / (name + '.pdf')), str(pages / name)], check=True)
        for page in range(1, count + 1):
            subprocess.run(['rsvg-convert', '--background-color', 'white', '--width', '720', '--output', str(pages / f'{name}-svg-{page}.png'), str(directory / f'{name}-{page}.svg')], check=True, env=svg_environment)
        checks[name] = {'pages': count, 'objects': dict(collections.Counter(objects)), 'unicode_characters': len(''.join(text)), 'ofd_bytes': (directory / (name + '.ofd')).stat().st_size}
assert (directory / 'graphics.txt').read_text() == (directory / 'graphics-roundtrip.txt').read_text()
# The acute tip lies outside SVG's default miter-limit=4 bevel. Check actual
# raster ink, so a correct-looking XML attribute alone cannot satisfy the probe.
tip = subprocess.check_output(['identify', '-format', '%[pixel:p{540,375}]', str(pages / 'graphics-svg-2.png')], text=True)
channels = list(map(float, re.findall(r'\d+(?:\.\d+)?', tip)))
assert len(channels) >= 3 and channels[0] < 80 and channels[1] < 130 and channels[2] > 140, ('acute miter tip missing', tip)
with tempfile.TemporaryDirectory(prefix='ofd-miter-control-') as temporary:
    control = ET.parse(directory / 'graphics-2.svg')
    for node in control.iter():
        if local(node) == 'path':
            node.attrib.pop('stroke-miterlimit', None)
    control_svg, control_png = Path(temporary) / 'default.svg', Path(temporary) / 'default.png'
    control.write(control_svg, encoding='utf-8', xml_declaration=True)
    subprocess.run(['rsvg-convert', '--background-color', 'white', '--width', '720', '--output', str(control_png), str(control_svg)], check=True, env=svg_environment)
    old_tip = subprocess.check_output(['identify', '-format', '%[pixel:p{540,375}]', str(control_png)], text=True)
    old_channels = list(map(float, re.findall(r'\d+(?:\.\d+)?', old_tip)))
    assert len(old_channels) >= 3 and all(value >= 240 for value in old_channels[:3]), ('miter negative control did not bevel', old_tip)
def git(*args):
    return subprocess.check_output(['git', '-C', str(root), *args], text=True).strip()
manifest = {'baseline': 'df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9', 'source_head': git('rev-parse', 'HEAD'),
            'source_diff_sha256': hashlib.sha256(subprocess.check_output(['git', '-C', str(root), 'diff', 'HEAD'])).hexdigest(),
            'environment': platform.platform(), 'checks': checks,
            'files': {str(path.relative_to(directory)): {'bytes': path.stat().st_size, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}
                      for path in sorted(directory.rglob('*')) if path.is_file() and path.name != 'manifest.json'}}
(directory / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n')
print(json.dumps(checks, ensure_ascii=False, indent=2))
