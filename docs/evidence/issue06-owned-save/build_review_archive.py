#!/usr/bin/env python3
"""Curate only explicitly selected owned-save evidence; never archive an artifact tree."""
import argparse
import hashlib
import io
import json
from pathlib import Path
import re
import zipfile
from pypdf import PdfReader

HEAD = '0ff9db5eb6a9baaeaa57497534fabf07f0e08848'
PAIR = ['0e2452473f47c8aa0f916423be6362276ad15666', 'f7e3a62b6e63522a836643e59ff8144162ab5bfa']
FONT = {'latin': ('b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5', 569208, 'licenses/NotoSans-LICENSE'),
        'cjk': ('3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa', 21659544, 'licenses/Ofdrw-CI-Noto-OFL.txt')}
SUBSETS = {'latin-input': ('8f5e1bad5b4980d66e541c39847f1f36f651491cd62aabe3cb503a6463881672', 34640),
           'latin-output': ('3573b2ed42d211e77a4eef50388ab0c833bdd0cfad4e15a8dff5c06b6c8bb600', 34580),
           'cjk-input': ('13ec0b3af15c206d44fdde4801f6820fe39b86b142c5995b4952907f4be4c102', 524668),
           'cjk-output': ('e0bbdf7cddcd724b7fb5fb73fd95d9f5ae81c3d327747581fa6839d03f0f0d7b', 524624)}
PINGFANG = '876af2cd4854644e7f3e7feb2f688997fdb3343c6df6693611209c9dfb47ccec'


def sha(data): return hashlib.sha256(data).hexdigest()
def encoded(value): return (json.dumps(value, indent=2, ensure_ascii=False) + '\n').encode()


p = argparse.ArgumentParser()
p.add_argument('--repo', type=Path, default=Path(__file__).resolve().parents[3])
p.add_argument('--output', type=Path, default=Path(__file__).resolve().parent)
p.add_argument('--generate', action='store_true', help='Generate the review ZIP after curating its explicit inventory.')
a = p.parse_args(); root = a.repo.resolve(); out = a.output.resolve()
b = Path('artifacts/issue06-q3-implementation/option-b-integration')
rows = []; payloads = {}; font_audit = []


def add(source, member, license_id='MIT project evidence', role='original evidence'):
    source = Path(source); data = (root / source).read_bytes()
    if member in payloads or source.suffix.lower() in {'.dll', '.nupkg', '.snupkg', '.ttf', '.otf', '.ttc'}:
        raise ValueError('Duplicate/forbidden raw archive member: ' + str(source))
    if any(x in source.parts for x in ['bin', 'obj', 'package-feed', 'public-consumer-packages']):
        raise ValueError('Excluded runtime/cache tree: ' + str(source))
    payloads[member] = data
    rows.append({'path': member, 'bytes': len(data), 'sha256': sha(data), 'source': source.as_posix(), 'license': license_id, 'role': role})
    return data


# Actual merge identity and existing source-pair/DLL evidence, not binary payloads.
for name in ['actual-merge-full-source-binding.json', 'actual-merge-source-binding.json', 'actual-merge-cli-build.json',
             'actual-merge-cli-build.log', 'attempt6-source-source-binding.json', 'attempt6-cli-build-command.json',
             'attempt6-diagnostic-dll-binding.json', 'attempt6-runtime-binary-snapshot.json']:
    add(b / name, 'identity/' + name)
merge = json.loads((root / b / 'actual-merge-full-source-binding.json').read_text())
assert merge['head'] == HEAD and merge['candidate_head06'] == PAIR[0] and merge['candidate_head09'] == PAIR[1]
assert len(merge['compared_source_script_fixture_files']) == 306 and not merge['errors']
add('LICENSE', 'licenses/PROJECT-MIT.txt', 'MIT')
add('THIRD-PARTY-NOTICES.md', 'licenses/THIRD-PARTY-NOTICES.md', 'project third-party notices')
for name in ['NOTICE.md', 'manifest.json', 'bundle.json', 'README.md']:
    add(Path('tests/fixtures/trusted-fonts') / name, 'font-provenance/' + name, 'OFL-1.1 font provenance')
for name in ['NotoSans-LICENSE', 'Ofdrw-CI-Noto-OFL.txt']:
    add(Path('tests/fixtures/trusted-fonts/licenses') / name, 'licenses/' + name, 'OFL-1.1')

# Six original CLI runs: only four licensed changed-job OFDs, never diagnostic ZIP outputs.
for family in ['latin', 'cjk']:
    for mode in ['noop-identical', 'noop-nohit', 'changed']:
        for name in ['stdout.log', 'stderr.log']:
            add(b / 'attempt6-cli-runtime' / (family + '-' + mode) / name, 'cli/' + family + '-' + mode + '/' + name)
        for name in ['numeric-report.json', 'stdout.log', 'stderr.log']:
            add(b / 'attempt6-owned-runtime' / (family + '-' + mode) / name, 'typed/' + family + '-' + mode + '/' + name)
for name in ['results-corrected.json', 'metadata-correction.json']:
    add(b / 'attempt6-cli-runtime' / name, 'cli/' + name)
add(b / 'attempt6-owned-runtime/results.json', 'typed/results.json')


def scan_ofd(data, expected, owner, depth=0):
    if depth > 8: raise ValueError('Unexpected nested ZIP depth.')
    count = 0
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for entry in z.infolist():
            value = z.read(entry.filename)
            if value[:4] in [b'\x00\x01\x00\x00', b'OTTO', b'ttcf', b'wOFF', b'wOF2'] or entry.filename.lower().endswith(('.ttf', '.otf', '.ttc', '.woff', '.woff2')):
                h = sha(value)
                if (h, len(value)) != expected[:2] or h == PINGFANG: raise ValueError('Unapproved/commercial embedded font: ' + owner + ':' + entry.filename)
                font_audit.append({'owner': owner, 'entry': entry.filename, 'payload_sha256': h, 'payload_bytes': len(value),
                    'source_font_sha256': expected[0], 'license': 'OFL-1.1', 'license_member': expected[2], 'kind': 'unchanged complete original font'})
                count += 1
            elif entry.filename.lower().endswith('.ofd') or value[:4] == b'PK\x03\x04':
                count += scan_ofd(value, expected, owner + ':' + entry.filename, depth + 1)
    return count


for family in ['latin', 'cjk']:
    for side in ['input', 'output']:
        member = 'documents/' + family + '-' + side + '.ofd'
        data = add(b / 'attempt6-cli-runtime' / (family + '-changed') / (side + '.ofd'), member, 'OFL-1.1 embedded font; MIT sample document')
        assert scan_ofd(data, FONT[family], member) == 1
        pdf_name = family + '-' + side
        member = 'documents/' + pdf_name + '.pdf'
        data = add(b / 'attempt6-visual' / (pdf_name + '.pdf'), member, 'OFL-1.1 derived font subset; MIT sample rendering')
        doc = PdfReader(io.BytesIO(data)); fonts = 0
        for page in doc.pages:
            for key, value in page['/Resources']['/Font'].items():
                f = value.get_object(); name = str(f['/BaseFont'])
                assert re.fullmatch('/[A-Z]{6}\\+Ofdrw-' + FONT[family][0], name), name
                desc = f['/DescendantFonts'][0].get_object()['/FontDescriptor'].get_object()
                assert '/FontFile2' in desc and '/FontFile' not in desc and '/FontFile3' not in desc
                raw = desc['/FontFile2'].get_object().get_data()
                assert (sha(raw), len(raw)) == SUBSETS[pdf_name]
                font_audit.append({'owner': member, 'resource': str(key), 'base_font': name, 'payload_sha256': sha(raw), 'payload_bytes': len(raw),
                    'source_ofd_member': 'documents/' + pdf_name + '.ofd', 'source_font_sha256': FONT[family][0], 'source_font_bytes': FONT[family][1],
                    'license': 'OFL-1.1', 'license_member': FONT[family][2], 'kind': 'actual PDF subset; different bytes from the complete source font'})
                fonts += 1
        assert fonts == 1

# Public sample DOCX must contain XML/relationship members only, no hidden font/binary.
docx_member = 'sample/generated-layout.docx'
data = add('e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx', docx_member, 'MIT original public deterministic sample')
with zipfile.ZipFile(io.BytesIO(data)) as z:
    docx_entries = z.namelist()
    assert all(not x.lower().startswith(('word/fonts/', 'word/embeddings/')) and x.lower().endswith(('.xml', '.rels')) for x in docx_entries)

# Bitmap-only native/default evidence; corresponding commercial-font binaries remain local.
for name in ['latin-input-1.png', 'latin-output-1.png', 'cjk-input-1.png', 'cjk-output-1.png',
             'generated-layout-native-1.png', 'generated-layout-native-2.png', 'generated-layout-default-1.png', 'generated-layout-default-2.png']:
    add(b / 'attempt6-visual' / name, 'pages/' + name, 'MIT sample bitmap rendering; no embedded font software')
for name in ['visual-acceptance2.json', 'text-readback.json', 'preparation.json',
             'latin-input-pdf-text.txt', 'latin-output-pdf-text.txt', 'cjk-input-pdf-text.txt', 'cjk-output-pdf-text.txt',
             'generated-layout-native-ofd-text.txt', 'generated-layout-default-ofd-text.txt',
             'generated-layout-native-pdf-text.txt', 'generated-layout-default-pdf-text.txt']:
    add(b / 'attempt6-visual' / name, 'visual/' + name)

# Exact full functional/TFM/font/package/oracle records; no local feed/nupkg/DLL trees.
low = b / 'attempt6-independent'
for name in ['summary.json', 'test-summary.json', 'test-env.json', 'test-commands.json', 'fonts-after.json',
             'after-source-check3.json', 'before-source-check3.json', 'copy-binding-after.json', 'prior-runtime-audit-after.json',
             'package-audit.json', 'package-command.json', 'package-e2e.log', 'public-consumer-commands.json',
             'consumer-restore.log', 'consumer-build.log', 'consumer-latin.log', 'consumer-cjk.log']:
    add(low / name, 'verification/' + name)
add(low / 'package-output/package-manifest.json', 'verification/package-manifest.json')
for name in ['Core', 'Packaging', 'Cli', 'Converter.Docx', 'Converter.Pdf', 'Signatures']:
    add(low / 'test-results' / (name + '.trx'), 'verification/trx/' + name + '.trx')
    add(low / 'test-logs' / ('test-' + name + '.log'), 'verification/logs/test-' + name + '.log')
for name in ['Reader', 'Layout', 'Converter.Pdf']:
    add(low / 'test-logs' / ('build-' + name + '.log'), 'verification/logs/build-' + name + '.log')
for name in ['Program.cs', 'PublicOutputVerifier.csproj']:
    add(low / 'public-consumer' / name, 'public-consumer/' + name)

# Eleven failures: only logs/typed numbers/state identity, not repeated OFDs/TTFs/temps.
failure = b / 'attempt6-failure-runtime1'
for name in ['results.json', 'evidence.json', 'notes.md']:
    add(failure / name, 'failures/' + name)
for case in ['same-path', 'output-symlink-input', 'wrong-trusted-sha', 'replacement-overflow', 'directory-target', 'late-input-marker']:
    for mode in (['typed'] if case == 'late-input-marker' else ['unmodified', 'typed']):
        folder = case + '-' + mode
        for name in ['stdout.log', 'stderr.log']:
            add(failure / folder / name, 'failures/' + folder + '/' + name)
        if mode == 'typed': add(failure / folder / 'numeric-report.json', 'failures/' + folder + '/numeric-report.json')
add(b / 'attempt6-failure-prep/run_failure_matrix.py', 'failures/run_failure_matrix.py')

excluded = []
for name in ['generated-layout-native.ofd', 'generated-layout-default.ofd', 'generated-layout-native.pdf', 'generated-layout-default.pdf']:
    path = b / 'attempt6-visual' / name; raw = (root / path).read_bytes()
    excluded.append({'source': path.as_posix(), 'bytes': len(raw), 'sha256': sha(raw), 'reason': 'LOCAL ONLY: native/default chain embeds commercial PingFang SC font software; public archive includes bitmap/text/Preview records only'})
font_map = {'status': 'exact archive font payload inventory and adjacent notices', 'raw_fonts': FONT,
    'payloads': font_audit, 'docx': {'member': docx_member, 'entries': docx_entries, 'embedded_font_or_binary': False},
    'commercial_exclusion': {'font': 'PingFang SC', 'bytes': 23278008, 'sha256': PINGFANG, 'binaries': excluded},
    'policy': 'Reject every font payload outside the two exact original OFL hashes and four pinned derived PDF subsets. Thus PingFang, old Arial Unicode and all unknown embedded payloads are excluded. Full TTFs/bundle ZIP are not duplicated; existing repository tests/fixtures/trusted-fonts bundle remains separately referenced.'}
notes = '''Owned-save curated review evidence\n\nActual product-source baseline:0ff9db5eb6a9baaeaa57497534fabf07f0e08848;306 compared source/script/fixture members are byteexact to validated owner06 0e24524 + owner09 f7e3a62. Later documentation/CI commits do not create new runtime evidence at the same Git HEAD. The archive preserves original logs/TRX/scalars/source/DLL hashes,not DLLs or local packages. Runtime and Low records retain their original source-pair identity.\n\nFour documents/*.ofd are ORIGINAL UNMODIFIED CLI inputs/changed outputs. Four PDFs derive from those same OFDs and contain the mapped OFL subsets; subset byte hashes differ from full-font source hashes. The diagnostic runtime uses independently timestamped ZIP outputs,so its original output hashbad622... is never substituted for the unmodified CLI output22a127... . Both records remain clearly identified.\n\nEight bitmap pages, original public DOCX with no font/binary ZIP member, text readback and actual Preview six-PDF/eight-page record are included. TWO native/default OFDs and TWO corresponding PDFs are LOCAL ONLY due to commercial PingFang SC payload; exclusion inventory retains path/hash/size. Original logs/Preview records may refer to those excluded local paths and machine-specific source/test paths, but the binaries/private contents are absent. These bitmap/record files are not a redistributable copy of the commercial font software.\n\nOnly actual checked samples/pages are accepted by the original visual record. Complex-script/layout support is not generalized from them. Packages remain11 locally bound hashes/manifests/oracle evidence,not a feed or release. Failures are11 separate full-original-CJK jobs,including one native-event observed partial-stage mutation and final Noop raw-input refusal;no generalized cancellation/Changed-late result is invented.\n\nAdjacent PROJECT-MIT/Noto license/NOTICE texts and font-license-map.json accompany every font payload mapping. No font raw files,caches,bin,obj,DLL,nupkg,commercial native/default OFD/PDF,or entire ignored tree is included.\n'''.encode()
for member, data, role in [('font-license-map.json', encoded(font_map), 'font/subset license mapping'), ('ARCHIVE-NOTES.txt', notes, 'curation scope and exclusions')]:
    payloads[member] = data; rows.append({'path': member, 'bytes': len(data), 'sha256': sha(data), 'source': 'generated by build_review_archive.py', 'license': 'MIT project evidence', 'role': role})
assert len(payloads) == len(rows)
manifest = {'status': 'explicit curated public review allowlist; no release/tag', 'actual_product_source_head': HEAD,
    'validated_source_pair': PAIR, 'source_identity_members': 306, 'members': rows, 'member_count': len(rows),
    'uncompressed_bytes': sum(x['bytes'] for x in rows), 'estimated_max_zip_bytes': sum(x['bytes'] for x in rows) + 1024 * len(rows),
    'explicit_exclusions': excluded + [{'source': 'all other artifacts, raw fonts, repository font bundle duplicate, old Arial Unicode/native payloads, cache/feed/nupkg/bin/obj/DLL trees/private material', 'reason': 'Not in the explicit allowlist.'}],
    'archive_identity_note': 'This external inventory lists every ZIP member; it is not itself a ZIP member, avoiding self-hash ambiguity.'}
out.mkdir(parents=True, exist_ok=True)
(out / 'font-license-map.json').write_bytes(encoded(font_map))
(out / 'member-inventory.json').write_bytes(encoded(manifest))
if a.generate:
    target = out / 'review-evidence.zip'
    if target.exists(): raise SystemExit('Refuse replacing an existing reviewed archive.')
    with zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for member, data in payloads.items():
            info = zipfile.ZipInfo(member, (2026, 10, 14, 0, 0, 0)); info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o100644 << 16; z.writestr(info, data)
    with zipfile.ZipFile(target) as z:
        assert z.namelist() == list(payloads) and all(sha(z.read(row['path'])) == row['sha256'] for row in rows)
    (out / 'archive-identity.json').write_bytes(encoded({'file': target.name, 'bytes': target.stat().st_size, 'sha256': sha(target.read_bytes()),
        'member_count': len(rows), 'member_inventory_sha256': sha((out / 'member-inventory.json').read_bytes()),
        'font_license_map_sha256': sha((out / 'font-license-map.json').read_bytes()), 'product_source_baseline': HEAD, 'runtime_source_pair': PAIR}))
    print((out / 'archive-identity.json').read_text())
else:
    print(json.dumps({'member_count': len(rows), 'uncompressed_bytes': manifest['uncompressed_bytes']}))
