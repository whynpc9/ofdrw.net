#!/usr/bin/env python3
"""Fail a tag release without a candidate-bound human Preview acceptance record."""
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
RECORD = ROOT / 'docs' / 'release-preview-acceptance.json'


def source_fingerprint():
    paths = subprocess.check_output(['git', 'ls-files', '-z'], cwd=ROOT).split(b'\0')
    digest = hashlib.sha256()
    for raw in sorted(path for path in paths if path):
        name = raw.decode()
        if name == 'docs/release-preview-acceptance.json':
            continue
        path = ROOT / name
        digest.update(raw + b'\0')
        digest.update(hashlib.sha256(path.read_bytes()).digest())
    return digest.hexdigest()


def verify(data, package_version):
    if data.get('status') != 'accepted':
        raise ValueError('Preview status is not accepted.')
    if not re.fullmatch(r'\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?', package_version or '') or data.get('package_version') != package_version:
        raise ValueError('Preview record does not match the release package version.')
    if data.get('source_fingerprint') != source_fingerprint():
        raise ValueError('Preview record does not match this source tree.')
    sample = ROOT / 'e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx'
    if data.get('sample_sha256') != hashlib.sha256(sample.read_bytes()).hexdigest():
        raise ValueError('Preview record sample hash differs from the checked-in DOCX.')
    for key in ('native_ofd_sha256', 'preview_pdf_sha256'):
        if not re.fullmatch(r'[0-9a-f]{64}', data.get(key, '')):
            raise ValueError(f'{key} is missing.')
    if data.get('conversion_mode') != 'native' or data.get('view_chain') != 'DOCX -> native OFD -> PDF -> macOS Preview':
        raise ValueError('Record must describe the Native OFD Preview chain.')
    if not data.get('reviewer') or not data.get('reviewed_at'):
        raise ValueError('Human reviewer and timestamp are required.')
    if data.get('checked_pages') != [1, 2] or data.get('open_findings'):
        raise ValueError('Both deterministic DOCX pages must be checked with no open findings.')


if __name__ == '__main__':
    try:
        if '--print-fingerprint' in sys.argv:
            print(source_fingerprint())
            sys.exit(0)
        if len(sys.argv) != 2:
            raise ValueError('Pass the release package version.')
        verify(json.loads(RECORD.read_text()), sys.argv[1])
        print('Candidate-bound Native OFD macOS Preview record accepted.')
    except (OSError, ValueError, json.JSONDecodeError, subprocess.CalledProcessError) as error:
        print(f'Preview release gate failed: {error}', file=sys.stderr)
        sys.exit(1)
