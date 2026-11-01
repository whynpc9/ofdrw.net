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
LFS_POINTER = re.compile(rb'version https://git-lfs.github.com/spec/v1\noid sha256:([0-9a-f]{64})\nsize ([0-9]+)\n?')


def git(*args, data=None):
    return subprocess.check_output(['git', *args], cwd=ROOT, input=data)


def indexed_lfs_pointers(entries):
    """Read only small, valid repository-standard pointers from the current index."""
    attributes = git('check-attr', '--cached', '-z', '--stdin', 'filter',
                     data=b''.join(path + b'\0' for path in entries)).split(b'\0')
    paths = [attributes[i] for i in range(0, len(attributes) - 1, 3)
             if attributes[i + 2] == b'lfs']
    if not paths:
        return {}
    object_ids = b''.join(entries[path] + b'\n' for path in paths)
    sizes = git('cat-file', '--batch-check=%(objectsize)', data=object_ids).splitlines()
    if len(sizes) != len(paths) or any(not size.isdigit() or int(size) >= 1024 for size in sizes):
        raise ValueError('Stage valid Git LFS pointers before computing the source fingerprint.')
    blobs = git('cat-file', '--batch', data=object_ids)
    offset, pointers = 0, {}
    for path, size in zip(paths, sizes):
        end = blobs.index(b'\n', offset)
        header = blobs[offset:end].split()
        if header != [entries[path], b'blob', size]:
            raise ValueError('Unexpected Git LFS index object response.')
        offset = end + 1
        pointer = LFS_POINTER.fullmatch(blobs[offset:offset + int(size)])
        if pointer is None:
            raise ValueError(f'Stage a valid Git LFS pointer before fingerprinting: {path.decode()}')
        pointers[path] = (pointer[1], int(pointer[2]))
        offset += int(size) + 1
    return pointers


def lfs_content_hash(path, identity):
    """Bind both pointer and hydrated forms to the verified oid/size identity."""
    oid, size = identity
    if path.is_symlink() or not path.is_file():
        raise ValueError(f'Git LFS working path must be a regular file: {path}')
    length = path.stat().st_size
    pointer = LFS_POINTER.fullmatch(path.read_bytes()) if length < 1024 else None
    if pointer is not None:
        if (pointer[1], int(pointer[2])) != identity:
            raise ValueError(f'Git LFS working pointer differs from the index: {path}')
    else:
        if length != size:
            raise ValueError(f'Git LFS payload size differs from the index: {path}')
        digest = hashlib.sha256()
        with path.open('rb') as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b''):
                digest.update(chunk)
        if digest.hexdigest().encode() != oid:
            raise ValueError(f'Git LFS payload hash differs from the index: {path}')
    canonical = b'version https://git-lfs.github.com/spec/v1\noid sha256:' + oid + b'\nsize ' + str(size).encode() + b'\n'
    return hashlib.sha256(canonical).digest()


def source_fingerprint():
    entries = {}
    for row in git('ls-files', '--stage', '-z').split(b'\0'):
        if not row:
            continue
        metadata, path = row.split(b'\t', 1)
        _, oid, stage = metadata.split()
        if stage != b'0':
            raise ValueError(f'Resolve the unmerged path before fingerprinting: {path.decode()}')
        entries[path] = oid
    pointers = indexed_lfs_pointers(entries)
    digest = hashlib.sha256()
    for raw in sorted(entries):
        name = raw.decode()
        if name == 'docs/release-preview-acceptance.json':
            continue
        path = ROOT / name
        digest.update(raw + b'\0')
        digest.update(lfs_content_hash(path, pointers[raw]) if raw in pointers
                      else hashlib.sha256(path.read_bytes()).digest())
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
    for key in ('native_ofd_sha256', 'preview_pdf_sha256',
                'default_ofd_sha256', 'default_preview_pdf_sha256'):
        if not re.fullmatch(r'[0-9a-f]{64}', data.get(key, '')):
            raise ValueError(f'{key} is missing.')
    if data.get('conversion_mode') != 'native' or data.get('view_chain') != 'DOCX -> native OFD -> PDF -> macOS Preview':
        raise ValueError('Record must describe the Native OFD Preview chain.')
    if data.get('default_conversion_mode') != 'default' or data.get('default_view_chain') != 'DOCX -> default OFD -> PDF -> macOS Preview':
        raise ValueError('Record must describe the default OFD Preview chain.')
    if not data.get('reviewer') or not data.get('reviewed_at'):
        raise ValueError('Human reviewer and timestamp are required.')
    if data.get('checked_pages') != [1, 2] or data.get('default_checked_pages') != [1, 2] or data.get('open_findings'):
        raise ValueError('Both pages in Native and default modes must be checked with no open findings.')


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
