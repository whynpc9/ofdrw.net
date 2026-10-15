#!/usr/bin/env python3
"""Reconstruct and verify the immutable R12 functional archive (requires zstd)."""
from pathlib import Path
import hashlib
import json
import shutil
import subprocess
import tarfile
import tempfile

root = Path(__file__).resolve().parent
manifest = json.loads((root / 'review-r12-functional-manifest.json').read_text())
with tempfile.TemporaryDirectory(prefix='ofd-r12-verify-') as tmp:
    combined = Path(tmp) / 'r12-functional.tar.zst'
    complete_hash = hashlib.sha256()
    complete_size = 0
    with combined.open('wb') as output:
        for index, row in enumerate(manifest['distribution']['parts'], 1):
            assert row['order'] == index
            source = root / row['file']
            digest = hashlib.sha256()
            size = 0
            with source.open('rb') as stream:
                while block := stream.read(1024 * 1024):
                    output.write(block)
                    digest.update(block)
                    complete_hash.update(block)
                    size += len(block)
            assert size == row['bytes'] and digest.hexdigest() == row['sha256'], source
            complete_size += size
    assert complete_size == manifest['archive_bytes']
    assert complete_hash.hexdigest() == manifest['archive_sha256']
    expected = {row['file']: row for row in manifest['files']}
    seen = set()
    process = subprocess.Popen(['zstd', '-q', '-dc', str(combined)], stdout=subprocess.PIPE)
    with tarfile.open(fileobj=process.stdout, mode='r|') as archive:
        for member in archive:
            assert member.isfile() and member.name in expected and member.name not in seen
            stream = archive.extractfile(member)
            digest = hashlib.sha256()
            size = 0
            while block := stream.read(1024 * 1024):
                digest.update(block)
                size += len(block)
            row = expected[member.name]
            assert size == row['bytes'] and digest.hexdigest() == row['sha256'], member.name
            seen.add(member.name)
    assert process.wait() == 0 and seen == set(expected)
print(f"Verified {len(seen)} files; original archive SHA256 {manifest['archive_sha256']}")
