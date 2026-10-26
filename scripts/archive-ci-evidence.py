#!/usr/bin/env python3
"""Losslessly archive all evidence files with a per-file SHA-256 inventory."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tarfile


def sha256(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def archive(source, destination):
    source, destination = source.resolve(), destination.resolve()
    if not source.is_dir() or destination.is_relative_to(source):
        raise ValueError('Evidence source must exist and destination must be outside it')
    files = sorted(p for p in source.rglob('*') if p.is_file())
    if not files or any(p.is_symlink() for p in source.rglob('*')):
        raise ValueError('Evidence must be nonempty and contain no symlinks')
    records = [{'path': p.relative_to(source).as_posix(), 'bytes': p.stat().st_size,
                'sha256': sha256(p)} for p in files]
    destination.parent.mkdir(parents=True, exist_ok=True)
    # A long window compresses repeated embedded fonts across files without
    # deleting or rewriting any OFD/PDF/SVG/page image in the evidence set.
    with subprocess.Popen(['zstd', '-q', '-f', '-T2', '-6', '--long=27', '-o', str(destination)],
                          stdin=subprocess.PIPE) as process:
        try:
            with tarfile.open(fileobj=process.stdin, mode='w|') as tar:
                tar.add(source, arcname=source.name)
        finally:
            process.stdin.close()
        if process.wait():
            raise ValueError('Evidence compression failed')
    subprocess.run(['zstd', '-q', '-t', str(destination)], check=True)
    manifest = {'sourceSha': os.environ.get('GITHUB_SHA'), 'workflowRun': os.environ.get('GITHUB_RUN_ID'),
                'root': source.name, 'files': records, 'inputBytes': sum(r['bytes'] for r in records),
                'archiveBytes': destination.stat().st_size, 'archiveSha256': sha256(destination)}
    destination.with_suffix(destination.suffix + '.json').write_text(json.dumps(manifest, indent=2) + '\n')
    print(f"Evidence: {len(files)} files, {manifest['inputBytes']} bytes -> {manifest['archiveBytes']} bytes")
    if os.environ.get('GITHUB_STEP_SUMMARY'):
        with open(os.environ['GITHUB_STEP_SUMMARY'], 'a', encoding='utf-8') as summary:
            summary.write(f"\nEvidence `{source.name}`: {len(files)} files, {manifest['inputBytes']:,} bytes -> {manifest['archiveBytes']:,} bytes (lossless).\n")
    return manifest


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('destination', type=Path)
    args = parser.parse_args()
    archive(args.source, args.destination)
