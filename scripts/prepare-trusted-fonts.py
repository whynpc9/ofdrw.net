#!/usr/bin/env python3
"""Prepare the pinned repository-local OFL fixture ZIP; Python standard library only."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
import struct
import tempfile
import zipfile

SPEC_SHA256 = 'a5215982dd59177a0846b183f6cc89e683c281be281993f97eda1181367e5264'
ALLOWED = frozenset(('fonts/NotoSans-Regular.ttf', 'fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf',
                     'licenses/NotoSans-LICENSE', 'licenses/Ofdrw-CI-Noto-OFL.txt', 'manifest.json', 'NOTICE.md'))
MAX_ARCHIVE = 24 * 1024 * 1024
MAX_MEMBER = 24 * 1024 * 1024
MAX_TOTAL = 32 * 1024 * 1024
MAX_METADATA = 64 * 1024
MAX_RATIO = 64
CHUNK = 81920


def digest_stream(stream):
    h = hashlib.sha256()
    stream.seek(0)
    while True:
        chunk = stream.read(CHUNK)
        if not chunk:
            return h.hexdigest()
        h.update(chunk)


def read_spec(path):
    with path.open('rb') as stream:
        raw = stream.read(MAX_METADATA + 1)
    if len(raw) > MAX_METADATA or hashlib.sha256(raw).hexdigest() != SPEC_SHA256:
        raise ValueError('Unrecognized pinned bundle manifest')
    spec = json.loads(raw)
    members = spec['members']
    if len(members) != 6 or {m['path'] for m in members} != ALLOWED:
        raise ValueError('Manifest must declare exactly four payloads and two metadata files')
    if spec['archive']['file'] != 'sha256-' + spec['archive']['sha256'] + '.zip':
        raise ValueError('Bundle filename must be content addressed')
    return spec


def preflight_directory(stream, size):
    """Bound/count central-directory records before ZipFile creates its member objects."""
    stream.seek(max(0, size - 65557))
    tail = stream.read(65557)
    at = tail.rfind(b'PK\x05\x06')
    if at < 0 or at + 22 > len(tail):
        raise ValueError('Missing ordinary ZIP end record')
    disk, cd_disk, count_disk, count, cd_size, cd_offset, comment = struct.unpack_from('<4H2IH', tail, at + 4)
    eocd = size - len(tail) + at
    if disk or cd_disk or count_disk != count or count != 6 or comment or at + 22 != len(tail):
        raise ValueError('Only a single-disk six-member ZIP without comments is admitted')
    if cd_size > MAX_METADATA or cd_offset + cd_size != eocd:
        raise ValueError('Invalid or oversized ZIP directory')
    stream.seek(cd_offset)
    cd = stream.read(cd_size)
    offset = records = 0
    while offset < len(cd):
        records += 1
        if records > 6 or offset + 46 > len(cd) or cd[offset:offset + 4] != b'PK\x01\x02':
            raise ValueError('Invalid ZIP directory record count')
        name, extra, note = struct.unpack_from('<3H', cd, offset + 28)
        offset += 46 + name + extra + note
        if offset > len(cd):
            raise ValueError('Truncated ZIP directory')
    if records != 6:
        raise ValueError('ZIP count does not match actual directory')
    stream.seek(0)


def validate_members(archive, spec):
    expected = {m['path']: m for m in spec['members']}
    infos = archive.infolist()
    if len(infos) != 6:
        raise ValueError('Unexpected member count')
    seen = set()
    total = 0
    for info in infos:
        name = info.filename
        mode = info.external_attr >> 16
        if name != info.orig_filename or name not in ALLOWED or name in seen or '\\' in name:
            raise ValueError('Unknown, duplicate or unsafe member path')
        if info.is_dir() or stat.S_IFMT(mode) not in (0, stat.S_IFREG) or info.flag_bits & 1:
            raise ValueError('Only unencrypted regular files are allowed')
        if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
            raise ValueError('Unsupported compression')
        if info.file_size != expected[name]['bytes'] or info.file_size > MAX_MEMBER or info.file_size < 0:
            raise ValueError('Member size differs from manifest or limit')
        if info.compress_size <= 0 or info.file_size > info.compress_size * MAX_RATIO:
            raise ValueError('Compression ratio exceeds limit')
        if name in ('manifest.json', 'NOTICE.md') and info.file_size > MAX_METADATA:
            raise ValueError('Oversized metadata')
        total += info.file_size
        if total > MAX_TOTAL:
            raise ValueError('Total expanded size exceeds limit')
        seen.add(name)
    if seen != ALLOWED:
        raise ValueError('Missing member')
    return infos


def copy_member(archive, info, destination, expected):
    h = hashlib.sha256()
    count = 0
    destination.parent.mkdir(parents=True, exist_ok=True)
    with archive.open(info) as source, destination.open('xb') as target:
        while True:
            chunk = source.read(min(CHUNK, expected['bytes'] - count + 1))
            if not chunk:
                break
            count += len(chunk)
            if count > expected['bytes'] or count > MAX_MEMBER:
                raise ValueError('Expanded member exceeded reservation')
            h.update(chunk)
            target.write(chunk)
    if count != expected['bytes'] or h.hexdigest() != expected['sha256']:
        raise ValueError('Expanded member size or SHA mismatch')


def verify_prepared(directory, spec):
    if directory.is_symlink() or not directory.is_dir():
        raise ValueError('Prepared target must be an owned regular directory')
    files = set()
    expected_top = {'fonts', 'licenses', 'manifest.json', 'NOTICE.md'}
    seen_top = set()
    with os.scandir(directory) as entries:
        for entry in entries:
            if entry.name not in expected_top or entry.name in seen_top or entry.is_symlink():
                raise ValueError('Prepared target has extra, duplicate or linked items')
            seen_top.add(entry.name)
            if entry.name in ('fonts', 'licenses'):
                if not entry.is_dir(follow_symlinks=False):
                    raise ValueError('Expected regular payload directory')
                seen_children = set()
                with os.scandir(entry.path) as children:
                    for child in children:
                        name = entry.name + '/' + child.name
                        if name not in ALLOWED or name in seen_children or child.is_symlink() or not child.is_file(follow_symlinks=False):
                            raise ValueError('Prepared payload path is not an allowed regular file')
                        seen_children.add(name)
                        files.add(name)
            elif not entry.is_file(follow_symlinks=False):
                raise ValueError('Expected regular metadata file')
            else:
                files.add(entry.name)
    if seen_top != expected_top or files != ALLOWED:
        raise ValueError('Prepared target has missing or extra files')
    for member in spec['members']:
        path = directory / member['path']
        if path.stat().st_size != member['bytes']:
            raise ValueError('Prepared member size mismatch')
        with path.open('rb') as stream:
            if digest_stream(stream) != member['sha256']:
                raise ValueError('Prepared member SHA mismatch')


def prepare(manifest, bundle, output):
    spec = read_spec(manifest)
    output = output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    digest = spec['archive']['sha256']
    destination = output / ('sha256-' + digest)
    lock = output / ('.' + digest + '.lock')
    stage = None
    # Exclusive cooperating-preparer lock; an occupied lock fails without overwriting it.
    fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    try:
        os.close(fd)
        with bundle.open('rb') as stream:
            size = os.fstat(stream.fileno()).st_size
            if size != spec['archive']['bytes'] or size > MAX_ARCHIVE or digest_stream(stream) != digest:
                raise ValueError('Bundle size or SHA mismatch')
            preflight_directory(stream, size)
            with zipfile.ZipFile(stream) as archive:
                infos = validate_members(archive, spec)
                if destination.exists() or destination.is_symlink():
                    verify_prepared(destination, spec)
                    if digest_stream(stream) != digest:
                        raise ValueError("Bundle changed during cache verification")
                else:
                    stage = Path(tempfile.mkdtemp(prefix='.trusted-fonts-', dir=output))
                    expected = {m['path']: m for m in spec['members']}
                    for info in infos:
                        copy_member(archive, info, stage / info.filename, expected[info.filename])
                    verify_prepared(stage, spec)
                    if digest_stream(stream) != digest:
                        raise ValueError('Bundle changed during preparation')
                    os.rename(stage, destination)
                    stage = None
        return {'bundleSha256': digest, 'directory': str(destination), 'environment': {
            'OFDRW_REQUIRE_TRUSTED_FONTS': '1',
            'OFDRW_TEST_TRUSTED_LATIN_FONT': str(destination / 'fonts/NotoSans-Regular.ttf'),
            'OFDRW_TEST_TRUSTED_CJK_FONT': str(destination / 'fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf')}}
    finally:
        if stage is not None:
            shutil.rmtree(stage)
        lock.unlink()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    default = Path(__file__).resolve().parent.parent / 'tests/fixtures/trusted-fonts/bundle.json'
    parser.add_argument('--manifest', type=Path, default=default)
    parser.add_argument('--bundle', type=Path)
    parser.add_argument('--output-dir', type=Path, required=True)
    args = parser.parse_args()
    spec = read_spec(args.manifest)
    bundle = args.bundle or args.manifest.parent / spec['archive']['file']
    print(json.dumps(prepare(args.manifest, bundle, args.output_dir), indent=2))


if __name__ == '__main__':
    main()
