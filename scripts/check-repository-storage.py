#!/usr/bin/env python3
"""Check staged Git objects, even when optional LFS evidence is not downloaded."""

import argparse
import hashlib
from pathlib import Path
import re
import subprocess
import sys


POINTER = re.compile(rb"version https://git-lfs.github.com/spec/v1\noid sha256:([0-9a-f]{64})\nsize ([0-9]+)\n")
MAX_GIT_BYTES = 5 * 1024 * 1024


def git(*args, data=None):
    return subprocess.check_output(["git", *args], input=data)


def check(materialized=()):
    entries = []
    errors = []
    for row in git("ls-files", "--stage", "-z").split(b"\0"):
        if not row:
            continue
        meta, path = row.split(b"\t", 1)
        mode, oid, stage = meta.split()
        if stage != b"0":
            errors.append(f"Unmerged path: {path.decode()}")
        elif mode != b"160000":
            entries.append((oid, path))
    attrs = git("check-attr", "--cached", "-z", "--stdin", "filter",
                data=b"".join(p + b"\0" for _, p in entries)).split(b"\0")
    filters = {attrs[i]: attrs[i + 2] for i in range(0, len(attrs) - 1, 3)}
    sizes = git("cat-file", "--batch-check=%(objectsize)",
                data=b"".join(oid + b"\n" for oid, _ in entries)).splitlines()
    regular_bytes = lfs_bytes = lfs_count = 0
    for (oid, path), raw_size in zip(entries, sizes):
        name = path.decode()
        size = int(raw_size)
        if name.startswith("artifacts/"):
            errors.append(f"Generated artifacts must stay untracked: {name}")
        if filters.get(path) != b"lfs":
            regular_bytes += size
            if size > MAX_GIT_BYTES:
                errors.append(f"Git blob exceeds 5 MiB; use LFS or remove it: {name}")
            continue
        pointer = POINTER.fullmatch(git("cat-file", "blob", oid.decode())) if size < 1024 else None
        if pointer is None:
            errors.append(f"LFS path contains a raw blob or invalid pointer: {name}")
            continue
        lfs_count += 1
        lfs_bytes += int(pointer[2])
        if any(name == prefix.rstrip("/") or name.startswith(prefix.rstrip("/") + "/")
               for prefix in materialized):
            file = Path(name)
            if not file.is_file() or file.stat().st_size != int(pointer[2]):
                errors.append(f"Missing LFS payload; run git lfs pull for: {name}")
            else:
                with file.open("rb") as source:
                    digest = hashlib.file_digest(source, "sha256").hexdigest().encode()
                if digest != pointer[1]:
                    errors.append(f"LFS payload hash mismatch: {name}")
    for error in errors:
        print(error, file=sys.stderr)
    print(f"Git blobs: {regular_bytes:,} bytes; LFS: {lfs_count} files / {lfs_bytes:,} payload bytes")
    return 1 if errors else 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-materialized", action="append", default=[], metavar="DIRECTORY")
    args = parser.parse_args()
    sys.exit(check(args.require_materialized))
