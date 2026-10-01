#!/usr/bin/env python3
"""Verify the exact retained table acceptance products; optionally extract them."""

import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import subprocess
import sys
import tarfile


def digest(path):
    value = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def safe_name(name):
    path = PurePosixPath(name)
    if path.is_absolute() or ".." in path.parts or "\\" in name:
        raise ValueError(f"Unsafe evidence path: {name}")
    return path


def verify(directory, destination=None):
    manifest = json.loads((directory / "bundle-manifest.json").read_text())
    archive_name = safe_name(manifest["archive"])
    if len(archive_name.parts) != 1:
        raise ValueError("Evidence archive must be in the manifest directory.")
    archive = directory / str(archive_name)
    if archive.stat().st_size != manifest["archiveBytes"] or digest(archive) != manifest["archiveSha256"]:
        raise ValueError("Evidence archive size or SHA-256 mismatch.")
    expected = manifest["files"]
    for name in expected:
        safe_name(name)
    seen = set()
    process = subprocess.Popen(["zstd", "-d", "-c", "--long=27", str(archive)],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    try:
        with tarfile.open(fileobj=process.stdout, mode="r|") as products:
            for member in products:
                relative = safe_name(member.name)
                if member.isdir():
                    continue
                if not member.isfile() or member.name not in expected or member.name in seen:
                    raise ValueError(f"Unexpected or duplicate evidence entry: {member.name}")
                record = expected[member.name]
                if member.size != record["bytes"]:
                    raise ValueError(f"Evidence size mismatch: {member.name}")
                value = hashlib.sha256()
                target = None
                if destination is not None:
                    output = destination.joinpath(*relative.parts)
                    output.resolve().relative_to(destination.resolve())
                    output.parent.mkdir(parents=True, exist_ok=True)
                    if output.exists() or output.is_symlink():
                        raise ValueError(f"Refusing to overwrite evidence: {output}")
                    target = output.open("wb")
                try:
                    with products.extractfile(member) as source:
                        for chunk in iter(lambda: source.read(1024 * 1024), b""):
                            value.update(chunk)
                            if target is not None:
                                target.write(chunk)
                finally:
                    if target is not None:
                        target.close()
                if value.hexdigest() != record["sha256"]:
                    raise ValueError(f"Evidence SHA-256 mismatch: {member.name}")
                seen.add(member.name)
        # Drain archive padding before waiting for the decoder.
        for _ in iter(lambda: process.stdout.read(1024 * 1024), b""):
            pass
        if process.wait() != 0:
            raise ValueError(process.stderr.read().decode(errors="replace"))
    finally:
        if process.poll() is None:
            process.terminate()
            process.wait()
        process.stdout.close()
        process.stderr.close()
    if seen != set(expected):
        raise ValueError(f"Missing evidence entries: {sorted(set(expected) - seen)}")
    prefix = "public-tables/review1/current/"
    direct_names = {name[len(prefix):] for name in expected if name.startswith(prefix) and
                    (name.endswith((".pdf", ".png")) or
                     PurePosixPath(name).name in {"acceptance.json", "artifact-manifest.tsv"})}
    preview = directory / "preview"
    actual_names = {path.relative_to(preview).as_posix() for path in preview.rglob("*") if path.is_file()}
    if actual_names != direct_names:
        raise ValueError("Missing or unexpected direct preview copies.")
    for name in direct_names:
        path = preview / name
        record = expected[prefix + name]
        if path.stat().st_size != record["bytes"] or digest(path) != record["sha256"]:
            raise ValueError(f"Direct preview copy mismatch: {path}")
    direct = len(direct_names)
    print(f"Verified {len(seen)} original evidence files and {direct} direct preview copies.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--evidence-dir", type=Path, default=Path(__file__).resolve().parents[1] /
                        "docs/validation/public-tables-evidence")
    parser.add_argument("--extract", type=Path, help="Extract into a new directory; existing files are not overwritten.")
    args = parser.parse_args()
    try:
        verify(args.evidence_dir, args.extract)
    except (OSError, ValueError, KeyError, tarfile.TarError) as error:
        print(f"Evidence verification failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
