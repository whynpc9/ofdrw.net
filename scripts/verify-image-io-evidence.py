#!/usr/bin/env python3
"""Verify the committed Issue 02 archive and every exact generated payload (stdlib only)."""
import argparse
import hashlib
import json
from pathlib import Path
import tarfile


def verify(directory):
    manifest = json.loads((directory / "manifest.json").read_text())
    archive = directory / manifest["archive"]["path"]
    if archive.stat().st_size != manifest["archive"]["bytes"]:
        raise ValueError("Evidence archive byte count mismatch")
    if hashlib.sha256(archive.read_bytes()).hexdigest() != manifest["archive"]["sha256"]:
        raise ValueError("Evidence archive hash mismatch")
    expected = {entry["path"]: entry for entry in manifest["files"]}
    if len(expected) != len(manifest["files"]):
        raise ValueError("Duplicate manifest paths")
    with tarfile.open(archive, "r:xz") as bundle:
        seen = set()
        for member in bundle:
            if not member.isfile() or member.name not in expected or member.name in seen:
                raise ValueError(f"Unexpected archive member: {member.name}")
            seen.add(member.name)
            entry = expected[member.name]
            if member.size != entry["bytes"]:
                raise ValueError(f"Payload byte count mismatch: {member.name}")
            digest = hashlib.sha256()
            stream = bundle.extractfile(member)
            while block := stream.read(81920):
                digest.update(block)
            if digest.hexdigest() != entry["sha256"]:
                raise ValueError(f"Payload hash mismatch: {member.name}")
        if seen != set(expected):
            raise ValueError("Missing evidence payloads")
    print(f"Verified {len(expected)} exact payloads; source {manifest['source_commit']}; Preview {manifest['preview']['status']}.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", nargs="?", type=Path,
                        default=Path(__file__).resolve().parents[1] / "docs/validation/issue02")
    verify(parser.parse_args().directory)
