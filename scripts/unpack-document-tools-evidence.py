#!/usr/bin/env python3
"""Extract the review bundle safely and verify every file against its manifest."""
import argparse, hashlib, json, pathlib, shutil, subprocess, tarfile
parser = argparse.ArgumentParser()
parser.add_argument('destination', type=pathlib.Path)
parser.add_argument('--bundle', type=pathlib.Path, default=pathlib.Path(__file__).resolve().parents[1]/'docs/evidence/document-tools/evidence.tar.zst')
args=parser.parse_args()
root=args.destination.resolve(); root.mkdir(parents=True,exist_ok=True)
with subprocess.Popen(['zstd','-d','-c',str(args.bundle)],stdout=subprocess.PIPE) as process:
    with tarfile.open(fileobj=process.stdout,mode='r|') as archive:
        for member in archive:
            path=pathlib.PurePosixPath(member.name)
            if path.is_absolute() or '..' in path.parts or not (member.isdir() or member.isfile()):
                raise ValueError(f'Unsafe bundle member: {member.name}')
            target=root.joinpath(*path.parts)
            if not target.resolve().is_relative_to(root): raise ValueError('Destination symlink escapes output')
            if member.isdir(): target.mkdir(parents=True,exist_ok=True)
            else:
                target.parent.mkdir(parents=True,exist_ok=True)
                with archive.extractfile(member) as source, target.open('wb') as output: shutil.copyfileobj(source,output)
    if process.wait()!=0: raise RuntimeError('zstd extraction failed')
manifest=json.loads((root/'manifest.json').read_text())
for path,expected in manifest['files'].items():
    file=root/path
    if file.stat().st_size!=expected['bytes'] or hashlib.sha256(file.read_bytes()).hexdigest()!=expected['sha256']:
        raise ValueError(f'Integrity mismatch: {path}')
print(f'Extracted and verified {len(manifest["files"])} files to {root}')
