#!/usr/bin/env python3
"""Maintainer-only, pinned four-RID build. Never called by consumer builds.

Requires macOS CLT/SDK matching PINNED_MAC, Docker, and the reviewed source SHA.
Builds and byte-compares two outputs for every RID before replacing any asset.
Linux containers have no network; the sole downloaded input is hash-checked.
"""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import plistlib
import re
import shutil
import struct
import subprocess
import tempfile
import urllib.request

ROOT = Path(__file__).resolve().parents[1]
NATIVE = ROOT / "native/Ofdrw.ProcessSupervisor"
MUSL_URL = "https://musl.libc.org/releases/musl-1.2.5.tar.gz"
MUSL_SHA = "a9a118bbe84d8764da0ea0d28b3ab3fae8477fc7e4085d90102b8596fc7c75e4"
IMAGES = {
    "linux-x64": ("linux/amd64", "gcc@sha256:82549aa8f90ada3236a8be70c74543132a76662ef33f0c3271ed802b81584a82"),
    "linux-arm64": ("linux/arm64", "gcc@sha256:ab78e7a512cead0e90364bb21f91c299a913e8b19fc481bf700d97218524f46d"),
}
PINNED_MAC = {
    "clt": "26.6.0.0.1781586589",
    "compiler": "Apple clang version 21.0.0 (clang-2100.1.1.101)",
    "sdkVersion": "26.5",
    "sdkBuild": "25F70",
    "sdkSettingsSha256": "f8d005f09381389167f9e0aeaa169bc9e7dff162ef22ca2fd8e98df7ff1acafe",
    "clangSha256": "f30550eab15fdf5ab8c0dc54c52679711241e5d4b636b027e18c09fef531775d",
    "ldSha256": "28d85b9af18c923db12e0b4ce70b80ee217f2b7cade0e872baa9e8ddc396c08d",
}
CFLAGS = ["-std=c11", "-Wall", "-Wextra", "-Werror", "-O2", "-fno-ident"]
LINUX_BUILD = r'''set -eu
export LC_ALL=C TZ=UTC SOURCE_DATE_EPOCH=1735689600 ZERO_AR_DATE=1
mkdir /work
cd /work
tar xzf /input/musl.tar.gz
cd musl-1.2.5
./configure --prefix=/opt/musl --disable-shared > /out/musl-configure.log
make -j8 > /out/musl-build.log 2>&1
make install > /out/musl-install.log 2>&1
gcc --version > /out/compiler.txt
ld --version > /out/linker.txt
for pass in first second; do
  mkdir /out/$pass
  /opt/musl/bin/musl-gcc -std=c11 -Wall -Wextra -Werror -O2 -fno-ident \
    -ffile-prefix-map=/input=/src -fno-pie -no-pie -nostdlib -static \
    -Wl,--build-id=none,-z,noexecstack,-Map=/out/$pass/link.map \
    /opt/musl/lib/crt1.o /opt/musl/lib/crti.o /input/supervisor.c \
    /opt/musl/lib/libc.a /opt/musl/lib/crtn.o -o /out/$pass/ofdrw-supervisor
done
readelf -h -l -d /out/first/ofdrw-supervisor > /out/readelf.txt
nm -u /out/first/ofdrw-supervisor > /out/undefined-symbols.txt
'''


def run(command, **kwargs):
    print("+", " ".join(str(c) for c in command), flush=True)
    return subprocess.check_output(command, text=True, **kwargs).strip()


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def save_json(path, value):
    path.write_text(json.dumps(value, indent=2) + "\n")


def evidence_text(text, build_root):
    text = text.replace(str(build_root), "<build>")
    text = re.sub(r"/tmp/cc[A-Za-z0-9]+\.o", "<supervisor.o>", text)
    text = re.sub(r"/var/folders/[^\s]+/supervisor-[A-Za-z0-9]+\.o", "<supervisor.o>", text)
    return "\n".join(line.rstrip() for line in text.splitlines()) + "\n"


def normalize_uuid(path, source_sha, rid):
    """Recent dyld requires LC_UUID. Set a source/RID-derived UUID before signing."""
    data = bytearray(path.read_bytes())
    count = struct.unpack_from("<I", data, 16)[0]
    offset = 32
    for _ in range(count):
        command, size = struct.unpack_from("<II", data, offset)
        if command == 0x1B:
            if size != 24:
                raise ValueError("invalid LC_UUID size")
            data[offset + 8:offset + 24] = hashlib.sha256((source_sha + ":" + rid).encode()).digest()[:16]
            path.write_bytes(data)
            return
        offset += size
    raise ValueError("linker omitted required LC_UUID")


def check_mac():
    sdk = Path(run(["xcrun", "--show-sdk-path"]))
    clt = run(["pkgutil", "--pkg-info-plist", "com.apple.pkg.CLTools_Executables"])
    actual = {
        "clt": plistlib.loads(clt.encode())["pkg-version"],
        "compiler": run(["xcrun", "clang", "--version"]).splitlines()[0],
        "sdkVersion": run(["xcrun", "--show-sdk-version"]),
        "sdkBuild": plistlib.loads((sdk / "System/Library/CoreServices/SystemVersion.plist").read_bytes())["ProductBuildVersion"],
    }
    actual["sdkSettingsSha256"] = sha(sdk / "SDKSettings.json")
    for name in ("clang", "ld"):
        actual[name + "Sha256"] = sha(Path(run(["xcrun", "--find", name])))
    if actual != PINNED_MAC:
        raise ValueError(f"macOS toolchain differs from pinned inputs: {actual}")
    actual["linker"] = run(["xcrun", "ld", "-v"], stderr=subprocess.STDOUT)
    return sdk, actual


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-sha256", required=True, help="SHA-256 explicitly frozen after native source review")
    parser.add_argument("--musl-source", type=Path, help="optional cached official tarball (still hash-checked)")
    parser.add_argument("--promote-execution-proof", type=Path,
                        help="promote an already complete verified candidate after all six ABI/libc execution records pass; does not rebuild")
    args = parser.parse_args()
    source = NATIVE / "supervisor.c"
    if sha(source) != args.source_sha256:
        raise ValueError("source does not match the supplied frozen SHA")
    spec = importlib.util.spec_from_file_location("verify_assets", ROOT / "scripts/verify-process-supervisor-assets.py")
    verifier = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(verifier)
    if args.promote_execution_proof:
        manifest = verifier.verify(require_publishable=False)
        evidence = json.loads(args.promote_execution_proof.read_text())
        verifier.verify_execution_evidence(manifest, evidence)
        manifest["executionEvidence"] = evidence
        manifest["publishable"] = True
        assets = NATIVE / "assets"
        save_json(assets / "manifest.json", manifest)
        (assets / "SupervisorAssetHashes.props").write_text(verifier.msbuild_props(manifest, assets))
        verifier.verify()
        return
    sdk, mac = check_mac()
    with tempfile.TemporaryDirectory(prefix="ofdrw-supervisor-build-") as temporary:
        stage = Path(temporary)
        inputs = stage / "input"
        inputs.mkdir()
        shutil.copyfile(source, inputs / "supervisor.c")
        tarball = inputs / "musl.tar.gz"
        if args.musl_source:
            shutil.copyfile(args.musl_source, tarball)
        else:
            with urllib.request.urlopen(MUSL_URL, timeout=120) as response:
                tarball.write_bytes(response.read())
        if sha(tarball) != MUSL_SHA:
            raise ValueError("musl source hash mismatch")
        entries = []
        for rid, (platform, image) in IMAGES.items():
            out = stage / rid
            out.mkdir()
            command = ["docker", "run", "--rm", "--network=none", "--platform", platform,
                       "-v", f"{inputs}:/input:ro", "-v", f"{out}:/out", image, "sh", "-c", LINUX_BUILD]
            run(command)
            linkmap = (out / "first/link.map").read_text()
            if "libgcc" in linkmap or "libclang_rt" in linkmap or (out / "undefined-symbols.txt").read_text().strip():
                raise ValueError(f"{rid}: compiler runtime or undefined symbols detected")
            toolchain = {"image": image, "platform": platform, "compiler": (out / "compiler.txt").read_text().splitlines()[0],
                         "linker": (out / "linker.txt").read_text().splitlines()[0],
                         "muslVersion": "1.2.5", "muslSource": MUSL_URL, "muslSourceSha256": MUSL_SHA,
                         "compilerRuntimeLinked": False}
            entries.append(record(verifier, rid, out, toolchain, ["sh", "-c", LINUX_BUILD], "Linux (minimum kernel not certified; see execution evidence)"))
        for rid, architecture, minimum in (("osx-x64", "x86_64", "10.15"), ("osx-arm64", "arm64", "11.0")):
            out = stage / rid
            out.mkdir()
            template = ["xcrun", "clang", *CFLAGS, "-arch", architecture, "-isysroot", str(sdk),
                        "-mmacosx-version-min=" + minimum, "-ffile-prefix-map=" + str(inputs) + "=/src",
                        "-Wl,-no_adhoc_codesign", str(inputs / "supervisor.c"), "-lproc"]
            for name in ("first", "second"):
                directory = out / name
                directory.mkdir()
                binary = directory / "ofdrw-supervisor"
                run([*template, "-Wl,-map," + str(directory / "link.map"), "-o", str(binary)])
                normalize_uuid(binary, args.source_sha256, rid)
                run(["codesign", "--force", "--sign", "-", "--identifier", "org.ofdrw.process-supervisor", "--timestamp=none", str(binary)], stderr=subprocess.STDOUT)
                run(["codesign", "--verify", "--strict", str(binary)], stderr=subprocess.STDOUT)
            toolchain = dict(mac, uuidRule="SHA256(sourceSha256 + ':' + rid) first 16 bytes, before signing",
                             codesignCommand=["codesign", "--force", "--sign", "-", "--identifier", "org.ofdrw.process-supervisor", "--timestamp=none", "<output>/ofdrw-supervisor"],
                             signatureVerificationCommand=["codesign", "--verify", "--strict", "<output>/ofdrw-supervisor"])
            entries.append(record(verifier, rid, out, toolchain, template + ["-Wl,-map,<output>/link.map", "-o", "<output>/ofdrw-supervisor"], "macOS " + minimum))
        if sha(source) != args.source_sha256:
            raise ValueError("source changed during build; discard outputs and freeze again")
        licenses = [{"path": "docs/licenses/MUSL-COPYRIGHT.txt", "sha256": sha(ROOT / "docs/licenses/MUSL-COPYRIGHT.txt"),
                     "source": MUSL_URL, "sourceSha256": MUSL_SHA}]
        manifest = {"schemaVersion": 1, "protocolVersion": 1, "sourcePath": "native/Ofdrw.ProcessSupervisor/supervisor.c",
                    "sourceSha256": args.source_sha256, "publishable": False, "files": entries, "licenses": licenses}
        complete = stage / "complete"
        complete.mkdir()
        for entry in entries:
            target = complete / entry["path"]
            target.parent.mkdir()
            shutil.copyfile(stage / entry["rid"] / "first/ofdrw-supervisor", target)
            target.chmod(0o755)
            # Link maps and native tool output remain reviewable beside the binary.
            for name in ("link.map",):
                (target.parent / name).write_text(evidence_text((stage / entry["rid"] / "first" / name).read_text(), stage))
            if entry["rid"].startswith("linux-"):
                (target.parent / "readelf.txt").write_text(evidence_text((stage / entry["rid"] / "readelf.txt").read_text(), stage))
        save_json(complete / "manifest.json", manifest)
        (complete / "NativeSupervisorAssetManifest.g.cs").write_text(verifier.csharp(manifest))
        (complete / "SupervisorAssetHashes.props").write_text(verifier.msbuild_props(manifest, complete))
        verifier.verify(complete, require_publishable=False)
        assets = NATIVE / "assets"
        assets.mkdir(exist_ok=True)
        for entry in entries:
            target = assets / entry["rid"]
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(complete / entry["rid"], target)
        shutil.copyfile(complete / "manifest.json", assets / "manifest.json")
        for name in ("NativeSupervisorAssetManifest.g.cs", "SupervisorAssetHashes.props"):
            shutil.copyfile(complete / name, assets / name)
    verifier.verify(require_publishable=False)


def record(verifier, rid, directory, toolchain, command, minimum):
    first = directory / "first/ofdrw-supervisor"
    second = directory / "second/ofdrw-supervisor"
    digest = sha(first)
    if digest != sha(second):
        raise ValueError(f"{rid}: independent compilation outputs differ")
    header = verifier.inspect_header(first.read_bytes(), rid)
    # Absolute temporary roots are normalized in documentary commands/maps.
    command = [arg.replace(str(directory.parent), "<build>") for arg in command]
    return {"rid": rid, "path": f"{rid}/ofdrw-supervisor", "sha256": digest, "size": first.stat().st_size,
            "format": header["format"], "architecture": rid.split("-", 1)[1], "minimumOS": minimum,
            "toolchain": toolchain, "buildCommand": command, "header": header,
            "reproducible": {"firstSha256": digest, "secondSha256": sha(second)}}


if __name__ == "__main__":
    main()
