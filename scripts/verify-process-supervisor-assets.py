#!/usr/bin/env python3
"""Verify the complete committed native delivery; uses only Python's stdlib.

Consumers validate these same bytes in managed MSBuild/runtime code. This script
is a maintainer/CI check and never invokes a compiler or executes an asset.
"""
import argparse
import hashlib
import json
from pathlib import Path
import struct
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "native/Ofdrw.ProcessSupervisor/assets"
RIDS = ("linux-x64", "linux-arm64", "osx-x64", "osx-arm64")
MUSL_SOURCE_SHA256 = "a9a118bbe84d8764da0ea0d28b3ab3fae8477fc7e4085d90102b8596fc7c75e4"
MUSL_NOTICE_SHA256 = "41c681ab36d2e05dc2fbfc875f6a68186ec8e425de61354adaee0845caee272e"
LINUX_IMAGES = {
    "linux-x64": "gcc@sha256:82549aa8f90ada3236a8be70c74543132a76662ef33f0c3271ed802b81584a82",
    "linux-arm64": "gcc@sha256:ab78e7a512cead0e90364bb21f91c299a913e8b19fc481bf700d97218524f46d",
}


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def require(condition, message):
    if not condition:
        raise ValueError(message)


def version(value):
    return f"{value >> 16}.{(value >> 8) & 255}.{value & 255}"


def inspect_header(data, rid):
    """Read bounded native headers without depending on platform binutils."""
    if rid.startswith("linux-"):
        require(len(data) >= 64 and data[:7] == b"\x7fELF\x02\x01\x01", f"{rid}: expected ELF64 little-endian")
        kind, machine = struct.unpack_from("<HH", data, 16)
        require(kind == 2, f"{rid}: expected ET_EXEC")
        require(machine == (62 if rid.endswith("x64") else 183), f"{rid}: incorrect ELF machine")
        phoff = struct.unpack_from("<Q", data, 32)[0]
        phsize, phnum = struct.unpack_from("<HH", data, 54)
        require(phsize == 56 and phoff + phsize * phnum <= len(data), f"{rid}: invalid program header table")
        types = [struct.unpack_from("<I", data, phoff + i * phsize)[0] for i in range(phnum)]
        require(1 in types and 2 not in types and 3 not in types, f"{rid}: dynamic table/interpreter or no load segments")
        return {"format": "ELF64", "machine": machine, "type": "ET_EXEC", "interpreter": None, "dynamic": False}
    require(len(data) >= 32 and data[:4] == b"\xcf\xfa\xed\xfe", f"{rid}: expected thin Mach-O 64 little-endian")
    cpu, _, kind, count, size = struct.unpack_from("<IIIII", data, 4)
    require(cpu == (0x1000007 if rid.endswith("x64") else 0x100000C) and kind == 2, f"{rid}: incorrect Mach-O architecture/type")
    require(32 + size <= len(data) and count <= size // 8, f"{rid}: invalid load commands")
    offset, minimum, sdk, dylibs, signed, uuid = 32, None, None, [], False, None
    for _ in range(count):
        require(offset + 8 <= 32 + size, f"{rid}: truncated load command")
        command, length = struct.unpack_from("<II", data, offset)
        require(length >= 8 and offset + length <= 32 + size, f"{rid}: invalid load command")
        if command == 0x1B:
            require(length == 24 and uuid is None, f"{rid}: malformed/repeated LC_UUID")
            uuid = data[offset + 8:offset + 24].hex()
        elif command == 0x32:
            require(length >= 24, f"{rid}: short LC_BUILD_VERSION")
            platform, min_value, sdk_value = struct.unpack_from("<III", data, offset + 8)
            require(platform == 1, f"{rid}: expected macOS platform")
            minimum, sdk = version(min_value), version(sdk_value)
        elif command == 0x24:
            require(length >= 16, f"{rid}: short LC_VERSION_MIN_MACOSX")
            min_value, sdk_value = struct.unpack_from("<II", data, offset + 8)
            minimum, sdk = version(min_value), version(sdk_value)
        elif command in (0xC, 0x80000018, 0x8000001F, 0x80000023):
            require(length >= 24, f"{rid}: short dylib command")
            start = struct.unpack_from("<I", data, offset + 8)[0]
            require(24 <= start < length, f"{rid}: invalid dylib name")
            name = data[offset + start:offset + length].split(b"\0", 1)[0].decode("ascii")
            require(name in ("/usr/lib/libSystem.B.dylib", "/usr/lib/libproc.dylib"), f"{rid}: unexpected dependency {name}")
            dylibs.append(name)
        elif command == 0x1D:
            signed = True
        offset += length
    require(offset == 32 + size, f"{rid}: load command size mismatch")
    require(minimum == ("10.15.0" if rid.endswith("x64") else "11.0.0"), f"{rid}: incorrect minimum macOS")
    require(sdk == "26.5.0" and dylibs and signed and uuid, f"{rid}: wrong SDK, missing system library/UUID, or unsigned binary")
    return {"format": "Mach-O64", "cpuType": cpu, "type": "MH_EXECUTE", "minimumOS": minimum, "sdk": sdk, "dylibs": dylibs, "uuid": uuid, "adHocSigned": True}


def csharp(manifest):
    lines = ["// Generated only after complete native asset verification. Do not edit.",
             "namespace Ofdrw.Net.Core.Processes", "{", "    internal static partial class NativeSupervisorAssets", "    {",
             "        private static Asset FindAsset(string rid)", "        {", "            switch (rid)", "            {"]
    for entry in manifest["files"]:
        lines.append(f'                case "{entry["rid"]}": return new Asset("{entry["rid"]}", {entry["size"]}L, "{entry["sha256"]}");')
    lines.extend(['                default: throw new System.PlatformNotSupportedException("Unsupported native supervisor RID: " + rid);',
                  "            }", "        }", "    }", "}", ""])
    return "\n".join(lines)


def msbuild_props(manifest, assets):
    hashes = [(entry["path"], entry["sha256"]) for entry in manifest["files"]]
    hashes.extend([("../supervisor.c", manifest["sourceSha256"]),
                   ("manifest.json", sha256(assets / "manifest.json")),
                   ("NativeSupervisorAssetManifest.g.cs", hashlib.sha256(csharp(manifest).encode()).hexdigest())])
    hashes.extend(("../../../" + item["path"], item["sha256"]) for item in manifest["licenses"])
    project = ET.Element("Project")
    properties = ET.SubElement(project, "PropertyGroup")
    ET.SubElement(properties, "ProcessSupervisorAssetsPublishable").text = str(manifest["publishable"]).lower()
    group = ET.SubElement(project, "ItemGroup")
    for path, digest in hashes:
        item = ET.SubElement(group, "ProcessSupervisorAssetHash", Include="$(MSBuildThisFileDirectory)" + path)
        ET.SubElement(item, "ExpectedHash").text = digest.upper()
    ET.indent(project, space="  ")
    return '<!-- Generated by build-process-supervisor.py; do not edit. -->\n' + ET.tostring(project, encoding="unicode") + "\n"


def verify_execution_evidence(manifest, evidence):
    require(evidence.get("sourceSha256") == manifest["sourceSha256"], "execution evidence refers to another source")
    checks = evidence.get("checks", [])
    expected = {("linux-x64", "glibc"), ("linux-x64", "musl"),
                ("linux-arm64", "glibc"), ("linux-arm64", "musl"),
                ("osx-x64", "macos-rosetta"), ("osx-arm64", "macos-native")}
    require(len(checks) == 6 and {(c["rid"], c["environment"]) for c in checks} == expected,
            "execution evidence must include all four ABIs and both Linux libc hosts")
    hashes = {entry["rid"]: entry["sha256"] for entry in manifest["files"]}
    cases = {(0, False, False), (7, False, False), (0, True, False), (7, True, False), (0, False, True)}
    for check in checks:
        require(check["binarySha256"] == hashes[check["rid"]], "execution evidence binary hash mismatch")
        results = check.get("results", [])
        require(len(results) == 5 and {(r["targetExit"], r["ignoredSigchld"], r["missingTarget"]) for r in results} == cases,
                "execution evidence is missing 0/7, ignored SIGCHLD, or ENOENT cases")
        require(check.get("command") and check.get("uname"), "execution environment/command is missing")
        for result in results:
            require(result.get("helperExitCode") == 0 and result.get("cleanupSuccess") is True,
                    "execution evidence contains a failed or unknown cleanup result")
            first, last = result["first"], result["last"]
            require(first[0] == last[0] == 0x4F465301 and first[2] > 0 and first[2] == last[2], "invalid protocol identity in evidence")
            require(last[1] == 4 and last[3] == 0, "missing successful terminal cleanup frame")
            require(first[1] == (3 if result["missingTarget"] else 1), "invalid first protocol frame")
            if result["missingTarget"]:
                require(first[3] == 2, "ENOENT was not preserved")
            else:
                require(result.get("exitFrame") == [0x4F465301, 2, first[2], result["targetExit"]], "real target exit status was not preserved")


def verify(assets=ASSETS, require_publishable=True):
    manifest = json.loads((assets / "manifest.json").read_text())
    require(manifest.get("schemaVersion") == 1 and manifest.get("protocolVersion") == 1, "unsupported manifest/protocol")
    require(not require_publishable or manifest.get("publishable") is True, "asset set is not marked publishable")
    require(manifest.get("sourcePath") == "native/Ofdrw.ProcessSupervisor/supervisor.c", "unexpected source path")
    require(sha256(ROOT / manifest["sourcePath"]) == manifest["sourceSha256"], "native source changed: rebuild all four assets from the frozen SHA")
    files = manifest.get("files", [])
    require(len(files) == 4 and sorted(f["rid"] for f in files) == sorted(RIDS), "all four unique RIDs are required")
    for entry in files:
        rid = entry["rid"]
        require(entry["path"] == f"{rid}/ofdrw-supervisor", f"{rid}: unexpected asset path")
        path = assets / entry["path"]
        require(path.is_file() and not path.is_symlink(), f"{rid}: missing asset or symlink")
        data = path.read_bytes()
        require(len(data) == entry["size"] and sha256(path) == entry["sha256"], f"{rid}: size/hash mismatch")
        require(inspect_header(data, rid) == entry["header"], f"{rid}: recorded native headers disagree")
        require(entry["reproducible"]["firstSha256"] == entry["sha256"] == entry["reproducible"]["secondSha256"], f"{rid}: rebuild mismatch")
        require(entry.get("buildCommand") and entry.get("toolchain") and entry.get("minimumOS"), f"{rid}: incomplete provenance")
        if rid.startswith("linux-"):
            require(entry["toolchain"]["image"] == LINUX_IMAGES[rid] and entry["toolchain"]["muslVersion"] == "1.2.5", f"{rid}: toolchain not pinned")
            require(entry["toolchain"]["muslSourceSha256"] == MUSL_SOURCE_SHA256, f"{rid}: incorrect musl source")
            require(entry["toolchain"]["compilerRuntimeLinked"] is False, f"{rid}: unexpected compiler runtime")
        else:
            expected_uuid = hashlib.sha256((manifest["sourceSha256"] + ":" + rid).encode()).digest()[:16].hex()
            require(entry["header"]["uuid"] == expected_uuid, f"{rid}: UUID not bound to source and RID")
            require(entry["toolchain"]["compiler"] == "Apple clang version 21.0.0 (clang-2100.1.1.101)" and entry["toolchain"]["sdkBuild"] == "25F70", f"{rid}: unexpected Apple compiler/SDK build")
        print(f"{rid}: {len(data)} bytes {entry['sha256']}")
    licenses = manifest.get("licenses", [])
    require(any(item["path"] == "docs/licenses/MUSL-COPYRIGHT.txt" for item in licenses), "missing musl license")
    require(sha256(ROOT / "docs/licenses/MUSL-COPYRIGHT.txt") == MUSL_NOTICE_SHA256, "musl notice differs from exact reviewed upstream bytes and individual source appendix")
    for item in licenses:
        require(sha256(ROOT / item["path"]) == item["sha256"], f"license mismatch: {item['path']}")
    require("musl 1.2.5" in (ROOT / "THIRD-PARTY-NOTICES.md").read_text(), "third party notices omit musl")
    if manifest.get("publishable") is True:
        verify_execution_evidence(manifest, manifest.get("executionEvidence", {}))
    require((assets / "NativeSupervisorAssetManifest.g.cs").read_text() == csharp(manifest), "generated C# differs from manifest")
    require((assets / "SupervisorAssetHashes.props").read_text() == msbuild_props(manifest, assets), "MSBuild hash list differs from complete delivery")
    return manifest


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--assets", type=Path, default=ASSETS)
    parser.add_argument("--write-csharp", type=Path, help="emit a verified netstandard-compatible hash table for Core compilation")
    parser.add_argument("--allow-unpublished", action="store_true", help="maintainer review of a complete candidate set; does not make it publishable")
    args = parser.parse_args()
    try:
        manifest = verify(args.assets, require_publishable=not args.allow_unpublished)
        if args.write_csharp:
            content = csharp(manifest)
            args.write_csharp.parent.mkdir(parents=True, exist_ok=True)
            if not args.write_csharp.exists() or args.write_csharp.read_text() != content:
                args.write_csharp.write_text(content)
    except (OSError, ValueError, KeyError, TypeError, struct.error) as error:
        print(f"native asset verification failed: {error}", file=sys.stderr)
        sys.exit(1)
