#!/usr/bin/env python3
"""Install a pinned, real CJK TrueType face on disposable CI runners.

Requires fonttools==4.59.2. Use --directory for an isolated local font directory
instead of installing into a system-discovered location.
"""

import argparse
import hashlib
import os
import shutil
from pathlib import Path
import subprocess
import sys
import tempfile
from urllib.request import urlopen

from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont


REVISION = "f8d157532fbfaeda587e826d4cd5b21a49186f7c"
BASE_URL = f"https://raw.githubusercontent.com/notofonts/noto-cjk/{REVISION}/Sans"
SHA256 = "990c807e79c25662a5a9ecf7f971baeb2bf2eab9a559e5ecf15cdfdb8561d21f"

# Generated with the pinned fonttools version and wght=400 below. Cache hits
# are checked against source-controlled hashes, never a cache-supplied manifest.
OUTPUT_HASHES = {
    "Ofdrw-CI-NotoSansCJKsc-Regular.ttf": "3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa",
    "Ofdrw-CI-Noto-OFL.txt": "6a73f9541c2de74158c0e7cf6b0a58ef774f5a780bf191f2d7ec9cc53efe2bf2",
}


def valid_cache(directory):
    return all((directory / name).is_file() and
               hashlib.sha256((directory / name).read_bytes()).hexdigest() == digest
               for name, digest in OUTPUT_HASHES.items())


def install_cached(cache, directory, register):
    directory.mkdir(parents=True, exist_ok=True)
    for name in OUTPUT_HASHES:
        if (cache / name).resolve() != (directory / name).resolve():
            shutil.copyfile(cache / name, directory / name)
    if register and sys.platform.startswith("linux"):
        subprocess.run(["fc-cache", "-f", str(directory)], check=True)
    print(f"Installed checksum-verified cached CI font: {directory}")


def default_directory():
    if sys.platform == "win32":
        return Path(os.environ["WINDIR"]) / "Fonts"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Fonts"
    return Path.home() / ".fonts"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--directory", type=Path)
    parser.add_argument("--source", type=Path, help="Use an already downloaded, checksum-verified source")
    parser.add_argument("--cache-dir", type=Path, help="Reuse checksum-verified static font and license")
    args = parser.parse_args()
    directory = args.directory or default_directory()
    if args.cache_dir and not args.source and valid_cache(args.cache_dir):
        install_cached(args.cache_dir, directory, not args.directory)
        return
    with tempfile.TemporaryDirectory(prefix="ofdrw-ci-font-") as temporary:
        source = args.source or Path(temporary) / "source.ttf"
        if not args.source:
            with urlopen(f"{BASE_URL}/Variable/TTF/NotoSansCJKsc-VF.ttf", timeout=120) as response:
                source.write_bytes(response.read())
        if hashlib.sha256(source.read_bytes()).hexdigest() != SHA256:
            raise ValueError("CI font SHA-256 mismatch")

        # PDFsharp needs static TrueType outlines. Pin Regular instead of using
        # the variable font's default Thin weight; bold/italic remain simulated.
        with TTFont(source, recalcTimestamp=False) as variable:
            font = instantiateVariableFont(variable, {"wght": 400}, inplace=True, updateFontNames=True)
            required = set(map(ord, "文档转换视觉回归样例红Alpha"))
            if not required.issubset(font.getBestCmap()):
                raise ValueError("CI font is missing required Latin/CJK glyphs")
            if "fvar" in font or "glyf" not in font:
                raise ValueError("CI font must contain static TrueType outlines")
            directory.mkdir(parents=True, exist_ok=True)
            target = directory / "Ofdrw-CI-NotoSansCJKsc-Regular.ttf"
            font.save(target)

        with urlopen(f"{BASE_URL}/LICENSE", timeout=60) as response:
            (directory / "Ofdrw-CI-Noto-OFL.txt").write_bytes(response.read())
    if not valid_cache(directory):
        raise ValueError("Generated CI font or license differs from the pinned output")
    if args.cache_dir:
        install_cached(directory, args.cache_dir, False)
    if not args.directory and sys.platform.startswith("linux"):
        subprocess.run(["fc-cache", "-f", str(directory)], check=True)
    print(f"Installed Noto Sans CJK SC Regular: {target} ({target.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
