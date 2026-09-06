#!/usr/bin/env python3
"""Install a pinned, real CJK TrueType face on disposable CI runners.

Requires fonttools==4.59.2. Use --directory for an isolated local font directory
instead of installing into a system-discovered location.
"""

import argparse
import hashlib
import os
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
    args = parser.parse_args()
    directory = args.directory or default_directory()
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
    if not args.directory and sys.platform.startswith("linux"):
        subprocess.run(["fc-cache", "-f", str(directory)], check=True)
    print(f"Installed Noto Sans CJK SC Regular: {target} ({target.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
