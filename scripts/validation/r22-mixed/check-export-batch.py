#!/usr/bin/env python3
"""Independently check the exact package consumer's bounded mixed export pages."""

import argparse
import hashlib
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import fontTools
import pypdf
from fontTools.ttLib import TTFont
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[3]
HERE = Path(__file__).resolve().parent
FONT = ROOT / "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-R22.ttf"
FLAGS = [
    "--disable-build-servers",
    "-m:1",
    "/nodeReuse:false",
    "/p:UseSharedCompilation=false",
]
REQUIRED = {
    "wide",
    "short-words",
    "long-upper",
    "unequal-latin",
    "upper-leading",
    "exact-fit",
    "two-word",
    "repeated-word",
    "two-word-short",
    "two-word-swapped",
    "two-word-long",
    "two-word-exact-fit",
}


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def rectangle_match(a, b):
    def v(r, key):
        return float(r.get(key, r.get(key.capitalize())))

    aw, bw = v(a, "right") - v(a, "left"), v(b, "right") - v(b, "left")
    ah, bh = v(a, "top") - v(a, "bottom"), v(b, "top") - v(b, "bottom")
    if min(aw, bw, ah, bh) <= 0:
        return False
    dx = abs(v(a, "left") + v(a, "right") - v(b, "left") - v(b, "right")) / 2
    dy = abs(v(a, "bottom") + v(a, "top") - v(b, "bottom") - v(b, "top")) / 2
    ox = max(
        0, min(v(a, "right"), v(b, "right")) - max(v(a, "left"), v(b, "left"))
    ) / min(aw, bw)
    oy = max(
        0, min(v(a, "top"), v(b, "top")) - max(v(a, "bottom"), v(b, "bottom"))
    ) / min(ah, bh)
    return dx <= 1 and dy <= 2 and ox >= 0.95 and oy >= 0.60


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "pages", type=Path, help="Exact consumer text-prelayout output directory"
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--pdfkit",
        action="store_true",
        help="Also run native macOS PDFKit whole selection/Find",
    )
    args = parser.parse_args()
    pages, out = args.pages.resolve(), args.output.resolve()
    out.mkdir(parents=True, exist_ok=True)
    assert pypdf.__version__ == "6.21.2" and fontTools.__version__ == "4.68.0"
    receipt = json.loads((pages / "r22-mixed-output-receipt.json").read_text())
    assert {r["name"].removeprefix("r22-mixed-") for r in receipt["cases"]} == REQUIRED
    for item in receipt["outputs"]:
        assert sha(pages / item["name"]) == item["sha256"], item["name"]
    cases, inputs = [], []
    for item in receipt["cases"]:
        pdf = pages / (item["name"] + ".pdf")
        source = item["source"]
        words = source.split(" ")
        cases.append(
            dict(
                name=item["name"],
                path=str(pdf),
                sha256=sha(pdf),
                expected=source,
                expectedSelection=[source],
                expectedWords=[
                    dict(text=w, count=words.count(w))
                    for w in dict.fromkeys(words)
                    if any("\u05d0" <= c <= "\u05ea" for c in w)
                ],
            )
        )
        inputs.append(
            dict(role=item["name"], path=str(pdf), sha256=sha(pdf), source=source)
        )
    (out / "inputs.json").write_text(
        json.dumps({"inputs": inputs}, ensure_ascii=False, indent=2) + "\n"
    )
    (out / "matrix.json").write_text(
        json.dumps({"cases": cases}, ensure_ascii=False, indent=2) + "\n"
    )
    env = os.environ.copy()
    env.update(
        DOTNET_CLI_HOME=str(out / "dotnet-home"),
        DOTNET_SKIP_FIRST_TIME_EXPERIENCE="1",
        DOTNET_CLI_TELEMETRY_OPTOUT="1",
    )
    Path(env["DOTNET_CLI_HOME"]).mkdir(exist_ok=True)
    project = HERE / "character-probe/PdfPigProbe.csproj"
    with (out / "reader-build.log").open("wb") as log:
        subprocess.run(
            [
                "dotnet",
                "build",
                str(project),
                "-c",
                "Release",
                "-o",
                str(out / "reader-bin"),
                *FLAGS,
            ],
            cwd=ROOT,
            env=env,
            stdout=log,
            stderr=subprocess.STDOUT,
            check=True,
        )
    subprocess.run(
        [
            "dotnet",
            str(out / "reader-bin/PdfPigProbe.dll"),
            str(out / "inputs.json"),
            str(out / "pdfpig.json"),
        ],
        cwd=ROOT,
        env=env,
        check=True,
    )
    pig = json.loads((out / "pdfpig.json").read_text())
    characters = module("r22_characters", HERE / "verify-character-gids.py")
    characters.original = TTFont(FONT)
    characters.original_cmap = characters.original.getBestCmap()
    visible = module("r22_visible", HERE / "verify-visible-glyphs.py")
    char_rows, visible_rows, geometry = [], [], []
    poppler = shutil.which("pdftotext")
    assert poppler, "pdftotext is required"
    for item in inputs:
        name = item["role"]
        raw = next(
            r for r in pig["rows"] if r["role"] == name and not r["useActualText"]
        )
        assert raw["sha256"] == item["sha256"]
        char_rows.append(characters.inspect(item, raw))
        visible_rows.append(
            visible.verify(pages / (name + ".ofd"), Path(item["path"]), FONT)
        )
        bbox = subprocess.check_output([poppler, "-bbox", item["path"], "-"])
        (out / (name + "-bbox.xhtml")).write_bytes(bbox)
        height = float(PdfReader(item["path"]).pages[0].mediabox.height)
        boxes = [
            dict(
                text=w.text,
                left=float(w.attrib["xMin"]),
                right=float(w.attrib["xMax"]),
                bottom=height - float(w.attrib["yMax"]),
                top=height - float(w.attrib["yMin"]),
            )
            for w in ET.fromstring(bbox).iter()
            if w.tag.endswith("word")
        ]
        used, checks = set(), []
        physical = [
            w for w in raw["words"] if any("\u05d0" <= c <= "\u05ea" for c in w["Text"])
        ]
        for word in physical:
            matches = [
                i
                for i, b in enumerate(boxes)
                if i not in used and rectangle_match(word["bounds"], b)
            ]
            if len(matches) == 1:
                used.add(matches[0])
            checks.append(
                dict(
                    pdfpigWord=word, popplerIndices=matches, passCheck=len(matches) == 1
                )
            )
        expected_count = sum(
            w["count"]
            for w in next(c for c in cases if c["name"] == name)["expectedWords"]
        )
        geometry.append(
            dict(
                name=name,
                passCheck=len(checks) == expected_count
                and all(c["passCheck"] for c in checks),
                checks=checks,
                popplerWords=boxes,
            )
        )
    result = dict(
        sourceSha=receipt["sourceSha"],
        platform=sys.platform,
        versions=dict(
            pypdf=pypdf.__version__,
            fontTools=fontTools.__version__,
            poppler=subprocess.run(
                [poppler, "-v"], capture_output=True, text=True
            ).stderr.strip(),
            pdfpig=pig["assembly"],
            pdfpigAssemblySha256=pig["assemblySha256"],
        ),
        inputManifestSha256=sha(out / "inputs.json"),
        sourceOutputReceiptSha256=sha(pages / "r22-mixed-output-receipt.json"),
        candidateManifestSha256=sha(pages.parent / "package-manifest.json")
        if (pages.parent / "package-manifest.json").exists()
        else None,
        verifierHashes={
            name: sha(HERE / name)
            for name in (
                "check-export-batch.py",
                "verify-character-gids.py",
                "verify-visible-glyphs.py",
                "pdfkit-probe.swift",
                "character-probe/Program.cs",
                "character-probe/PdfPigProbe.csproj",
            )
        },
        characters=char_rows,
        visible=visible_rows,
        physicalWordGeometry=geometry,
        pdfkit="not run on this platform",
        Preview="not run by automation",
    )
    ok = (
        all(r["pass"] for r in char_rows)
        and all(r["allGlyphsMatched"] for r in visible_rows)
        and all(r["passCheck"] for r in geometry)
    )
    if args.pdfkit:
        assert sys.platform == "darwin"
        binary = out / "pdfkit-probe"
        subprocess.run(
            ["swiftc", str(HERE / "pdfkit-probe.swift"), "-o", str(binary)], check=True
        )
        subprocess.run(
            [str(binary), str(out / "matrix.json"), str(out / "pdfkit.json")],
            check=True,
        )
        kit = json.loads((out / "pdfkit.json").read_text())
        kit_checks = []
        for row in kit["rows"]:
            word_boxes = next(
                r["popplerWords"] for r in geometry if r["name"] == row["name"]
            )
            used = set()
            for find in row["finds"]:
                for match in find["matches"]:
                    a = match["bounds"][0]
                    candidates = [
                        i
                        for i, b in enumerate(word_boxes)
                        if i not in used and rectangle_match(a, b)
                    ]
                    if len(candidates) == 1:
                        used.add(candidates[0])
                    kit_checks.append(
                        dict(
                            name=row["name"],
                            query=find["query"],
                            passCheck=len(candidates) == 1,
                        )
                    )
            ok &= row["exactSource"] and all(
                f["count"] == f["expectedCount"] for f in row["finds"]
            )
        ok &= all(r["passCheck"] for r in kit_checks)
        result["pdfkit"] = dict(
            version=kit["pdfkitVersion"], checks=kit_checks, raw="pdfkit.json"
        )
    result["status"] = "PASS" if ok else "FAIL"
    (out / "receipt.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    )
    print(
        f"{result['status']}: {len(cases)} mixed pages, {sum(r['expectedGlyphs'] for r in visible_rows)} original glyphs"
    )
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
