"""Independent scalar/CID/GID/outline/baseline check for the fixed R22 corpus."""

import hashlib
import json
import re
import sys
from io import BytesIO
from pathlib import Path

from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.ttLib import TTFont
from pypdf import PdfReader
from pypdf.generic import ContentStream


def sha(data):
    return hashlib.sha256(data).hexdigest()


def outline(font, name):
    pen = DecomposingRecordingPen(font.getGlyphSet())
    font.getGlyphSet()[name].draw(pen)
    return pen.value


def unicode_map(data):
    text = data.decode("latin-1")
    result = {}
    for group in re.findall(r"beginbfchar(.*?)endbfchar", text, re.S):
        for a, b in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", group):
            result[int(a, 16)] = bytes.fromhex(b).decode("utf-16-be")
    for group in re.findall(r"beginbfrange(.*?)endbfrange", text, re.S):
        for a, b, c in re.findall(
            r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", group
        ):
            for cid in range(int(a, 16), int(b, 16) + 1):
                result[cid] = chr(int(c, 16) + cid - int(a, 16))
    return result


def width(font, cid):
    d = font["/DescendantFonts"][0].get_object()
    W = d.get("/W", [])
    i = 0
    while i < len(W):
        start = int(W[i])
        v = W[i + 1]
        i += 2
        if isinstance(v, list):
            if start <= cid < start + len(v):
                return float(v[cid - start])
        else:
            w = float(W[i])
            i += 1
            if start <= cid <= int(v):
                return w
    return float(d.get("/DW", 1000))


def inspect(item, pig):
    path = Path(item["path"])
    assert sha(path.read_bytes()) == item["sha256"]
    reader = PdfReader(path)
    page = reader.pages[0]
    assert int(page.get("/Rotate", 0)) == 0
    assert float(page.mediabox.left) == float(page.mediabox.bottom) == 0
    fonts = {}
    for name, ref in page["/Resources"]["/Font"].items():
        f = ref.get_object()
        if "/DescendantFonts" not in f:
            continue
        d = f["/DescendantFonts"][0].get_object()
        if "/FontFile2" not in d["/FontDescriptor"]:
            continue
        data = d["/FontDescriptor"]["/FontFile2"].get_data()
        tt = TTFont(BytesIO(data))
        assert tt["head"].unitsPerEm == original["head"].unitsPerEm
        gidmap = d.get("/CIDToGIDMap", "/Identity")
        gidmap = gidmap.get_object() if hasattr(gidmap, "get_object") else gidmap
        fonts[str(name)] = (
            f,
            tt,
            unicode_map(f["/ToUnicode"].get_data()),
            None if str(gidmap) == "/Identity" else gidmap.get_data(),
            sha(data),
        )
    # References admit translation only; reject affine transforms rather than silently ignore.
    x = y = lx = ly = 0.0
    size = 0
    selected = None
    ctm = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    stack = []
    draws = []

    def show(s):
        nonlocal x
        raw = s.original_bytes if hasattr(s, "original_bytes") else bytes(s)
        if selected not in fonts:
            return
        f, tt, cmap, gidmap, fontsha = fonts[selected]
        assert len(raw) % 2 == 0
        for i in range(0, len(raw), 2):
            cid = int.from_bytes(raw[i : i + 2], "big")
            value = cmap.get(cid, "")
            if any("\u05d0" <= c <= "\u05ea" for c in value):
                gid = (
                    cid
                    if gidmap is None
                    else int.from_bytes(gidmap[2 * cid : 2 * cid + 2], "big")
                )
                name = tt.getGlyphName(gid)
                baseline = [x + ctm[4], y + ctm[5]]
                same = [
                    l
                    for l in pig["letters"]
                    if l["Value"] == value
                    and abs(l["baseline"]["x"] - baseline[0]) <= 0.1
                    and abs(l["baseline"]["y"] - baseline[1]) <= 0.1
                ]
                nominal = original_cmap.get(ord(value)) if len(value) == 1 else None
                draws.append(
                    {
                        "unicode": value,
                        "cid": cid,
                        "embeddedGid": gid,
                        "originalGid": original.getGlyphID(nominal)
                        if nominal
                        else None,
                        "baseline": baseline,
                        "outlineEqual": bool(nominal)
                        and outline(tt, name) == outline(original, nominal),
                        "advanceEqual": bool(nominal)
                        and tt["hmtx"][name][0] == original["hmtx"][nominal][0],
                        "pdfWidth": width(f, cid),
                        "nominalPdfWidth": 1000
                        * tt["hmtx"][name][0]
                        / tt["head"].unitsPerEm,
                        # PDF integer widths may round or truncate 1000-em metrics.
                        "pdfWidthNominal": abs(
                            width(f, cid)
                            - 1000 * tt["hmtx"][name][0] / tt["head"].unitsPerEm
                        )
                        < 1,
                        "pdfpigMatchCount": len(same),
                        "embeddedFontSha256": fontsha,
                    }
                )
            x += width(f, cid) * size / 1000

    for args, op in ContentStream(page.get_contents(), reader).operations:
        if op == b"q":
            stack.append(ctm)
        elif op == b"Q":
            ctm = stack.pop()
        elif op == b"cm":
            a, b, c, d, e, f = map(float, args)
            assert (a, b, c, d) == (1, 0, 0, 1), (item["role"], args)
            ctm = (1, 0, 0, 1, ctm[4] + e, ctm[5] + f)
        elif op == b"BT":
            x = y = lx = ly = 0.0
        elif op == b"Tf":
            selected = str(args[0])
            size = float(args[1])
        elif op == b"Tm":
            a, b, c, d, x, y = map(float, args)
            assert (a, b, c, d) == (1, 0, 0, 1)
            lx = x
            ly = y
        elif op in [b"Td", b"TD"]:
            lx += float(args[0])
            ly += float(args[1])
            x = lx
            y = ly
        elif op == b"Tj":
            show(args[0])
        elif op == b"TJ":
            for v in args[0]:
                if isinstance(v, (int, float)):
                    x -= float(v) * size / 1000
                else:
                    show(v)
    expected = sum("\u05d0" <= c <= "\u05ea" for c in item["source"])
    return {
        "role": item["role"],
        "sha256": item["sha256"],
        "hebrewDrawCount": len(draws),
        "expectedCount": expected,
        "pass": len(draws) == expected
        and all(
            x["outlineEqual"]
            and x["advanceEqual"]
            and x["pdfWidthNominal"]
            and x["pdfpigMatchCount"] == 1
            for x in draws
        ),
        "draws": draws,
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Decode bounded Type0/TrueType scalar-to-original-GID draws and associate independent PdfPig baselines."
    )
    parser.add_argument("inputs", type=Path)
    parser.add_argument("pdfpig", type=Path)
    parser.add_argument("font", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument(
        "--record-controls",
        action="store_true",
        help="Record expected unsupported controls without a failing exit code",
    )
    args = parser.parse_args()
    assert (
        sha(args.font.read_bytes())
        == "d7441d42a24882db09640ea4bb768b17a106cdb31ee9749a52cb34a075b8fbcc"
    )
    original = TTFont(args.font)
    original_cmap = original.getBestCmap()
    inputs = json.loads(args.inputs.read_text())["inputs"]
    pig = json.loads(args.pdfpig.read_text())
    rows = []
    for item in inputs:
        raw = next(
            x
            for x in pig["rows"]
            if x["role"] == item["role"] and not x["useActualText"]
        )
        assert raw["sha256"] == item["sha256"]
        try:
            row = inspect(item, raw)
        except Exception as error:
            row = {"role": item["role"], "unsupported": str(error), "pass": False}
        rows.append(row)
    args.output.write_text(
        json.dumps({"rows": rows}, ensure_ascii=False, indent=2) + "\n"
    )
    print(
        json.dumps(
            [{k: v for k, v in row.items() if k != "draws"} for row in rows],
            ensure_ascii=False,
        )
    )

    sys.exit(0 if args.record_controls or all(row["pass"] for row in rows) else 1)
