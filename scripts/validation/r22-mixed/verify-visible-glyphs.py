"""Verify the bounded one-font R22 fixture's actual OFD XML against PDF painted paths.

Requires pypdf 6.21.2 and fontTools 4.68.0. This is an independent validator
for generated R22 pages, not an arbitrary OFD/PDF parser or logical reader.
"""

import hashlib
import json
import sys
import xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from fontTools.pens.basePen import BasePen
from fontTools.ttLib import TTFont
from pypdf import PdfReader
from pypdf.generic import ContentStream

NS = {"o": "http://www.ofdspec.org"}


class CubicPen(BasePen):
    def __init__(self, gs):
        super().__init__(gs)
        self.ops = []

    def _moveTo(self, p):
        self.ops.append(("m", p))

    def _lineTo(self, p):
        self.ops.append(("l", p))

    def _curveToOne(self, p1, p2, p3):
        self.ops.append(("c", p1 + p2 + p3))

    def _qCurveToOne(self, p1, p2):
        p0 = self._getCurrentPoint()
        a = tuple(p0[i] + 2 * (p1[i] - p0[i]) / 3 for i in (0, 1))
        b = tuple(p2[i] + 2 * (p1[i] - p2[i]) / 3 for i in (0, 1))
        self._curveToOne(a, b, p2)

    def _closePath(self):
        self.ops.append(("h", ()))

    def _endPath(self):
        raise ValueError("open font contour")


def mul(a, b):
    return [
        a[0] * b[0] + a[2] * b[1],
        a[1] * b[0] + a[3] * b[1],
        a[0] * b[2] + a[2] * b[3],
        a[1] * b[2] + a[3] * b[3],
        a[0] * b[4] + a[2] * b[5] + a[4],
        a[1] * b[4] + a[3] * b[5] + a[5],
    ]


def transform(m, x, y):
    return (m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5])


def verify(ofd, pdf, fixture):
    assert (
        hashlib.sha256(Path(fixture).read_bytes()).hexdigest()
        == "d7441d42a24882db09640ea4bb768b17a106cdb31ee9749a52cb34a075b8fbcc"
    )
    original = TTFont(fixture)
    expected = []
    with ZipFile(ofd) as z:
        page = ET.fromstring(z.read("Doc_0/Pages/Page_0/Content.xml"))
        res = ET.fromstring(z.read("Doc_0/PublicRes.xml"))
        # Resource identity is resolved from the actual archive, never from rendered SVG metadata.
        fonts = res.findall(".//o:Font", NS)
        assert len(fonts) == 1, "Expected one bounded R22 font"
        fontnode = fonts[0]
        fontfile = fontnode.find("o:FontFile", NS).text
        base = res.attrib.get("BaseLoc", "Res")
        fontpath = "Doc_0/" + base.strip("/") + "/" + fontfile
        tt = TTFont(BytesIO(z.read(fontpath)))
        assert tt["head"].unitsPerEm == original["head"].unitsPerEm
        height = float(page.find(".//o:PhysicalBox", NS).text.split()[3])
        gs = tt.getGlyphSet()
        for t in page.findall(".//o:TextObject", NS):
            c = t.find("o:TextCode", NS)
            txt = c.text or ""
            if not any("\u05d0" <= ch <= "\u05ea" for ch in txt):
                continue
            assert (
                t.get("CTM") is None
                and t.get("HScale") is None
                and t.get("Font") == fontnode.get("ID")
            )
            assert (
                t.get("Fill", "true") == "true" and t.get("Stroke", "false") == "false"
            )
            boundary = list(map(float, t.attrib["Boundary"].split()))
            gids = list(map(int, t.find(".//o:Glyphs", NS).text.split()))
            x = float(c.get("X"))
            y = float(c.get("Y"))
            dx = list(map(float, c.get("DeltaX", "").split()))
            dy = list(map(float, c.get("DeltaY", "").split()))
            scale = float(t.get("Size")) / tt["head"].unitsPerEm
            assert len(gids) == len(txt) and len(dx) == len(gids) - 1
            for i, gid in enumerate(gids):
                name = tt.getGlyphName(gid)
                source = txt[-1 - i]
                assert original.getGlyphID(original.getBestCmap()[ord(source)]) == gid
                p = CubicPen(gs)
                gs[name].draw(p)
                original_pen = CubicPen(original.getGlyphSet())
                original.getGlyphSet()[original.getGlyphName(gid)].draw(original_pen)
                assert p.ops == original_pen.ops
                origin = (
                    (boundary[0] + x) * 72 / 25.4,
                    (height - boundary[1] - y) * 72 / 25.4,
                )
                ops = []
                for op, args in p.ops:
                    coords = []
                    for j in range(0, len(args), 2):
                        coords.extend(
                            (
                                origin[0] + args[j] * scale * 72 / 25.4,
                                origin[1] + args[j + 1] * scale * 72 / 25.4,
                            )
                        )
                    ops.append((op, tuple(coords)))
                expected.append(dict(source=source, gid=gid, origin=origin, ops=ops))
                if i < len(dx):
                    x += dx[i]
                if i < len(dy):
                    y += dy[i]
    reader = PdfReader(pdf)
    page = reader.pages[0]
    m = [1, 0, 0, 1, 0, 0]
    stack = []
    paths = []
    current = []
    assert len(reader.pages) == 1 and int(page.get("/Rotate", 0)) == 0
    alpha = 1.0
    color = (0.0, 0.0, 0.0)
    states = page["/Resources"].get("/ExtGState", {})
    for args, op in ContentStream(page.get_contents(), reader).operations:
        if op == b"q":
            stack.append((m[:], alpha, color))
        elif op == b"Q":
            m, alpha, color = stack.pop()
        elif op == b"gs":
            alpha = float(states[args[0]].get_object().get("/ca", alpha))
        elif op == b"rg":
            color = tuple(map(float, args))
        elif op == b"g":
            color = (float(args[0]),) * 3
        elif op == b"cm":
            m = mul(m, list(map(float, args)))
        elif op in (b"m", b"l", b"c"):
            coords = []
            for i in range(0, len(args), 2):
                coords.extend(transform(m, float(args[i]), float(args[i + 1])))
            current.append((op.decode(), tuple(coords)))
        elif op == b"h":
            current.append(("h", ()))
        elif op in (b"f", b"f*", b"B", b"B*", b"b", b"b*", b"S", b"s"):
            if current:
                assert op == b"f" and alpha == 1 and color == (0.0, 0.0, 0.0), (
                    "Expected opaque black nonzero fill for the fixed R22 glyphs"
                )
                paths.append(current)
                current = []
        elif op == b"n":
            current = []
    used = set()
    checks = []
    for e in expected:
        candidates = []
        for i, p in enumerate(paths):
            if (
                i in used
                or len(p) != len(e["ops"])
                or any(
                    a[0] != b[0] or len(a[1]) != len(b[1]) for a, b in zip(p, e["ops"])
                )
            ):
                continue
            error = max(
                [abs(x - y) for a, b in zip(p, e["ops"]) for x, y in zip(a[1], b[1])]
                or [0]
            )
            if error <= 0.002:
                candidates.append((i, error))
        ok = len(candidates) == 1
        if ok:
            used.add(candidates[0][0])
        checks.append(
            dict(
                source=e["source"],
                gid=e["gid"],
                origin=e["origin"],
                candidates=candidates,
                passCheck=ok,
            )
        )
    return dict(
        ofd=str(ofd),
        pdf=str(pdf),
        expectedGlyphs=len(expected),
        paintedPaths=len(paths),
        matchedPaths=len(used),
        allGlyphsMatched=bool(expected)
        and len(paths) == len(expected)
        and all(x["passCheck"] for x in checks),
        checks=checks,
        scope="Hebrew visible original-GID paths only; carriers not used as character evidence",
    )


if __name__ == "__main__":
    r = verify(Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]))
    Path(sys.argv[4]).write_text(json.dumps(r, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({k: v for k, v in r.items() if k != "checks"}))
    sys.exit(0 if r["allGlyphsMatched"] else 1)
