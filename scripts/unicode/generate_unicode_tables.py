#!/usr/bin/env python3
"""Generate the frozen Unicode property tables used by Ofdrw.Net.Layout text segmentation.

Usage:
    python3 -I scripts/unicode/generate_unicode_tables.py <ucd-dir> [output.cs]

<ucd-dir> holds the UCD files listed under "generatorInputs" in
ucd-17.0.0-manifest.json, either flat (by file name) or in their UCD
sub-directories. Every input is verified against its pinned SHA-256 before it
is parsed; a mismatch aborts without writing output. The generated file is
deterministic for identical inputs.
"""
import hashlib
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(HERE, "ucd-17.0.0-manifest.json")
DEFAULT_OUTPUT = os.path.join(HERE, "..", "..", "src", "Ofdrw.Net.Layout", "Text", "Unicode",
                              "UnicodeData.g.cs")

BIDI_CLASSES = ["L", "R", "AL", "EN", "ES", "ET", "AN", "CS", "NSM", "BN", "B", "S", "WS", "ON",
                "LRE", "LRO", "RLE", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI"]
BIDI_LONG = {
    "Left_To_Right": "L", "Right_To_Left": "R", "Arabic_Letter": "AL", "European_Number": "EN",
    "European_Separator": "ES", "European_Terminator": "ET", "Arabic_Number": "AN",
    "Common_Separator": "CS", "Nonspacing_Mark": "NSM", "Boundary_Neutral": "BN",
    "Paragraph_Separator": "B", "Segment_Separator": "S", "White_Space": "WS", "Other_Neutral": "ON",
    "Left_To_Right_Embedding": "LRE", "Left_To_Right_Override": "LRO",
    "Right_To_Left_Embedding": "RLE", "Right_To_Left_Override": "RLO",
    "Pop_Directional_Format": "PDF", "Left_To_Right_Isolate": "LRI",
    "Right_To_Left_Isolate": "RLI", "First_Strong_Isolate": "FSI", "Pop_Directional_Isolate": "PDI",
}
LINE_BREAK_CLASSES = ["XX", "BK", "CR", "LF", "CM", "NL", "SG", "WJ", "ZW", "GL", "SP", "ZWJ", "B2",
                      "BA", "BB", "HY", "HH", "CB", "CL", "CP", "EX", "IN", "NS", "OP", "QU", "IS",
                      "NU", "PO", "PR", "SY", "AI", "AK", "AL", "AP", "AS", "CJ", "EB", "EM", "H2",
                      "H3", "HL", "ID", "JL", "JV", "JT", "RI", "SA", "VF", "VI"]
GRAPHEME_CLASSES = ["Other", "CR", "LF", "Control", "Extend", "ZWJ", "Regional_Indicator",
                    "Prepend", "SpacingMark", "L", "V", "T", "LV", "LVT"]
INCB = ["None", "Consonant", "Linker", "Extend"]

# Line-break value flags (stored above the 6-bit class).
LB_PI, LB_PF, LB_EAST_ASIAN, LB_PICTOGRAPHIC_UNASSIGNED, LB_MARK = 64, 128, 256, 512, 1024


def sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for block in iter(lambda: stream.read(1 << 16), b""):
            digest.update(block)
    return digest.hexdigest()


def locate(root, relative):
    for candidate in (os.path.join(root, relative), os.path.join(root, os.path.basename(relative))):
        if os.path.isfile(candidate):
            return candidate
    raise SystemExit(f"missing UCD input {relative} under {root}")


def lines(path):
    with open(path, encoding="utf-8") as stream:
        for raw in stream:
            yield raw.rstrip("\n")


def parse_range(text):
    text = text.strip()
    if ".." in text:
        first, last = text.split("..")
        return int(first, 16), int(last, 16)
    value = int(text, 16)
    return value, value


def fill_property(path, default, mapper=lambda value: value, field=1, honor_missing=True):
    """Return a full code point array: @missing defaults first, then explicit data."""
    values = [default] * 0x110000
    explicit = []
    for line in lines(path):
        if line.startswith("# @missing:"):
            if not honor_missing:
                continue
            body = line[len("# @missing:"):].split("#")[0]
            parts = [part.strip() for part in body.split(";")]
            if len(parts) == 2:
                first, last = parse_range(parts[0])
                mapped = mapper(parts[1])
                for code in range(first, last + 1):
                    values[code] = mapped
            continue
        content = line.split("#", 1)[0].strip()
        if not content:
            continue
        parts = [part.strip() for part in content.split(";")]
        explicit.append((parse_range(parts[0]), parts[field] if field < len(parts) else None, parts))
    return values, explicit


def ranges(values):
    starts, data = [], []
    previous = None
    for code, value in enumerate(values):
        if value != previous:
            starts.append(code)
            data.append(value)
            previous = value
    return starts, data


def emit_array(name, kind, items, per_line=12, hexadecimal=False):
    out = [f"    internal static readonly {kind}[] {name} = new {kind}[]", "    {"]
    for offset in range(0, len(items), per_line):
        chunk = items[offset:offset + per_line]
        text = ", ".join((f"0x{item:X}" if hexadecimal else str(item)) for item in chunk)
        out.append(f"        {text},")
    out.append("    };")
    return out


def main():
    if len(sys.argv) not in (2, 3):
        raise SystemExit(__doc__)
    root = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) == 3 else DEFAULT_OUTPUT
    with open(MANIFEST, encoding="utf-8") as stream:
        manifest = json.load(stream)
    paths = {}
    for relative, expected in manifest["generatorInputs"].items():
        path = locate(root, relative)
        actual = sha256(path)
        if actual != expected:
            raise SystemExit(f"{relative}: SHA-256 {actual} differs from pinned {expected}")
        paths[os.path.basename(relative)] = path

    # General category (used only for the UAX #14 LB1/LB15/LB19/LB30b flags).
    category = ["Cn"] * 0x110000
    for line in lines(paths["DerivedGeneralCategory.txt"]):
        content = line.split("#", 1)[0].strip()
        if not content:
            continue
        code_range, value = [part.strip() for part in content.split(";")]
        first, last = parse_range(code_range)
        for code in range(first, last + 1):
            category[code] = value

    # Bidi_Class.
    bidi, explicit = fill_property(paths["DerivedBidiClass.txt"], "L", lambda value: BIDI_LONG[value])
    for (first, last), value, _ in explicit:
        for code in range(first, last + 1):
            bidi[code] = value
    bidi_values = [BIDI_CLASSES.index(value) for value in bidi]

    # Bidi_Paired_Bracket and Bidi_Paired_Bracket_Type.
    brackets = []
    for line in lines(paths["BidiBrackets.txt"]):
        content = line.split("#", 1)[0].strip()
        if not content:
            continue
        code, pair, kind = [part.strip() for part in content.split(";")]
        brackets.append((int(code, 16), int(pair, 16), 1 if kind == "o" else 0))

    # Line_Break plus the property flags consumed by the tailorable rules.
    line_break, explicit = fill_property(paths["LineBreak.txt"], "XX")
    for (first, last), value, _ in explicit:
        for code in range(first, last + 1):
            line_break[code] = value
    east_asian, explicit = fill_property(paths["EastAsianWidth.txt"], "N")
    for (first, last), value, _ in explicit:
        for code in range(first, last + 1):
            east_asian[code] = value
    pictographic = [False] * 0x110000
    for line in lines(paths["emoji-data.txt"]):
        content = line.split("#", 1)[0].strip()
        if not content:
            continue
        code_range, prop = [part.strip() for part in content.split(";")]
        if prop != "Extended_Pictographic":
            continue
        first, last = parse_range(code_range)
        for code in range(first, last + 1):
            pictographic[code] = True
    line_values = []
    for code in range(0x110000):
        value = LINE_BREAK_CLASSES.index(line_break[code])
        gc = category[code]
        if gc == "Pi":
            value |= LB_PI
        if gc == "Pf":
            value |= LB_PF
        if east_asian[code] in ("F", "W", "H"):
            value |= LB_EAST_ASIAN
        if pictographic[code] and gc == "Cn":
            value |= LB_PICTOGRAPHIC_UNASSIGNED
        if gc in ("Mn", "Mc"):
            value |= LB_MARK
        line_values.append(value)

    # Grapheme_Cluster_Break, Indic_Conjunct_Break and Extended_Pictographic.
    grapheme, explicit = fill_property(paths["GraphemeBreakProperty.txt"], "Other")
    for (first, last), value, _ in explicit:
        for code in range(first, last + 1):
            grapheme[code] = value
    incb = ["None"] * 0x110000
    for line in lines(paths["DerivedCoreProperties.txt"]):
        content = line.split("#", 1)[0].strip()
        if not content:
            continue
        parts = [part.strip() for part in content.split(";")]
        if len(parts) == 3 and parts[1] == "InCB":
            first, last = parse_range(parts[0])
            for code in range(first, last + 1):
                incb[code] = parts[2]
    grapheme_values = [GRAPHEME_CLASSES.index(grapheme[code]) | (INCB.index(incb[code]) << 4) |
                       ((1 << 6) if pictographic[code] else 0) for code in range(0x110000)]

    # Script, as ISO 15924 codes.
    aliases = {}
    for line in lines(paths["PropertyValueAliases.txt"]):
        match = re.match(r"^sc\s*;\s*(\w+)\s*;\s*(\w+)", line)
        if match:
            aliases[match.group(2)] = match.group(1)
    script, explicit = fill_property(paths["Scripts.txt"], "Unknown")
    for (first, last), value, _ in explicit:
        for code in range(first, last + 1):
            script[code] = value
    script_codes = sorted({aliases[value] for value in script} | {"Zyyy", "Zinh", "Zzzz"})
    script_values = [script_codes.index(aliases[value]) for value in script]

    out = [
        "// <auto-generated>",
        "// Generated by scripts/unicode/generate_unicode_tables.py from the Unicode Character",
        f"// Database {manifest['unicodeVersion']}. Inputs are pinned by SHA-256 in",
        "// scripts/unicode/ucd-17.0.0-manifest.json. Unicode License v3; see",
        "// docs/licenses/UNICODE-LICENSE.txt. Do not edit by hand.",
        "// </auto-generated>",
        "namespace Ofdrw.Net.Layout.Text.Unicode;",
        "",
        "internal static class UnicodeData",
        "{",
        f"    internal const string Version = \"{manifest['unicodeVersion']}\";",
        "",
        "    internal static readonly string[] BidiClassNames = { " +
        ", ".join(f'"{value}"' for value in BIDI_CLASSES) + " };",
        "    internal static readonly string[] LineBreakClassNames = { " +
        ", ".join(f'"{value}"' for value in LINE_BREAK_CLASSES) + " };",
        "    internal static readonly string[] GraphemeClassNames = { " +
        ", ".join(f'"{value}"' for value in GRAPHEME_CLASSES) + " };",
        "    internal static readonly string[] ScriptCodes = { " +
        ", ".join(f'"{value}"' for value in script_codes) + " };",
        "",
    ]
    for prefix, values, kind in (("Bidi", bidi_values, "byte"), ("LineBreak", line_values, "ushort"),
                                 ("Grapheme", grapheme_values, "byte"), ("Script", script_values, "byte")):
        starts, data = ranges(values)
        out += emit_array(prefix + "Starts", "int", starts, hexadecimal=True)
        out += emit_array(prefix + "Values", kind, data, per_line=24)
        out.append("")
    out += emit_array("BracketCodes", "int", [item[0] for item in brackets], hexadecimal=True)
    out += emit_array("BracketPairs", "int", [item[1] for item in brackets], hexadecimal=True)
    out += emit_array("BracketOpening", "byte", [item[2] for item in brackets], per_line=24)
    out.append("}")
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    with open(output, "w", encoding="utf-8", newline="\n") as stream:
        stream.write("\n".join(out) + "\n")
    print(f"wrote {output}")


if __name__ == "__main__":
    main()
