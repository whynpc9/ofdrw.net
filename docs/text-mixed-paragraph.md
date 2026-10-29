# Bounded mixed Latin/Hebrew paragraph query

`HarfBuzzTextPrelayout.CreateMixedParagraph()` supplies the pinned HarfBuzz run
shaper to `TextMixedParagraphEngine`. The separate `TextMixedParagraphRequest`
accepts one LTR paragraph of 3–128 BMP source characters: ASCII Latin words,
Hebrew letters U+05D0–U+05EA, and single ASCII spaces between words. It requires
both scripts, one caller-supplied immutable regular static TrueType face covering
all characters, horizontal layout and finite positive size/width. A script
change inside a word, controls, combining marks, digits, punctuation, bidi
format controls, supplementary characters, variation selectors, ZWJ and other
scripts fail explicitly. It makes no general Unicode bidi claim.

The query identifies word and space source spans, shapes Latin words LTR and
Hebrew words RTL with absolute UTF-16 source clusters, and places the resulting
visual glyph sequences in an LTR paragraph. Adjacent Hebrew words and their
internal spaces form one visual RTL phrase: word positions reverse within that
phrase while the result's run list remains in logical source order. It chooses
breaks only after whole words and reshapes every final word/space. A soft
break retains its trailing space in the source interval and total advance;
visible fit excludes that space. The final result owns the exact shaped runs
returned to the caller, including the selected face, engine version,
script, direction, GIDs, offsets and advances. It never substitutes a host
font or falls back to approximate character widths.

The query bounds font size, width, line count, shape calls and cumulative font
copy work. Missing glyphs, overlong words, malformed shape results, budget
exhaustion and cancellation fail without a partial result. The fixed R22 DejaVu Sans
2.37 fixture is derived from the upstream release archive with its license and byte
hash in `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/upstream-manifest.json`.
The E2E checks source spans, descending Hebrew HarfBuzz clusters, legal soft
wraps, coverage failure, unsupported punctuation, budget and cancellation.

The public API in this delivery is query-only. It returns source-addressed shaped
runs and query-relative coordinates; it does not create an OFD, PDF, SVG or PNG.
Output support and reader compatibility remain in draft PR #71. In particular,
a successful query must not be treated as permission to serialize the result
through an unrelated writer. The [query validation record](validation/r22-mixed-query.md)
tracks this candidate's source and package gates.
