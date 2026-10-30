# R22 bounded two-word export checkpoint

This is the historical checkpoint leading to PR #76. For the source-baseline
reconciliation and the subsequent repeated-letter candidate, see
[remaining scope](r22-remaining-scope.md). The dated samples and hashes below
are not reassigned to later output.

Status: candidate implementation; final source, isolated-package, Preview and
CI acceptance are in progress. This does not close the full R22 roadmap item.

`TextMixedParagraphOfdWriter.CreateSinglePage` now has a candidate profile for
one or two adjacent Hebrew words between two ASCII Latin words, on one
unwrapped LTR line. It retains the exact R22 DejaVu Sans face at 4 mm, 1–8
characters per Latin run and 2–8 distinct letters within each Hebrew word.
Repeating a complete word is supported by the candidate; repeating a letter
within a word, other fonts/sizes, leading/trailing Hebrew, separated RTL
phrases, three words and wrapping remain rejected before creating output.
The broader paragraph query remains separate.

OFD TextObjects and TextCodes retain logical source order. `CGTransform` and
explicit glyph positions preserve the original GIDs. The PDF exporter paints
those original outlines and emits one scalar at each glyph's actual origin,
in visual order. It no longer replaces a Hebrew word with one semantic CID or
changes that CID's nominal font width. Shared characters and words at different
advances can reuse truthful per-character mappings.

Only consecutive verified RTL words separated by an actually empty font-space
glyph may have their PDF paint order reversed. The words must share a baseline
and font reference; transformed, clipped, stroked, scaled or overlapping ink is
excluded. Disjointness uses actual original outline control-point bounds,
including glyphs extending beyond declared OFD Boundaries. Other graphic
objects stop the phrase. This changes PDF content order without changing the
OFD logical source or moving any visible glyph.

The [replacement reader protocol](r22-reader-replacement-candidate.md) was
validated against the eleven hash-fixed existing references, including the
five additional samples. The reference character check independently matched
nominal original glyph outlines/advances and PdfPig baselines. Positive
logical text, logical Find and independent word rectangles passed; all declared
negative classifications stayed negative. The earlier PDFKit logical-character
index association remains falsified at 0/8 twice. These historical references
were known inputs, not a blind test.

A diagnostic through public positioned-glyph models, OFD write/read and the
modified product exporter produced three exact two-word PDFs. All 21 visible
original glyph paths and independently decoded character mappings/baselines
passed. Terra `gpt-6.1-terra` High inspected all three exact PDFs in Preview and
six PDF/direct-SVG page rasters. Logical Find, reversed-query absence and both
repeated-word locations passed. Actual Preview Cmd+A/C → new TextEdit plain
text Cmd+V → UTF-8 save recovered each source exactly, without operator
normalization or reversal. This is the saved native paste result; the system
pasteboard's raw representation was not independently read. The three owned
PDFs were closed; an Open chooser remained and hidden windows could not be
globally enumerated. This evidence belongs to those frozen diagnostic bytes.

The high-level public API candidate subsequently generated twelve single-page
mixed samples, including distinct/repeated words, swapped order, two short
words, two eight-letter words and exact fit. All 77 visible original glyphs
matched independently; PDFKit returned exact source and the declared positive
and reversed Find counts for all twelve. These candidate results do not inherit
the diagnostic Preview gate. The final package samples require their own
functional, semantic, PNG/SVG and Preview checks.

The independent checks and pinned tool versions are described in
[`scripts/validation/r22-mixed`](../../scripts/validation/r22-mixed/README.md).
The initial geometry tolerances were fixed before inspecting product PDFs:
0.1 pt character baseline, 0.002 pt visible contour coordinates, 1/2 pt word
center, 95%/60% word overlap. Shifted, missing and duplicated visible-glyph
controls were rejected. No product outline is proved by a hidden carrier.

At source `6ac2cb74a78a79313a6c7d5f0434016ecac93b13`, the full local
macOS Release suite passed 4,424 tests with zero failures/skips; the separate
mandatory trusted-font gate passed 74/74. The API documentation rebuild passed.
One twelve-package candidate, version `0.1.0-r22twoword.1`, was built once and
consumed through fresh package caches. Its immutable package-manifest SHA-256
is `069c0d6a511a7053c882e2c4f8370ea5e232b158f6712058ad5ade739f9ffb30`.
The selected eight runtime DLLs were checked against their nupkg payloads and
isolated cache, including `lib/netstandard2.1` and the actual HarfBuzz load.
The exact package outputs separately passed all twelve PDFKit logical-reader
and physical-word tests and all 77 scalar/GID/baseline/visible-contour checks.
Those are local functional and headless results, not final package Preview
acceptance or the later final-main cross-platform matrix.

PDF→OFD still depends on PdfPig default word extraction in its semantic-layer
and text-fallback paths. The export reader decision does not fix or accept RTL
reverse import. Unknown imported mappings/R02, R20, general bidi and other
scripts remain outside this increment. No tag, release or public NuGet
publication is authorized here. The old `4d9f003a` matrix and package manifest
`2bbe2b2d9638be438ae495947855b38a756bfa8827883b08ef4f25370f223c7d`
retain their original identities.
