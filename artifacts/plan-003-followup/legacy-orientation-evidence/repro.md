# Existing BuiltIn source landscape setup

The exploratory control used exactly the same mixed model builder as the final
`DifferentPageSizesKeepHeadersFieldsTablesAndImagesThroughDrawingSeam` regression,
except its second section tuple was `(420d, 300d, "Landscape")` rather than the
final `(320d, 440d, "SecondPortrait")`. The source model width/height were420/300.
Both its no-italic control (original MigraDoc RenderPages path) and styled
adapter path yielded page2 MediaBox300x420, Rotate0. The retained PDFs, pdfinfo
and manifest record this observation. They are negative evidence, not accepted
landscape output. The production ApplyPageSetup method has zero diff from the
frozen3247c76f base. The narrow correction required for source page setup is a
separate follow-up; this task does not change it.

The final prepared MigraDoc test bypasses this source-setup issue, sets neutral
width/height300/420 with explicit Landscape for its second section, and proves
original RenderPages and the production adapter both output actual420x300 pages,
with unchanged glyph anchors and exactly one middle-run italic shear.
