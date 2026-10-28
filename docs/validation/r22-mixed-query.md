# R22 bounded mixed-paragraph query candidate

This branch contains the query and final line-layout API only. It is based on
PR #71's current base `3a31a2cefda73f993b4ba91cd7ff8fc896777637` and is
intended to deliver logical source intervals, final shaped glyphs and visual
query-relative coordinates for the declared Latin/Hebrew LTR profile. It makes no OFD, PDF,
SVG, PNG or reader-compatibility claim.

The HarfBuzz E2E query checks one Hebrew word, two adjacent Hebrew words,
repeated words, separated RTL phrases, leading and trailing RTL words, exact
fit, and soft wraps immediately before and after an RTL phrase. Each case
checks source continuity, visual run order, within-word glyph clusters,
non-overlapping run geometry, trailing space ownership and width. It also
checks missing font coverage, punctuation refusal, shaping budget and
cancellation. The DejaVu Sans 2.37 derived fixture, source hash and license
are recorded in the shared font manifest and notices.

The branch must pass its own final-SHA E2E, isolated package consumer and
cross-platform checks before review. Earlier PR #71 results are useful
diagnostics but do not transfer to this branch's acceptance. The output writer,
shared packaging/converter changes, Preview copy/Find failures and their
separate delivery gates remain with draft PR #71.
