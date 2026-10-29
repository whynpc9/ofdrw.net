# R22 bounded mixed-paragraph query candidate

This branch contains the query and final line-layout API only. It includes
main through the merged shared PDF bold repair PR #73
(`27f1944384eb3800d2ff21ff3ed63faeef16ef41`) and delivers logical source
intervals, final shaped glyphs and visual query-relative coordinates for the
declared Latin/Hebrew LTR profile. It makes no OFD, PDF, SVG, PNG or
reader-compatibility claim.

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
diagnostics but do not transfer to this branch's acceptance. The mixed-script
output writer, PDF reader copy/Find failures and their separate delivery gates
remain with draft PR #71; PR #73's shared bold converter repair is already in main.
