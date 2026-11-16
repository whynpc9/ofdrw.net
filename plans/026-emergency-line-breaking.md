# Plan 026: Enforce the accepted-line budget and preserve emergency break choices

## Scope and baseline

- Priority P2, medium risk; depends on Plan 020 ([#129](https://github.com/whynpc9/ofdrw.net/pull/129), merged).
- Use the post-020 baseline `bf492ef1ef1c4628e8d913cc5ee48f4d44799a07`. Before changing the engine, compare it with the current base, and record its actual output for the cases below.
- Follow the maintainer's [decision in #141](https://github.com/whynpc9/ofdrw.net/issues/141#issuecomment-6904414154): keep exactly the same line boundaries for successful inputs. Fix the budget, the RTL cluster lookup, and the repeated enumeration of emergency break candidates in separate commits.
- Product scope: `ParagraphLayoutEngine.cs`. Tests: `EmergencyBreakTests.cs`. Record completion in the Plan 026 row of `plan-index.md`. Keep the paragraph shaper, normal fit search, bidi, justification, DOCX/Flow callers, public defaults, and APIs unchanged.

## Why the search must keep the full range

`BreakGraphemes` uses a binary search over all grapheme boundaries in `(start, limit]`, followed by an upward scan with a consecutive-miss limit. Reshaped widths are not monotonic, so changing the binary search's initial range can change the result. The rejected `4 × MaximumEm` cutoff drops a fitting late candidate under legal negative tracking. It can also change the result even when the old cut is inside the truncated range.

Use the complete ordered range, the same midpoint rounding on the same relative indices, the same line rebuilds, and the same upward scan and stopping rules. `Estimate` is not used to omit a trial. The width comparison and grapheme-overflow source range stay unchanged.

## Steps

1. **Record old output.** Use the repository font fixtures at 4 mm. Store SHA-256 of invariant `SourceStart:SourceLength` lists for the eight cases below in a test-only commit. Add direct assertions for the negative-tracking case (`0:115,115:13`) and the synthetic nonmonotonic case (first cut at 20, even though cuts 11–19 overflow). Confirm the `ffi` and Arabic lam-alef regressions still pass.

   | Case | Input | Face | Width | Options |
   |---|---|---|---:|---|
   | Latin | `a` × 600 | Noto Sans | 30 mm | |
   | Mixed widths | `WiMl` × 150 | Noto Sans | 30 mm | |
   | Ligatures | `ffi` × 120 | DejaVu Sans fixture | 20 mm | |
   | Arabic | `لا` × 120 | DejaVu Sans fixture | 20 mm | |
   | Spaced | `a` × 400 | Noto Sans | 30 mm | tracking 0.5 mm |
   | Indented | `a` × 400 | Noto Sans | 30 mm | first-line indent 10 mm |
   | Thai | `ภาษาไทย` × 60 | Noto Sans Thai | 30 mm | |
   | After words | `word word ` + `a` × 300 + ` end` | Noto Sans | 30 mm | |

2. **Limit accepted lines when collecting drafts.** In `Session.Run`, check `drafts.Count` before adding the next yielded line. The Nth line is admitted, the N+1st causes `BudgetExceeded` with a `MaximumLines` diagnostic, and no further paragraphs are enumerated. This applies to empty paragraphs too. Remove the redundant late check in `Paginate`. Retain a separate `long` trial-build counter, capped at `(MaximumLines + 1) × (MaximumNearMisses + 3 × MaximumFitAttempts)` to allow all trials for the one pending line. Keep cancellation and `MaximumShapeCalls` independent; the trial-build diagnostic must identify trial work.

3. **Index RTL successor glyphs.** Only a tracked RTL run needs the lookup. In existing `clusters` traversal order, assign each in-range code point once. For overlapping ranges, the first match still wins; duplicate starts still use the existing final `LastGlyph`. Uncovered successors and successors outside this run still use the glyph of the current cluster. Test the affected glyph advance directly using an overlapping result returned by a custom `ITextShaper`. Existing HarfBuzz RTL, ligatures, and LTR tracking must keep their placement.

4. **Reuse emergency boundaries within a paragraph.** On the first emergency fit for a paragraph, build its complete ordered list of grapheme breaks once. Use the slice `(start, limit]`. Its first and last indices can be found by binary search; the fitting search inside the slice stays relative to its first index. This removes repeated scans and allocations of the remaining boundary list. It still builds long trial lines, so it does not establish linear runtime for long inputs. A change to the probe order or a way to eliminate long trials is separate work and needs its own compatibility proof.

## Acceptance

- Keep the baseline hashes and the direct compatibility assertions unchanged. Include the existing `EmergencyBreakRebuildsCutsInsideALigature` and `LetterSpacingCountsEveryGraphemeOfALigatureAndWidensRtlGaps` in verification.
- Five LF with `MaximumLines=5` must yield five lines; 1,000 LF must hit `MaximumLines` at the sixth pending line. Also test emergency text with a small `MaximumLines`, and a separate low `MaximumShapeCalls` request.
- With Noto Sans Regular, 4 mm and an 80 mm frame, report lines, `ShapeCalls`, and median elapsed time over multiple runs for 16k, 32k, and 90k repetitions of `a` in the same Release environment, for baseline and candidate. Record errors where a result is unavailable. The 90k candidate must cover `[0,90000)` without gaps or overlaps at the public defaults. Timing is evidence, not a test threshold. If this instead hits `MaximumShapeCalls`, investigate and report the actual cause; keep defaults and search order unchanged.
- Run the Release solution build, all solution tests (including the Layout project), the DOCX Native paragraph test filter, and local packaged consumer E2E. Use the task `DOTNET_CLI_HOME` and flags from `AGENTS.md`. Materialize fixtures with `git lfs pull --include="e2e/**,tests/**" --exclude=""` and run `python3 scripts/check-repository-storage.py --require-materialized e2e --require-materialized tests` after staging.
- Generate F08 using `OFDRW_PARAGRAPH_ARTIFACTS` and `EveryF08OutputFeatureRoundTripsAndExports`. Generate the deterministic DOCX fixture in explicit and default Native modes with the final paragraph engine enabled. Check OFD text/pagination and PDF text readback, render every PDF page to PNG, then inspect every page in macOS Preview. Keep input/output hashes, environment, source SHA, access paths and per-page observations in `artifacts/`; close the reviewed Preview windows. Record each chain separately: layout → OFD → PDF, DOCX → explicit Native OFD → PDF, and DOCX → default Native OFD → PDF.
- Stop and report engine drift that invalidates the baseline, any changed successful-input line choice or prior ligature failure, or a candidate that cannot admit the 90k case without changing a public default or the search order.
