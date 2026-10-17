# DOCX native font identity repair — 2026-10-17

Baseline: `f90b904ee6d71029bb6d30d12c5fd7ea941ff644`. The repair is the commit containing this report; the final full commit identity and artifact hashes are recorded in the task's ignored `artifacts/font-fix/` directory. The original cross-platform acceptance evidence is unchanged.

## Behavior

- Authentic configured Noto CJK faces now embed in native/default OFD. Equivalent Noto family aliases retain authenticity; the catalog no longer manufactures a Noto alias from a different system CJK family.
- Native Latin advances use the same configured physical face as embedded text, including requested bold/italic selection. A separate bug that embedded the regular face while measuring a real bold face is corrected. OFD `FamilyName` carries the family without the internal style suffix.
- Embedded OFD resources already take the strict embedded-font path in `DocumentFontContext`; these outputs no longer reach `CjkViewerFontLoader` or its host stand-ins. PDF font parsing/export code is unchanged.
- Intentional system-name CJK compatibility (SimSun/SimHei/etc. without embedding, including its legacy Latin metric policy) is retained. Missing/unconfigured fonts remain on the existing name-only/host-resolution path; this repair does not establish deterministic rendering for absent fonts. Supply the actual font through `FontDirectories` for the verified profile.

## Local evidence

Environment: macOS arm64, .NET 10, isolated managed worktree `xplat-font-identity/ofdrw.net`. Commands use task-local `DOTNET_CLI_HOME`, single-node builds, disabled build servers/shared compilation. No merge, tag, publication, golden change or tolerance change.

- Fixed Noto source: SHA-256 `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa`; reused the existing repository OFL fixture, no new binary fixture.
- Regression coverage: native/default, subset/full embedding, original A/space advances (608/224 units per 1000 em), original Latin/CJK glyph bytes through OFD and PDF (allowing only zero alignment padding), a resolver that rejects host queries, distinct physical bold-face embedding, and system-name compatibility.
- DOCX affected tests: 147 passed, zero skipped (`docx-complete.trx`). PDF identity/isolation/style selection: 85 passed, zero skipped (`pdf-verified.trx`), followed by 3/3 strengthened glyph-byte identity checks (`identity-glyphs.trx`). These are affected projects/filters, not a full-repository test claim.
- Local eleven-package consumer: PASS at unpublished version `0.1.0-xplat-font-fix.1`; existing macOS visual corpus all 12 full-page/region rows RMSE 0 under unchanged 0.035 gates.
- Generated original/Noto fixtures × explicit native/default: 4 cases, 8 pages. Exact OFD text including spaces matches DOCX. Whitespace-normalized PDF extraction matches once with no duplicate source text. All four canonical OFD structures match.
- All three Noto OFD resources reference the same embedded subset (`9f04c8dfad3b6068a42031aa0691c4db2ec9039182453b22a9168d069d0019e5`); sample OFD grows from about 4.3 KB name-only to about 291 KB with the real subset. Final subtitle Boundary: `73.919 40.005 62.171 4.815`.
- Automatic rendering: all 8 PDF pages rendered at 144 DPI. PNG visual review: both distinct page images reviewed; all four cases have identical per-page PNG hashes.
- macOS Preview: all 4 final PDFs freshly opened, pages 1–2 individually inspected/captured, document windows closed. Chinese/English, proportional spacing, local bold/italic/color, table fill/borders, alignment and explicit pagination passed for this fixture. No claims for unrepresented images, headers/footers or automatic page numbering. See `visual-acceptance.json`, `preview/`, `final-macos/generation.json` and `artifact-hashes.json` in task artifacts.

## Affected retest matrix and remaining gates

| Scope | Retest on the immutable repair commit | Acceptance |
| --- | --- | --- |
| macOS / Windows / WSL | Same two DOCX fixtures × native/default, same pinned Noto bytes | Exact source text; original font identity in OFD/PDF; no host-dependent Latin Boundary/DeltaX; native/default canonical equality |
| PDF and direct OFD readers | Existing Preview/Suwell workflows, each generated document and both pages | Chinese, Latin spacing/overlap, local styles, table, pagination; font substitution must not be reported as original identity |
| Targeted tests | `FontIdentityRegressionTests`, `DocxFontIdentityTests`, native/generated style tests, font isolation/style tests | Real trusted-font tests enabled with `OFDRW_REQUIRE_TRUSTED_FONTS=1`; zero skips for the selected gate |
| Package consumer | Same candidate package bytes on each target, existing generated DOCX and visual corpus checks | Preserve existing approved golden/thresholds |

Windows reviewed golden remains **NOT COMPLETED**: the existing consumer deliberately throws `PlatformNotSupportedException` because no reviewed Windows baseline exists. Adding a directory or promoting this repair's output to golden would not establish review. A Windows candidate and explicit independent visual review must precede a separately authorized baseline change.

The two Audit023 symbolic-link tests remain a Windows **environment prerequisite**, not a product fix: use an isolated test runner already granted symbolic-link creation capability, then rerun the two unchanged tests. This repair neither changes their assertions nor changes Windows security/developer settings. They passed in the affected macOS test run. If target privileges are unavailable, keep the target gate blocked; do not silently skip it.

Cross-platform and Suwell acceptance of this repair are **NOT COMPLETED in this task** and remain assigned to the source conversation. This report is local repair evidence, not a release or overall cross-platform pass.
