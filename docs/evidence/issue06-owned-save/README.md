# Issue06 keyword location and replacement evidence

The implementation and bounded local acceptance are complete. PR review and CI closure are recorded separately in the PR; this document does not grant merge or release authority.

The actual preserved-history merge is `0ff9db5eb6a9baaeaa57497534fabf07f0e08848`, with parents `0e2452473f47c8aa0f916423be6362276ad15666` (Issue06) and `f7e3a62b6e63522a836643e59ff8144162ab5bfa` (accepted shared foundation). Its 306 source, script and fixture files are byte-identical to the independently accepted generated pair. The only shared integration resolution is the reviewed Reader friend union. The actual merge also passed a forced Release CLI build. Later documentation/evidence commits retain this runtime baseline; no claim that generated-pair tests originally ran on a merged Git HEAD.

The stacked PR targets `codex/typed-object-owned-save-foundation` at that exact foundation commit. The complete Issue06 increment relative to that base comprised 77 files before these documentation/evidence and CI-trigger additions, rather than only the final 22-file correction checkpoint.

## Recorded acceptance

| Gate | Actual result and scope |
| --- | --- |
| Functional regression | Independent isolated Release tests: 1245/1245, zero failures/skips: Core 67, Packaging 932, DOCX 63, PDF 144, CLI 35, Signatures 4 |
| Original fonts | 65/65 real-font tests: trusted geometry 33, keyword locator 26, replacement 6; unchanged full original Noto Latin/CJK font payloads |
| Build and packages | Reader/Layout/PDF netstandard2.0 and 2.1; 11 actual local packages at `0.1.0-issue06-local0e24524-f7e3a62`; installed CLI and package-only consumer E2E |
| CLI save | Six genuine runs: Latin/CJK × Changed, identical-text Noop, no-hit Noop; all exit 0, inputs preserved, no temporary residue; four Noop outputs byte-exact |
| Public geometry | Package-only Reader/Locator verified actual uninstrumented Changed outputs: original font SHA/length/resource, nonzero exact glyph IDs, advance, baseline and all ink-box corners. CJK dash glyph 14, X=17/18.388/19.776; Latin glyph 16, X=17/18.288/19.576; Y=31 mm |
| Failure preservation | 11 actual jobs: ordinary and typed variants of same-path, symlink-to-input, wrong trusted SHA, replacement overflow and directory target; one observed partial-stage Noop source change rejected before commit. Old target preserved and temporary files removed |
| macOS Preview | Six freshly opened PDFs, eight pages: Latin/CJK input/output (one each), explicit native and default DOCX→OFD→PDF (two each). Chinese/English, local styles, tables, alignment and pagination inspected; Preview quit afterward |

The actual CJK Changed private save used a single context and four separate default ceilings of 536,870,912 bytes: storage 86,682,639; copy 276,478,354; validation work 517,229,010; output upper bound 78,256,486. Actual/reserved I+O+A peak was 34,799,738. These are logical cumulative accounting and observed file sizes, not process memory or CPU guarantees. No original font shrinking or budget reset was used. CJK output grew by seven bytes (11,599,908→11,599,915).

Functional tests and visual acceptance are separate. Preview followed the actual native/default OFD products, never direct DOCX→PDF. All eight page PNGs are retained. The deterministic layout sample has no images, real header/footer or automatic page numbers; those are not claimed inspected. The original source, font, DLL, command, environment, output and page hashes are retained, as are the original stale CLI status and its explicit metadata correction. Typed diagnostic output has different ZIP timestamps and is not substituted for the actual CLI output.

## Public archive and licenses

Download [review-evidence.zip](review-evidence.zip), then use [member-inventory.json](member-inventory.json) for exact member paths, lengths and SHA256 values, and [font-license-map.json](font-license-map.json) for original/subset font provenance. The explicit allowlist is encoded in [build_review_archive.py](build_review_archive.py). It preserves raw logs/TRX, numeric records, source/build/package identity, public consumer code/results, failure records and visual/text readback without publishing caches, binaries or package feeds.

Archive: 146 members, 24,611,338 bytes, SHA256 `391c1301ab9db96a7cadb9c914904200eb1cc659c233bfde3b01e82c3925cc47`; [archive-identity.json](archive-identity.json) binds the inventory and font map. Every member was read back and checked against its exact source bytes before commit.

The four CLI OFDs contain only verified original Noto OFL fonts. The four corresponding PDF font payloads are derived subsets, with their own SHA/length mapped to the licensed original; they are not described as byte-identical full fonts. OFL notices and the repository MIT license are included. The source DOCX contains XML/relationships only and no embedded font binary.

The two native/default layout OFDs and their PDFs embed system PingFang SC bytes and remain local only. Their hashes, eight public bitmap page images and inspection records are retained; excluded binaries are explicitly listed. PingFang payload identity is 23,278,008 bytes, SHA256 `876af2cd4854644e7f3e7feb2f688997fdb3343c6df6693611209c9dfb47ccec`. No redistribution permission is inferred from installed system fonts. Prior commercial Arial Unicode artifacts and whole historical artifact trees are also excluded.

## Limits

Acceptance covers the listed fixtures/pages and the original Regular-font, simple-scalar, horizontal-positive, translation-only replacement profile. Unsupported geometry, overlap, missing glyphs, insufficient original space and quota exhaustion reject the batch. The late failure observation proves the recorded Noop raw-input rejection, not a late Changed/cancellation claim. Existing Issue09 multi-M behavior outside the selected profile remains outside this ticket and is not declared fixed. No GitHub merge, NuGet publication, tag or production signing acceptance is included.

Usage and signature rewriting semantics are in the [tutorial](../../tutorials/16-keyword-replacement.md).

## PR19 mandatory font gate correction

The first Codex review correctly found that the dedicated `run-trusted-font-tests.py` filter omitted `TextReplacementRealFontTests`. The full 1245-test regression above had executed those six cases, but the dedicated gate selected only geometry/locator. Commit `6053a8afb4341de5ec12f58c5d42232fbdf77c18` adds the missing class and clarifies the scope; product code, font payloads, budgets and page artifacts remain unchanged.

Independent verification used an exact Git archive of that commit and the actual mandatory runner, with fresh results and forced trusted-font environment. It passed **65/65 with zero skips/failures**: geometry 33, locator 26, replacement 6. All 377 archived files matched the commit before/after. See [per-case summary](font-gate-r1/summary.json), [actual invocation](font-gate-r1/invocation.json), [raw log](font-gate-r1/runner.log), [TRX](font-gate-r1/results/trusted-fonts.trx) and [member hashes](font-gate-r1/member-hashes.json). This targeted tool correction supplements the unchanged archive rather than relabeling its original source identity. It required no new rendering or Preview inspection.

## PR19 platform path correction

The second Codex review found that Core's file-intent comparison always ignored case, rejecting distinct `Input.ofd` / `input.ofd` on case-sensitive Unix even though CLI/Layout used platform-specific comparison. Commit `6b12e018c21f905904d1994b6c5b8f3875e1e1cf` aligns this one Core comparison with the existing Windows/macOS-insensitive, other-platform-ordinal policy. Ticket/once-only checks, canonical/reparse rejection, charges, publication and failure handling remain in place; the fixed foundation ref is unchanged. Actual Core ticket and CLI Noop/Changed, same-path and symlink regressions were added. The tutorial's Reader namespace was also corrected.

Independent exact-archive verification passed Core **69/69**, CLI **39/39**, affected Packaging classes **74/74**, all zero failures/skips, plus Core/Reader/Layout/PDF Release builds for netstandard2.0 and 2.1. All 387 archived files matched the commit before/after. See the [report](path-r2/report.md), [summary](path-r2/summary.json), [static closure](path-r2/pr19-owned-path-case-closure.md), [raw TRX/log hashes](path-r2/member-hashes.json) and commands/environment in that folder.

This local run actually exercised macOS case-policy rejection, valid-OFD same-canonical rejection and real symlink-to-input rejection, preserving input/target and leaving no temporary stage. It did not exercise Linux independent case-only publication or Windows link-API unavailability. The successful Linux Noop/Changed and fresh public Reader/Locator neighbor checks are separate branches in the actual CLI regressions and are verified by the latest Ubuntu PR CI. Windows link-API unavailability, if encountered, explicitly reports `NOT EXECUTED` and must not be interpreted as alias-rejection coverage merely because its test returned. Current-head CI/review closure is recorded on [PR19](https://github.com/whynpc9/ofdrw.net/pull/19).

This changes file admission, not page content or rendering. The archive's original source/build identities and Preview pages remain explicitly historical and unchanged; the new path gate supplements them without regenerating the archive or relabeling old tests as current-head execution.

## PR19 missing output parent correction

The third Codex review found that owned staging did not create missing output parents, unlike the existing CLI `AtomicOutput`. Commit `439cf4404c63362ceeceb56e77078fcabb966763` prepares the directory in both Noop/Changed stage constructors after scope admission and before stage creation. The private helper prepays bounded copy/work, checks the current destination, creates the directory, and rechecks context/token/destination; final source/claim/rename and cleanup remain unchanged. No new authority, budget reset, font service or foundation-ref change is involved.

Independent exact-archive verification passed CLI **44/44** and affected Packaging **74/74**, zero failures/skips, plus Core/Reader/Layout/PDF dual-TFM builds. Five actual CLI cases cover nested missing-parent Noop/Changed, file ancestor, real symlink ancestor, and invalid input that must not create the parent. The macOS symlink branch executed; Windows-only API unavailability is explicitly scoped if encountered. All 406 archived files and the diagnostic copy's original files remained identical before/after. See [report](path-r3/report.md), [summary](path-r3/summary.json), [raw TRX/source/DLL/command hashes](path-r3/member-hashes.json) and [static closure](path-r3/pr19-missing-output-parent-closure.md).

Because directory creation adds real charges, a new single full-original CJK Changed run used the current source and the existing typed observer of the actual private dispatcher. It reached **ActualCommitted**, Completed=true/Error=null, created the previously missing nested parent, preserved input/full 21,659,544-byte font, and left no temporary files. Each ceiling remained 536,870,912: storage **86,682,543**, copy **276,591,506**, validation work **517,318,480**, output upper bound **78,256,486**; actual/reserved I+O+A peak **34,799,738**. These are this run's measured logical counters, not the old archive's 517,229,010 work value or process memory.

The actual [current CJK OFD](path-r3/current-cjk-changed.ofd) is 11,599,915 bytes, SHA256 `a7a95ed2916e40daca8549b9e185d4f9017bb32f709dc57a6bc9c57a1805bfbd`. Its full embedded font SHA `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa` is mapped to the existing [Noto OFL notice](../../../tests/fixtures/trusted-fonts/licenses/Ofdrw-CI-Noto-OFL.txt) in the member manifest. [All five uncompressed entry comparisons](path-r3/zip-entry-comparison.json), including exact Content XML and font bytes, match both old uninstrumented and old typed outputs; raw ZIP hashes differ. The unchanged Content/font is explicitly mapped to the original accepted page PNG/Preview identities; no new PNG render or Preview session is claimed. The old ZIP, counters and visual records stay unchanged and historical.

The [probe](path-r3/probe/Cli.OwnedDiagnostics.cs) and [harness](path-r3/probe/harness/Program.cs) only observe the original private dispatcher after terminal completion; source/DLL bindings are retained. To rebuild that observer, extract the exact source commit into `probe/source`, copy the addon into its CLI source directory, and build the provided harness project. It uses the existing CLI.Tests friend, with no reflection, new friend or caller-supplied credentials. This supplements the prior package/font/visual evidence and does not repeat those unchanged gates.
