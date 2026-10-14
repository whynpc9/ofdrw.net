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
