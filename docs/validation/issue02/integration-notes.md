# 2026-10-15 accepted main integration

Runtime source `f53545ea83390318384be6fbe2e7cb06dfbe80f4` merges `792d5c0` and main `ea198538f26b1552a9902dcae898ffce8ef4362d`. The original PR #11 stays based on main. Six actual conflicts were resolved locally; no password package, GitHub merge, tag, NuGet publish or dispatch is part of this change.

Astra High bounded design and final read-only audit found no deterministic merge defect. PDF keeps the image-only strict entry/budgets/cache/geometry while main owns ordinary drawing, font subset/coverage, typed/composite projection, full matrices, flow and model contracts. Shared ConvertPackageAsync performs main graph/annotation preflight; each unique nested package validates within the existing strict/legacy error boundary before font context/cache insertion. The same operation material is passed through nested forms. All package-load options, including both annotation limits, remain copied/snapshotted and inherited.

CLI takes the union of image commands, document tools and replace-text before general parsing. CI retains main branch filters and all existing gates, and all three image test classes run on Linux/Windows/macOS. E2E retains main flow/table/graphics/subset probes plus image APIs, installed CLI and import -> watermark -> split -> PNG. The default feed is still verify-package-artifacts.PACKAGES, exactly11, with optional extensions independent and Issue11 excluded.

New tests prove shared typed composite placements/holes in direct images and repeated nested signature forms, unsupported typed clip output preservation, and strict already-loaded page IDs. Main Reader now refuses empty/non-numeric/duplicate page IDs before either renderer; existing tests assert both refusals and unchanged targets for those inputs, while missing-ID legacy behavior and valid-document signature-tolerance pixel tests remain. No production coverage or failure assertion was weakened.

First remote CI found host Arial CJK dependence in seven image-fixture cases on Ubuntu/Windows. The fixture correction reads the repository-pinned OFL CJK ZIP entry with exact size and SHA-256 verification and binds it as an embedded document font. Production fallback/coverage is unchanged; 143 image tests pass locally. Existing current public source fixtures already used that exact explicit OFL input, so this fixture-only correction does not invalidate the runtime-source/hash-specific Preview artifacts.

Sol Low independently ran2249/2249 complete Release tests and65/65 mandatory original-font tests, zero skips, and read back the actual11 package identities/manifests. Local consumer version0.1.0-issue02.main1 retains both image objects after image import/watermark/split/export. Preview on2026-10-15 reopened17 public pages and2 local original-source diagnostic pages; all passed and all task windows/open panels were closed. The exact bounded record and archive hashes are in manifest.json and preview-acceptance.json.

# Historical dependency integration

# Issue02 / Issue03 dependency integration

Issue02 pre-integration head: `b29ed9db8d157569881d9dcd2a6d56f085f32baa`.
Reviewed dependency: PR #12 / `codex/document-tools` at `df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9` (runtime source `31569bd`).
Integration uses a history-preserving merge in the same `codex/ofd-image-io` worktree; original PR #11 is retained. No GitHub PR merge or package publication.

CLI whitelist/dispatch conflicts take the union; both command-local partial implementations remain. PDF auto-merge retains Issue02 strict signature policy and all Issue03 ordinary-element changes; Reader, Writer, Pruner and resource ownership remain Issue03 implementations. The source patch against the dependency keeps image-only strict opt-in and existing default public-PDF signature compatibility.

The new `MaxAnnotationObjectCount` / `MaxAnnotationXmlBytes` options are validated/snapshotted by OfdToImageConverter and inherited by nested appearance reads. Their counters remain cumulative within each Reader invocation, including repeated references and hidden primitives. They are not a new export-global annotation ledger. Existing outer-plus-unique-nested expanded bytes/entries/pages accounting and payload/form/font caching remain cumulative.

Cross-feature regressions cover image-entry annotation quota propagation/snapshots, cache-reused nested annotated signatures, exact sheared annotation clips, unsupported annotation whole-export failure (API/CLI target preservation), complete glyph CTM at fixed baseline anchors and upright controls, and centered image import -> watermark -> reordered split -> image export. Existing Issue02 strict metadata/cancel/entry-limit tests and all Issue03 contract tests remain. Non-finite source boxes now fail in the shared Reader before the image-specific origin check; the old diagnostic assertion is adjusted without weakening error type or preserved-output assertions.

## Public evidence font audit

The prior Issue02 baseline-native/default OFD payloads each embedded a full 23,278,008-byte Arial Unicode MS / ArialUnicodeMS font. Its font name table contains the Monotype license notice (OS/2 fsType 0x8 is not blanket public redistribution authority). The prior source.pdf and baselines also embedded subsets. They remain local diagnostic artifacts only and will not enter the new current public evidence bundle.

New public fixtures use the repository-pinned Noto Sans CJK SC Regular plus its OFL, and original MIT rectangle-glyph fixtures/generator. The Issue03 matrix generates a licensed DOCX copy and embedded Noto Native/default OFDs; its old Preview acceptance is not inherited. All acceptance uses newly generated integration files with source/dependency hashes and actual page/file paths. Existing Git history is preserved as instructed; replacement of the current evidence does not rewrite historical blobs.

CI pull-request base filters include the exact dependency branch so a stacked PR still runs the same platform/image/package/document-tool gates.
