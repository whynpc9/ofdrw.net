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
