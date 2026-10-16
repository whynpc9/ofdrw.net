# Plain text OFD implementation evidence

2026-10-16 local acceptance for Plan 028. Implementation/source identity: [source.json](source.json). Complete scope, commands, results and boundaries: [implementation report](../../../plans/028-followup-implementation.md). Usage: [plain-text-ofd.md](../../plain-text-ofd.md).

- Functional: full suite 2868/2868, trusted fonts 65/65, zero skips; [full-test-summary.json](full-test-summary.json), TRX in full-tests/, layout-tests/, cli-tests/, trusted-font-tests/.
- Dual netstandard targets and XML API: [compiled-targets.json](compiled-targets.json), [API gate](api-docs.log).
- Installed candidate: 11 packages at `0.0.0-audit028`, [package manifest](package-manifest.json), [consumer log](package-e2e.log).
- Current UTF-8→adapter→native OFD→PDF: [notice source](visual/notice.txt), [OFD](visual/notice.ofd), [PDF](visual/notice.pdf), [expectations](visual/source-expectations.json). 2 notice + 3 control + 1 empty pages inspected.
- Fresh DOCX→native/default OFD→PDF: 2 pages each, from generated-layout.docx; [native](visual/generated-docx-native.pdf), [default](visual/generated-docx-default.pdf).
- Separate PNG and Preview visual review: all 10 pages passed in the scoped samples. [Page index](preview/pages.json) records original viewed file URLs, page numbers, AX and original JPEG hashes. Source PDFs remained byte-identical before/after Preview.
- [manifest.json](manifest.json) records evidence hashes and bytes. Raw AX/text fixtures retain significant whitespace/CRLF using scoped .gitattributes. No system font files are copied here. Pure-text OFD uses name-only SimSun/portable Flow metrics; this host's PDF export used Arial Unicode fallback. No universal font coverage or reversible source-byte extraction claim.

Raw command paths point to the original managed worktree and its artifacts/plan-028 folder. The files here are retained copies of those actual bytes; Preview was performed on the original URLs in the AX records. Full candidate nupkgs remain locally in artifacts/package-e2e/0.0.0-audit028/packages and were not published.
