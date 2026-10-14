# Issue09 typed object acceptance

The functional and package gates passed on product candidate `60163cca80a4da62bf748ba19a300a8d714b80c4`. **Overall acceptance remains open: macOS Preview could not start because the Mac was locked, and current-head PR review/CI is pending.** PNG inspection is auxiliary and does not replace Preview. Later documentation/evidence-only commits do not change or relabel this measured product source.

The branch preserves original B `c68a6c7` and accepted shared foundation `f7e3a62` through the normal merge `45aeec5`, whose second parent is exact Issue06 `7caa8b54f2b28e0479d051e6ddf9b6d3227b5c3d`. The stacked PR targets `codex/keyword-locate-replace`; its delta is Issue09. No merge, tag or package publication is implied.

| Gate | Actual scope and result |
| --- | --- |
| Independent functional regression | Release 1320/1320, zero failures/skips: Core 70, Packaging 987, PDF 152, DOCX 63, Signatures 4, CLI 44; required full original Noto fonts enabled |
| Builds and packages | Ten libraries × netstandard2.0/netstandard2.1: 20/20; 11 fresh local packages at `0.1.0-issue09.c.60163cc`, unchanged before/after consumption |
| Public package consumer | Fresh private NuGet cache, ten library PackageReferences, no ProjectReference; installed CLI package executes successfully; ten loaded/installed DLL payloads match package bytes |
| Public API sample | Construct → write/read → public Split clone → placement and Area CTM edit → write/read; outlines/bookmark/GoTo, sharing, source independence, original fonts/image and extraction asserted |
| Static review | Astra C review findings repaired; final narrow compatibility repair has no blocker. [Exact-source report](static-review/review.md) |
| PNG auxiliary inspection | Six pages: before/after two-page sample and unchanged original V1 input newly exported with the package converter; no observed missing glyphs, diagonal bridges, triangular table gaps, annotation duplication, unexpected blank pages or overlap |
| macOS Preview | Not completed. Native UI reported locked Mac and failed automatic unlock; manual unlock requested. No PDF was opened in Preview by this run yet |

[Independent report](functional/verdict.md), [machine evidence](functional/evidence.json), [original TRX](functional/trx/), and [sample readback](sample/results.json) retain source, commands, package/DLL identity, font bytes and output sizes. The earlier `8d907a1` failed full run (1318/1320) remains local historical evidence; the two actual compatibility regressions were repaired before this successful new run.

## Reviewable products

Open [final native OFD](sample/after.ofd) or its [OFD → PDF export](sample/after.pdf). Compare [before PDF](sample/before.pdf), [after page 1](sample/after-1.png) and [after page 2](sample/after-2.png). The sample's intentional 38 mm text clip is labelled. Only the left composite placement moves (+8,+6); the image Area CTM translation is edited separately. Resource viewports clip overflow; NonZero stays solid and EvenOdd has a hole. The local annotation is drawn once. The sample has explicit footer/page labels, but no actual repeating header/footer system or automatic page-number feature is claimed tested.

The original multi-M defect is preserved in [historical PDF](historical/historical-original-v1.pdf) and [historical page 1](historical/historical-original-v1-1.png). Its original source was `37e8f0f19512ad8d02b1921831676b8fa9c1aca9`. The [exact unchanged input OFD](sample/original-v1.ofd), SHA `9c3444ca5aa6a1ce3e2d957db697cdfcb3afd61fc126ff76823260bef4114c2f`, is now exported to [repaired PDF](sample/original-v1.pdf). Its literal remains `M 0 0 L 170 0 L 170 30 L 0 30 C M 0 10 L 170 10 M 85 0 L 85 30`; no independent-path workaround or vendor upgrade was used. `AddMove` emits each transformed actual move within one compound fill/stroke.

Sample native OFD changes 11,892,889 → 11,892,891 bytes (+2); PDF 155,008 → 155,041 (+33). Original V1 OFD is 11,892,467 bytes; its new PDF is 118,716. Full original Noto fonts intentionally remain embedded; no shrinking or new font service was added. SVGs remain local because each embeds roughly 29.6 MB of font data; their hashes and generated-byte sizes are in the readback record. No SVG browser visual verdict is claimed.

## Reproduce and limits

The public generator is [Program.cs](../../../e2e/Ofdrw.Net.TypedObjects.E2E/Program.cs); its project switches to PackageReference only when `OfdrwPackageVersion` is set. Use the checked-in trusted-font bundle via `scripts/prepare-trusted-fonts.py`, the [checkerboard](sample/checkerboard.png), and the preserved original V1 input. Build Release with single-node/no-build-server flags and invoke the built DLL directly with output-directory, font-directory, checkerboard and V1 input arguments. The consumer source and package identities are bound in the functional evidence; candidate packages are local verification artifacts, not published NuGet versions.

Supported: constructible outlines/bookmarks/GoTo; local annotation appearances; one filled Path Area per Clip, compound subpaths and NonZero/EvenOdd, successive Clip intersection; shared/nested acyclic Text/Image/Path composites with explicit CTM and viewport clipping. Multiple Areas persist but general union export is refused. Unknown XML preserves its envelope/position; removal or replacement that would discard opaque owners refuses. Complex Text clips, DrawParam, group opacity, thumbnail/substitution and opaque drawings preserve and fail closed. Strict keyword editing rejects typed clips/composites; Merge/Mix rejects this graph scope. These results cover the stated synthetic samples and pages, not arbitrary complex Word or OFD fidelity.

## Licenses and identity

Both public OFDs contain only full Noto OFL originals, recorded in [font payloads](font-payloads.json). PDFs contain derived subsets. [Latin license](licenses/NotoSans-LICENSE), [CJK OFL](licenses/Ofdrw-CI-Noto-OFL.txt) and [repository MIT](licenses/MIT-LICENSE) accompany the artifacts. No system PingFang/Arial binary, package feed, cache or DLL is redistributed. [Member manifest](member-manifest.json) binds the allowlisted raw records and products by exact length/hash; hashes establish identity, not visual acceptance.
