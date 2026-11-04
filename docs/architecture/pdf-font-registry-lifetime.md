# PDF font registry lifetime (spike 014)

Status: investigation complete; recommendation C (no eviction, better signals, process recycling). Recorded 2026-11-04 on branch `t3/implement-advisor-plans` (base `d34d0a3`), Linux x64, .NET 10.0.401, PdfSharpCore 1.3.67. No source code was changed for this spike.

OFD→PDF registers every embedded font face in a process-wide `IFontResolver` (`OfdEmbeddedFontResolver`) that never evicts. The budgets are `PdfFontRegistry.MaximumRegisteredFontBytes` (256 MiB) and `MaximumRegisteredFaces` (4,096). When either is reached, every later conversion that needs a new face fails with `InvalidOperationException` until the process restarts (`conversion-contracts.md`, "达到限额会明确失败"). This note answers the four spike questions.

## 1. Exhaustion curve (measured)

A throwaway console app (not committed) converted one-page OFDs in a single process. Each document embedded **one distinct** subset face built with the SDK's own `TrueTypeSubset` from the pinned CI fonts. Only the subset glyph set varied. Registry counters were read through `PdfFontRegistry` plus reflection on the private dictionaries.

| Run | Subset face size | Failed at document | Registered faces | Registered bytes | Managed heap after GC | Working set |
|---|---|---|---|---|---|---|
| CJK (Noto Sans CJK SC, 40 Han characters) | ≈1.03 MB | 261 | 260 | 267,658,732 | 2,213 MiB | 3,219 MiB |
| Latin (Noto Sans, 5–19 letters) | ≈167 KB | 1,603 | 1,602 | 268,330,148 | 1,109 MiB | 2,109 MiB |

Intermediate CJK samples: 101 documents → 104 MB registered, 906 MiB managed; 201 → 207 MB, 1,728 MiB. Intermediate Latin samples: 501 documents → 84 MB registered, 348 MiB managed; 1,001 → 168 MB, 694 MiB.

Findings:

- Both runs fail on the **byte** budget long before the 4,096-face budget. The face budget only matters for subsets smaller than 64 KiB (256 MiB / 4,096), which these producers do not emit.
- Each document adds one face; identical bytes would be stored once, but per-document subsets never repeat.
- Retained managed memory grows about **8.5 MB per CJK document** and **0.7 MB per Latin document**. That is 4–8× the bytes the registry counts, because PdfSharpCore keeps its own parsed copies (below). The 256 MiB budget therefore understates the real retained memory: a host near the budget holds 2–3 GiB.
- Not measured: documents with several faces each, nested seal faces, and conversion throughput under concurrency.

## 2. PdfSharpCore 1.3.67 cache inventory

Source: `https://github.com/ststeiger/PdfSharpCore`, commit `d6ac8b092129a4f797365bbbf3eea2723d6c3ebd`, the commit recorded in the 1.3.67 nuspec.

| Class | Field | Key | Retains | Source |
|---|---|---|---|---|
| `Fonts.FontFactory` | `FontResolverInfosByName` (static) | face/typeface name | resolver info | `PdfSharpCore/Fonts/FontFactory.cs:375` |
| `Fonts.FontFactory` | `FontSourcesByName` (static) | face name | `XFontSource` holding the font bytes (`Drawing/XFontSource.cs:152`) | `PdfSharpCore/Fonts/FontFactory.cs:386` |
| `Fonts.FontFactory` | `FontSourcesByKey` (static) | font checksum key | `XFontSource` | `PdfSharpCore/Fonts/FontFactory.cs:396` |
| `Fonts.FontDescriptorCache` | `_cache` (singleton) | descriptor key | parsed `OpenTypeDescriptor` | `PdfSharpCore/Fonts/FontDescriptorCache.cs:172` |
| `Fonts.OpenType.OpenTypeFontfaceCache` | `_fontfaceCache`, `_fontfacesByCheckSum` (singleton) | name / checksum | parsed `OpenTypeFontface` | `PdfSharpCore/Fonts.OpenType/OpenTypeFontfaceCache.cs:140,145` |
| `Drawing.FontFamilyCache` | `_familiesByName` (singleton) | family name | `FontFamilyInternal` | `PdfSharpCore/Drawing/FontFamilyCache.cs:131` |

Measured: `FontSourcesByName` grew by 2 entries per document and `FontDescriptorCache` by 1. `OpenTypeFontfaceCache` stayed at 0 on this path.

None of these caches has a remove, clear or reset API, public or internal: no `Remove(`/`Clear()` calls exist under `Fonts/`, `Fonts.OpenType/` or `FontFamilyCache.cs`. `GlobalFontSettings.FontResolver` cannot be replaced after first use (`GlobalFontSettings.cs:82`). Clearing would require reflection on private static fields of a third-party library. The caches are guarded by PdfSharpCore's internal locks, which callers cannot take, so clearing while another thread draws could not be made safe.

## 3. Options and verdicts

- **A. Per-conversion reference counting and release.** Releasing a face from `OfdEmbeddedFontResolver` frees nothing: `FontSourcesByName` and `FontDescriptorCache` still hold the bytes and parsed descriptors under the face name. Freeing memory would also require clearing PdfSharpCore's private caches by reflection. If PdfSharpCore later asks `GetFont` for a released name, the result is a resolver miss or wrong glyphs. **Verdict: not feasible with supported APIs.** Not prototyped, because there is no supported API to prototype against.
- **B. Generation reset when idle.** The same blocker: a reset would have to clear six private PdfSharpCore caches by reflection under a global lock that PdfSharpCore's own code does not honour. It would break with any PdfSharpCore update. **Verdict: not recommended.** Not prototyped.
- **C. No eviction, better signals.**
  - Keep the fail-closed budget.
  - Report a `PDF_FONT_REGISTRY_PRESSURE` warning through `OfdToPdfConversionResult.Diagnostics` (plan 005) once either budget passes 80%.
  - Expose the registered face count next to `PdfFontRegistry.RegisteredFontBytes`.
  - Document process recycling as the supported mitigation for long-running services.

  **Verdict: recommended.** It needs no third-party internals and gives hosts an early, testable signal.

The real fix is a PDF backend whose font caches are scoped to a document. That belongs to the deferred backend decision (D10).

## 4. Recommendation

Adopt C as a follow-up plan:

1. A `PDF_FONT_REGISTRY_PRESSURE` warning at 80% of either budget.
2. A public `PdfFontRegistry.RegisteredFaces` count.
3. A host guidance section.

Until then, services should recycle conversion workers well before the budget. A practical rule is to restart a worker once `PdfFontRegistry.RegisteredFontBytes` passes half of `MaximumRegisteredFontBytes`, or after a few hundred documents with distinct embedded CJK subsets. Size worker memory for several GiB per worker, not for 256 MiB.

## Open questions

- Multi-face documents and nested seal OFDs were not measured. Each extra distinct face consumes budget at the same rate.
- Whether raising `MaximumRegisteredFontBytes` is safe depends on host memory, given the 4–8× retained-memory factor. It was not measured under concurrency.
- Side observation, not product-facing: `TrueTypeSubset.Create` appears to alter the `OpenTypeFace` it receives, because a second subset cut from the same instance lacked requested glyphs. The SDK's callers parse a fresh face each time. It is worth an internal check before the helper is reused.
