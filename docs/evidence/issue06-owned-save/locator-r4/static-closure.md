# PR19 R4 CanReplace ownership flag — static closure

No blocking finding in the frozen three-file diff. Final source checkpoint `73d68c88c604f502c1f923bff3e4613ea30956b0`; all three exact committed blobs independently match the hashes below. Read-only actual-source review, with no dotnet, product edits, commits or GUI.

Analysis.Flush now prepays one Field per matched fragment before the new reference comparisons, selects the source at **match.ClusterStart**, and combines the existing per-source Editable flag with ReferenceEquals against that matched source. This correctly restricts ownership eligibility to one captured occurrence. The fresh-per-occurrence capture and shared-per-run source references establish that identity; it does not rely on mutable model IDs, paths or source-free XML provenance. Multiple TextCodes within a single occurrence remain eligible. Search segmentation, matching, fragment geometry, HasXml policy, callback/final-font guards and editor authentication are unchanged.

The work is bounded by admitted match.ClusterCount: Field charges the existing finite comparison counter and 16 logical Work units per fragment before comparison; chunked outer work reservation remains unchanged. This adds no allocation, callback, whole-source validation, shared owner authority or runtime context.

Regressions retain cross-object hits and exact baseline fragments but require CanReplace false; verify a later-owner-only match remains true, so segment[0] cannot accidentally determine ownership; and retain true eligibility plus successful cross-slot replacement/readback within one object. The new two replacement-route cases are actual Writer-to-Reader source-backed fixtures: every hit fragment must be authoritative XML, occurrence paths differ, Find remains complete, and both explicit-hit and keyword Replace must throw exactly UneditableKeywordSource and return no result. They verify original logical text, raw source XML, font bytes and preserved package entries remain unchanged. The deterministic metric provider is appropriate for this ownership-flag regression; these tests do not claim production font/CLI qualification.

Before/after exact hash readback:

- `src/Ofdrw.Net.Reader/Keywords/Internal/OfdKeywordAnalysis.cs`: `24e08884d0186cf93b8bcb48feb3bc528f6310bf9324830d5a7aaa670a68d4ff`
- `tests/Ofdrw.Net.Packaging.Tests/KeywordLocatorTests.cs`: `771196d086852a8071c7591834dd7ce1bded51bc419538b4124b99b75ce4d37c`
- `tests/Ofdrw.Net.Packaging.Tests/TextReplacementTests.cs`: `97ace954bcf2c5b89e63021e5cb5894ec26007f351ee3b37f6adb5dcdeef10a6`

Root explicitly narrowed the final validation scope to affected locator/replacement tests for this small flag change. That direction supersedes the broader full-CJK/65-case recommendations in the initial proposal; those are **not outstanding mandatory gates** for R4. No concrete new defect in this patch warrants expanding that scope. The earlier full CJK439 budget evidence remains historical and is not represented as a new R4 measurement. Fonts, pages and shared ownership protocols are unchanged; existing Preview content mapping may be retained without claiming a new visual review. This static closure is eligible for the planned affected Low tests and is not an overall runtime acceptance.
