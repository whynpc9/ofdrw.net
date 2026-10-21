# Local archival preflight

`Ofdrw.Net.Packaging.Preflight.ArchivalPreflightChecker` inspects an existing OFD ZIP using six finite local rules. Its profile is `ofdrw-local-preflight/0.1-draft`, rule revision 1. It is not OFD-H, OFD, GB/T 42133 or archival certification. Standard mappings are unverified and every `Clause` is empty.

```csharp
using Ofdrw.Net.Packaging.Preflight;

// Caller supplies a stable snapshot, at position zero when seekable.
using var input = File.OpenRead("candidate.ofd");
var report = await new ArchivalPreflightChecker().CheckAsync(
    input, new ArchivalPreflightOptions(), cancellationToken);
var fontPresence = report.GetRuleOutcome("APF-003");
// Inspect RunState, RuleResults and CoverageGaps before applying your own policy.
```

The checker leaves the input open, restores seekable input to zero on success, rejection, cancellation or failure, and never writes to it. A non-seekable stream is consumed from its current position into bounded memory; no temporary file is used. The caller must prevent concurrent input mutation. Position restoration failures and ordinary I/O failures propagate; argument errors and cancellation also propagate. Malformed/unsafe ZIP input or an exhausted budget produces `RunState=Incomplete` and Unknown results for every unfinished rule, retaining earlier observations. This API does not accept caller-owned mutable package archives.

The loader's existing ZIP limits are copied from `LoadOptions`. Additional defaults are 32 MiB cumulative distinct XML bytes, 8 MiB per XML, reader depth 64, 200,000 XML nodes, 100,000 reference visits, and 10,000 results. At least twelve result slots are required, six reserved for interruption diagnostics. XML bytes/depth/nodes are checked before tree construction. Repeated references consume reference/result budgets; distinct parsed XML is cached. These are local operational limits, not standard limits or a guarantee about peak process memory. Loader central-directory allocation limits remain those of the existing public loader.

| Rule | Applicable local condition | Important limits |
| --- | --- | --- |
| APF-001 | Exactly one recognized DocBody, one DocRoot, readable Document, at least one existing Page target | `RequireSingleDocument=false` makes this rule NotApplicable; other rules still traverse all recognized documents. Page drawing content is not validated. |
| APF-002 | Recognized resource/index declarations, page/template BaseLoc, resource FontFile/MediaFile resolve inside the package and exist | Only known edges are inspected. Existing adjacent PageRes.xml is included; an absent convention-based PageRes.xml is not an error. |
| APF-003 | Each recognized Font declaration has one nonempty package FontFile; visible known TextObject font IDs are resolvable in the inspected resource scope | Nonempty bytes may be an invalid font. Format, glyph coverage, subset correctness, embedding rights, licensing and appearance are not checked. Duplicate IDs or unmodeled resource scopes remain Unknown. |
| APF-004 | Known root Encryptions.xml/PasswordCrypto markers are absent and inspected controls reveal no unsupported encryption scope | Root Encryptions.xml fails this rule without parsing/decrypting it; orphan PasswordCrypto and unreadable/unknown controls remain Unknown. `RequireNoEncryption=false` makes the rule NotApplicable while retaining marker evidence. |
| APF-005 | Declared attachment index and each unambiguous Attachment/FileLoc exist inside the package | Empty payload is permitted; an empty declared index remains Unknown. A lone FileLoc attribute is supported; attribute/element conflicts are Unknown. External=true/1 fails. Orphan Attachments.xml is Unknown. |
| APF-006 | Declared signature list, Signature BaseLoc, SignedValue and listed Reference FileRef targets can be inventoried | Empty declarations/lists or missing targets fail; malformed/unsupported signature structures or transforms are Unknown. No hashes, signature values, identity, certificate chain, timestamp or legal validity are verified. Fail severity is Warning and still makes the summary Fail. |

Only `http://www.ofdspec.org` and `http://www.ofdspec.org/2016` are recognized. Child fields must have the owning namespace and expected parent. Unknown control shapes, extensions, annotation/composite scopes and orphan conventional resources are reported conservatively as coverage gaps. Arbitrary unknown XML may contain other references: this check is not a complete external-link audit, schema validator or resource-ID closure proof.

DocRoot and DocBody/Signatures resolve from package root; page/template, attachment-index, signature BaseLoc and SignedValue resolve from their containing XML. Resource BaseLoc is resolved relative to its resource table, and FontFile/MediaFile relative to that base. A leading `/` overrides to package root. Signature Reference/@FileRef resolves from package root. Package entry lookup follows the existing loader's case-insensitive semantics; case-insensitive duplicate ZIP names are rejected. These local path policies have not been certified against a standard.

URI schemes, UNC and double-slash paths are classified before package resolution and never fetched. Query, fragment, percent escapes, control characters and ambiguous backslash references remain Unknown; root escape fails. No network client, host-font discovery, external process, writer, password or signature verifier is invoked. XML prohibits DTD and has a null resolver. External/ambiguous location metadata is redacted rather than exposing credentials or query values; known supported namespace selectors retain expanded names. Unknown namespace strings are replaced with `<unsupported-namespace>` to avoid copying untrusted URI credentials. Payload text, font bytes, attachment bytes and signature bytes are never copied into report evidence.

`RuleResults` preserves observations for individual objects, with report-local result and evidence IDs, source entry, selector, stable reason, observed/expected metadata, severity, recommendation and local source. `GetRuleOutcome` aggregates Fail, then Unknown, then Pass; no applicable observations means NotApplicable. `SummaryOutcome` uses the same order and also remains Unknown for an incomplete run. All NotApplicable stays NotApplicable. A definite failure coexists with Unknown observations; a zero-object result never proves success. `RunState=Complete` means inspection finished, even when unsupported semantics yield Unknown. Always retain the complete report alongside any application decision.

Reports use public read-only properties and can be serialized by the caller's serializer (for example System.Text.Json with JsonStringEnumConverter). The library introduces no serialization or network dependency and no CLI. The existing `ValidationReport` and permissive `OfdPackageStructureChecker` retain their original contracts; there is no implicit lossy boolean adapter.

Executable boundary examples are in `ArchivalPreflightTests`. Its four-byte font and SignedValue sentinels intentionally prove presence only and must never be used as rendering or signature-validity evidence. No pages are generated or modified by preflight, so page rendering, PNG inspection and Preview acceptance do not apply to this implementation.

## Opt-in archival rule subset and independent copy

`ArchivalRuleService.CheckAsync` adds two separate, standard-mapped observations while retaining the six APF results unchanged in `LocalPreflight`. The mapping is to **GB/T 42133-2022 section 6.2.2(a)** (remove `Document.xml/Permissions`) and **6.2.2(b)** (remove `Document.xml/VPreferences`). The edition and clauses were checked against the [standard PDF hosted by Jianghan University Library](https://tsgdag.jhun.edu.cn/_upload/article/files/93/cc/4d8bf26643f49ee8a76ee25d4ef6/82cca584-f57d-47db-9363-a1497f03326a.pdf), SHA-256 `7d76c086e1384ff200c0df8f5a40bf67660e841340c5e2630e582c81c22f1b82`, PDF page 8 (printed page 4). `VerifiedSubset` describes only this mapping. It is not a claim of complete GB/T 42133 conformance or certification; the report's `Certification` remains `NotAssessed`. Other clauses have not been implemented or inferred from the number of rules. GB/T 33190-2016 allows Permissions in general OFD, so its document-root clause does not justify this archival removal.

The new check requires a seekable, stable input at position zero and leaves it open and rewound. If the local inspection is incomplete, the document is ambiguous, or a foreign namespace affects the scope, the added rules return `Unknown`. A recognized declaration is `Fail`; its absence is `Pass` for this narrow rule. The caller must inspect `LocalPreflight` and the added results separately. An archival-subset Pass does not establish that images, fonts, actions, signatures, attachments or every resource reference are archivally conforming.

`ArchivalCopyConverter.ConvertAsync(inputPath, outputPath, options)` is an explicit, separate operation. It accepts only a clean single-document local preflight, rejects signatures and suspected orphan signature payloads with `SignedInputRequiresResigning`, and refuses unknown fields inside either target declaration. The internal handlers remove only supported `Permissions` and `VPreferences` nodes from `Document.xml`. R16 typed objects guide recognized shapes, but the copy handler does not rewrite pages or other package payloads through the general writer. Signed input is not modified, silently unsigned or represented as valid after editing. No-op input produces no output.

The converter snapshots the bounded source file, writes a temporary ZIP in the destination directory, reruns local and subset checks, and compares the entry set and every non-target entry's bytes before publishing. It also emits `<output>.audit.xml` with source/output SHA-256, changed and unchanged entry hashes, removed declarations, clauses, information-loss notice and signature disposition. Failure removes temporary files. The output and audit paths must not already exist. The input remains untouched. The checks prove that this operation preserves the existing page, resource and attachment payload edges byte for byte; they do not establish that the original input had complete semantic resource closure beyond the finite APF scope. The handler fails rather than hiding that boundary.

```csharp
var options = new ArchivalPreflightOptions();
using (var input = File.OpenRead("source.ofd"))
{
    var observations = await new ArchivalRuleService().CheckAsync(input, options);
    // Inspect observations.LocalPreflight, observations.RuleResults and Scope.
}
var copy = await new ArchivalCopyConverter().ConvertAsync("source.ofd", "archive-copy.ofd", options);
// Retain copy.OutputPath and copy.AuditPath together.
```

## Versioned coverage inventory

The [GB/T 42133-2022 clause ledger](archival-standard-coverage.md) tracks sections 5-7 and normative Annex A, including clauses that remain open or need external evidence. `ArchivalCoverageService.CheckAsync` adds three opt-in, read-only observations: `AC-001` maps to 6.2.1(a), `DocType=OFD-A`; `AC-002` maps to 6.2.1(c), exactly one `DocBody`; `AC-003` maps to 6.2.2(c), document actions limited to Goto. A document without actions gives `NotApplicable` for AC-003; unsupported action shapes give `Unknown`. It retains the unchanged R17 report rather than expanding its copy handler or changing the six local APF rules.

```csharp
using var file = File.OpenRead("candidate.ofd");
var coverage = await new ArchivalCoverageService().CheckAsync(file, new ArchivalPreflightOptions());
// Inspect coverage.RuleResults and coverage.FoundationalRules separately.
// coverage.StandardOutcome cannot be Pass while clause coverage is incomplete.
```

`CoverageStatus` is always `Incomplete` and `Certification` is `NotAssessed`. A definite violation is `StandardOutcome=Fail`; otherwise the full-standard outcome is `Unknown`, even when all mapped conditions pass. Input rejection, unsupported root namespaces and foreign `DocBody` or `DocType` forms beside a recognized valid declaration are Unknown. Missing or explicitly wrong recognized mandatory declarations are Fail even when a foreign same-name declaration is present. The PDF source and SHA are pinned in each rule definition and in the ledger; no standard PDF bytes are shipped with the package. R17's audited copy can be checked before and after with this service, but that copy only removes its two documented declaration types.
