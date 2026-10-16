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
