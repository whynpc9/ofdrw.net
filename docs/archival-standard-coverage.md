# GB/T 42133-2022 archival coverage ledger

This ledger inventories the treatment and software provisions of **GB/T 42133-2022, 信息技术 OFD档案应用指南**. It is an engineering traceability record, not a conformance certificate. The [public standard copy](https://tsgdag.jhun.edu.cn/_upload/article/files/93/cc/4d8bf26643f49ee8a76ee25d4ef6/82cca584-f57d-47db-9363-a1497f03326a.pdf) is 14 PDF pages, SHA-256 `7d76c086e1384ff200c0df8f5a40bf67660e841340c5e2630e582c81c22f1b82`. Clause labels below were checked against the rendered pages because this PDF's extractable text is garbled. The [national-standard catalog](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=CF545613ABF25B011116F79DCCD3C45B) lists this edition as current. No PDF copy is committed here.

**Coverage terms:** `Exact` means a separately verified, executable condition with positive/negative/uncertain fixtures. `Partial` means local preflight or an existing operation supplies only some evidence for the clause. `Open` means no clause-level rule/handler has been verified. `External` means source history, human judgment, a renderer/reader, another standard, PKI or a target archive must supply evidence. These are implementation statuses, not per-file outcomes. On a particular file `Pass`, `Fail`, `Unknown` and `NotApplicable` retain their four-state meanings; an `Open` clause cannot become Pass merely because its object is absent. A rule's NotApplicable requires complete scope and proven absence of applicable objects. R17's `AR-001/002` and R20's `AC-001`–`AC-009`, `AC-011` and `AC-012` have bounded executable mappings. `AC-010`, `AC-013`, `AC-014` and `AC-015` are violation-only subsets; their absent cases remain Unknown. The six `APF-*` rules remain local policy and do **not** become standard clauses by resemblance.

Sections 1, 3 and 4 define scope, terms and abbreviations rather than per-package acceptance predicates. Section 2 incorporates the following **undated** normative references; their applicable current editions and amendments have not been pinned or evaluated for this ledger. A referenced-standard condition remains open until its exact text, version and test evidence are recorded. The annex also points to GB/T 20090 (all parts) for AVS; that format family has no interoperability evidence here.

| Section 2 reference | Needed for |
| --- | --- |
| GB/T 13000; GB 18030 | Character set and mapping, including 6.2.6(d). |
| GB/T 18894 | Record-unit organization in 5.2(b). |
| GB/T 33190 | Base OFD structure, font exemption, metadata and format semantics throughout section 6. |
| GB/T 35275; GB/T 38540 | Signature/seal format and trust boundary in 6.13. |
| DA/T 46; DA/T 54 | Metadata names in 6.2.1(b). |
| DA/T 47 | Long-term format qualities in 5.2(a). |

| Clause | Checked condition or treatment boundary | Coverage and required evidence |
| --- | --- | --- |
| 5.1(a) | Archival text/graphics scenario | External: intended use and archive policy. |
| 5.1(b) | Selected photographic records | External: collection scope and excluded specialist images. |
| 5.1(c) | Management statistics/records | External: record classification. |
| 5.2(a) | Long-term format qualities | External: DA/T 47 and independent portability/display evidence. |
| 5.2(b) | Record-unit conversion/organization | External: GB/T 18894 and source-record provenance. |
| 6.1 (first paragraph) | All display/print/interaction data and related resources inside package | Partial: APF-002/003/005 visits known edges; unknown XML, full resource closure and external semantics remain open. |
| 6.1 (second paragraph) | Optional metadata, signature, tags and attachments | Open: optionality and their own clauses need separate evaluation. |
| 6.2.1(a) | Root `DocType` identifies OFD-A | **Exact:** `AC-001`, recognized root and unqualified attribute only. |
| 6.2.1(b) | Custom metadata names and preferred separate storage | External: DA/T 46/54 mapping and GB/T 33190 allowance not yet evaluated. |
| 6.2.1(c) | No multiple document mechanism; split or merge source if needed | **Exact** for final package cardinality: `AC-002`. Split/merge provenance and appended file list remain Open. |
| 6.2.1(d) | Retain existing `Versions` | Open: requires before/after provenance, not one final package. |
| 6.2.1(e) | Remove files unrelated to root and transitively referenced content | Partial: APF visits known references. R18's opt-in physical attachment removal proves ownership and reference closure for a requested payload, not whole-package archival reachability or deletion of every unrelated file. |
| 6.2.2(a) | Remove document `Permissions` | **Exact:** R17 `AR-001`; opt-in audited copy handler for recognized XML only. |
| 6.2.2(b) | Remove document `VPreferences` | **Exact:** R17 `AR-002`; same copy boundary. |
| 6.2.2(c) | Remove document actions other than intra-document Goto | **Exact read-only type check:** `AC-003` on one recognized `Document.xml`. No removal handler; unsupported or opaque forms are Unknown. Bookmark/page-target validity and other action locations remain separately Open. |
| 6.2.2(d) | Summarize attachments not retained | Open: retention decision, digest and description require source provenance. |
| 6.2.2(e) | Remove `Extensions`, preserving any visual effect | External: source-aware appearance comparison and page transformation. A node-absence check alone cannot pass this clause. |
| 6.2.3(a) | Promote most used page area to document default | Open: count page areas, preserve appearance and verify output. |
| 6.2.3(b) | Prefer omitting redundant page area | Open: optional optimization, requires page/default comparison. |
| 6.2.3(c) | Remove non-Goto page actions | **Exact read-only type check:** `AC-004` traverses recognized page and template descriptions; unresolved page entries and opaque action forms are Unknown. Goto target validity is separate. No removal handler. |
| 6.2.3(d) | Distinguish same-named layers | **Exact read-only name check:** `AC-011` detects duplicate unqualified Layer names within each recognized page/template content; unresolved page or foreign layer names are Unknown. No rename handler or source naming provenance. |
| 6.2.3(e) | Remove PageBlock nesting or bound unavoidable nesting | Open: full page/template traversal and rendering equivalence. |
| 6.2.4(a) | Prefer original layer order and body/background organization | External: generation intent and visual comparison. |
| 6.2.4(b) | Place scan image and recognized text in appropriate layers | External: scan/OCR provenance and page inspection. |
| 6.2.4(c) | Prefer semantic content order | External: semantic reading order and accessibility review. |
| 6.2.5(a) | Remove non-Goto outline actions | **Exact read-only type check:** `AC-005` visits nested recognized `OutlineElem` actions. Opaque outline/action scope is Unknown. No removal handler. |
| 6.2.5(b) | Remove outline jumps to absent bookmark/page | **Exact read-only target check:** `AC-006` resolves recognized Goto bookmark names and page IDs against the selected document. Ambiguous inventories are Unknown. No removal handler. |
| 6.2.6(a) | Place resource types in public/document tables | Open: validate every resource table and reference scope. |
| 6.2.6(b)(1) | Embed/subset fonts referenced by filled/stroked TextObject | Partial: APF-003 only checks file presence and known font IDs; real font validity, glyph use and GB/T 33190 exemption remain Open. |
| 6.2.6(b)(2) | Include fonts used in clip text | Open: clip traversal and actual glyph dependency. |
| 6.2.6(c)(1) | Preserve glyph outline tables when subsetting | Open: inspect actual embedded OpenType tables and outline correctness. |
| 6.2.6(c)(2) | Preserve required metrics/program/name tables | Open: same font-byte evidence. |
| 6.2.6(c)(3) | Retain other tables needed for faithful drawing | External: glyph rendering and platform comparison. |
| 6.2.6(d) | Prefer GB 18030 character set and UCS mapping | External: GB 18030 / GB/T 13000 and text/font mapping tests. |
| 6.2.6(e) | Restrict raster resource formats or convert before archiving | Open: decode authenticated media payloads, not filename alone. |
| 6.2.6(f) | No custom expansion data in raster resources | Open: decode each supported image format and extensions. |
| 6.2.6(g) | Remove audio/video and retain summaries elsewhere | Open: source-aware transformation and summary evidence. |
| 6.3.1(a) | Default color space may be defined | Open: permission, not a mandatory presence rule; object absence is NotApplicable only after full scope. |
| 6.3.1(b) | Restrict color-space types | Open: traverse resource scopes and referenced color spaces. |
| 6.3.1(c) | Prefer a color profile | Open: optional guidance plus color-management evidence. |
| 6.3.2(a) | Remove clip that fully contains object boundary | Open: actual transformed clip geometry and visual equivalence. |
| 6.3.2(b) | Remove zero-area clip and hide object | Open: geometry, transforms and visibility comparison. |
| 6.3.3(a) | Bound visible text with appropriate tight rectangle | External: glyph geometry and appearance check. |
| 6.3.3(b) | Remove object actions, migrating intra-document links | **Partial:** `AC-010` fails a recognized retained object action in page, template or annotation appearance. Absence is Unknown without source action inventory and link-annotation migration/visual/target evidence. |
| 6.3.3(c) | Prefer removing stroke attributes when Stroke=false | Open: optional cleanup with style inheritance. |
| 6.3.3(d) | Prefer removing redundant composite drawing attributes | Open: inherited resource semantics and appearance. |
| 6.3.3(e) | Prefer removing image/video fill attributes | Open: object-type and rendering semantics. |
| 6.4(a) | Prefer shared color DrawParam for repeated paths | Open: optional deduplication with source equivalence. |
| 6.4(b) | Prefer shared stroke DrawParam for repeated paths | Open: same, including dash/cap/join. |
| 6.5(a) | Prefer resource table suited to image sharing scope | Open: cross-page/page-local use inventory. |
| 6.5(b) | `ImageObject Interpolate`, when used, is false | **Exact read-only attribute check:** `AC-007` scans recognized page/template and retained annotation appearance objects; true fails, false passes, complete absence of attribute is NotApplicable, opaque graphic scope is Unknown. |
| 6.6(a) | Prefer one text object per matching-font line | External: source line and font-run semantics. |
| 6.6(b) | Split text outside built-in font/character exemption | Open: GB/T 33190 exemption plus glyph-run verification. |
| 6.6(c) | Identify text size with `Size` | **Exact read-only attribute check:** `AC-008` requires a positive unqualified `Size` on recognized text objects in page/template and retained annotation appearance content. Opaque graphic scope is Unknown; actual glyph geometry remains external. |
| 6.6(d) | Prefer `HScale` for horizontal-only transform | Open: exact transform decomposition and appearance. |
| 6.6(e) | Prefer shared DrawParam for repeated text colors | Open: optional resource deduplication. |
| 6.7 | Treat video as image with video action | Open: media/action handling and rendering. |
| 6.8 | Prefer vector resource table suited to reuse scope | Open: cross-page and page-local composite inventory. |
| 6.9(a) | No external-resource action on objects/outline/page/document | **Partial:** `AC-013` fails retained, recognized Action/URI declarations with absolute HTTP(S), FTP or file targets at document, page/template, outline and recognized graphic-object locations, including retained annotation appearances. Unknown covers URI schemes or resource-backed actions not proved external, opaque locations and full action/resource closure. No removal handler. |
| 6.9(b) | Prefer keeping removed-action summaries | External: before/after action provenance. |
| 6.10.1(a)(1) | Freeze appearance-affecting annotation parameters | External: original/final appearance comparison. |
| 6.10.1(a)(2) | Remove annotation appearance PageBlock nesting | **Exact read-only structure check:** `AC-012` follows retained annotation appearances and fails nested recognized PageBlocks; opaque appearances are Unknown. No flattening handler. |
| 6.10.1(a)(3) | Set retained annotation ReadOnly/NoZoom/NoRotate | **Exact read-only attribute check:** `AC-009` follows the recognized annotation index and requires all three true on each retained `Annot`; unresolved or foreign scope is Unknown. Other annotation-treatment subclauses remain open. |
| 6.10.1(b) | Alternative: flatten annotation appearance into page | External: source ownership, visual equivalence and information-loss record. |
| 6.10.2 | Store post-archive management annotations separately | External: archive lifecycle and separate-file evidence. |
| 6.11 | Prefer retaining tags and closing their referenced content | Partial: known package edges only; full tag content closure and provenance Open. |
| 6.12(a) | Remove extension information | **Partial:** `AC-014` fails a recognized `Document/Extensions` declaration with a target; empty or foreign declarations and other extension locations remain Unknown. No removal handler. |
| 6.12(b) | Freeze extension appearance/print effects | External: source-specific interpretation and visual/print comparison. |
| 6.13.1(a) | Signature format and qualified CA for non-de-technicalized signature | External: GB/T 35275, CA qualification, chain/revocation and trusted time. APF-006 is inventory only. |
| 6.13.1(b) | Seal format, registration and qualified certificates | External: GB/T 38540, registration and qualified CA evidence. |
| 6.13.1(c) | Archive authority accepts lasting certificate status services | External: archive authority decision and service continuity. |
| 6.13.2(a)(1) | Record signer/time/digest in retained annotation | External: authenticated source signature and new annotation provenance. |
| 6.13.2(a)(2) | Remove retained annotation's SignRef | Open: only as part of authenticated signature de-technicalization. |
| 6.13.2(b) | Flatten signature annotation if annotation not retained | External: appearance/ownership comparison. |
| 6.13.2(c) | Remove signature record and folder after de-technicalization | Open: signed-byte policy and complete reference closure. R17 copy refuses signed input. |
| 6.13.3(a)(1) | Seal appearance to annotation with signer/certificate/time/digest | External: authenticated seal and appearance evidence. |
| 6.13.3(a)(2) | Alternative seal appearance to page with Darken | External: visual check and compositing support. |
| 6.13.3(b) | Remove seal record and folder after de-technicalization | Open: only after authenticated source and closed references. |
| 6.14 | Retain multi-version information | Open: before/after bytes and references required. |
| 6.15.1(a) | Prefer retaining text-format attachments | External: retention policy and actual content type. |
| 6.15.1(b) | Prefer retaining documented formats with archived documentation | External: documentation and archive accession evidence. |
| 6.15.1(c) | Prefer retaining qualifying signed attachments | External: a/b, signature standard and trust verification. |
| 6.15.2(a) | Convert eligible removed attachments and append to main document | External: Annex A formats, conversion fidelity, order and page checks. |
| 6.15.2(b)(1) | Prefer externalizing removed audio/video with summary | External: retained digest, description and external object evidence. |
| 6.15.2(b)(2) | Prefer externalizing oversized non-viewing images with summary | External: size/use policy, digest and description. |
| 6.15.2(b)(3) | Prefer externalizing data spreadsheets with summary | External: semantic classification and information loss. |
| 6.15.2(b)(4) | Prefer externalizing animation/3D-heavy web/slides with summary | External: content inspection and preservation decision. |
| 6.16 | Long-term OFD unencrypted; decrypt before archiving | **Partial:** `AC-015` fails a retained root `Encryptions.xml` marker regardless of local APF policy. Absence remains Unknown without general encryption discovery, source decryption provenance and plaintext proof. APF-004 remains a separate local-policy check. |
| 7.1(1) | Writer supports relevant generation/conversion and archival processing | External: product capability and conversion E2E, not package-only. |
| 7.1(2) | Writer follows the prescribed treatment | External: full clause-level end-to-end evidence still missing. |
| 7.2(1) | Reader fully outputs OFD, checks format and reports untreated items | External: independent reader and GUI/log behavior. |
| 7.2(2) | Reader disables editing of OFD-A and signals it | External: target reader UI and permissions tests. |

Annex A is normative only when attachment conversion under 6.15.2(a) applies. It is not evidence that the project can convert every listed format. Each row needs a redistributable source fixture, target-format inspection, content/ordering/visual checks and a recorded decision where the table calls for judgment.

| Annex A.1 row | Attachment class | Coverage |
| --- | --- | --- |
| 1 | Raster image | Open; image-to-page conversion and judgment for starred class. |
| 2 | Text | Open; encoding and content fidelity. |
| 3 | Word processing | Open; DOC/DOCX/WPS/RTF/UOF/UOT vary; existing DOCX converter is not complete annex coverage. |
| 4 | Spreadsheet | Open; sheets, formulas, print area and starred judgment. |
| 5 | Presentation | Open; notes/animation and starred judgment. |
| 6 | Fixed-layout document | Open; OFD/PDF/XPS/PS/CEB/SEP/AIP interoperability. |
| 7 | Web page | Open; resource capture and starred judgment. |
| 8 | Audio | Open; AVS conversion and content proof. |
| 9 | Video | Open; AVS conversion and content proof. |
| 10 | CAD | Open; whether view/output is the main purpose. |
| 11 | Database | Open; form or image treatment by content. |
| 12 | Email | Open; web/form treatment and headers/attachments. |
| 13 | Social media | Open; text/image/spreadsheet treatment by content. |

## Result and transformation boundaries

`ArchivalCoverageService.CheckAsync` reads a stable seekable package and returns R17's report plus `AC-001`–`AC-015`. It never writes or changes the default six-rule local preflight. Known mandatory declaration, non-Goto action, missing outline target, interpolation, text-size, annotation-flag, retained object-action, external URI action, extension-declaration or encryption-marker violations are `Fail`; rejected or unsupported scope is `Unknown`. Action type checks do not establish valid targets outside the separately checked outline target scope. `AC-010`, `AC-013`, `AC-014` and `AC-015` remain Unknown without complete source, resource, extension or plaintext evidence when no known violation is found. Its `StandardOutcome` can be `Fail` or `Unknown`, **never Pass**, because the ledger still has open clauses. A missing mandatory `DocType` or document is Fail. This report does not convert, render, decrypt or validate certificates.

R17's `ArchivalCopyConverter` independently copies only unsigned, locally clean, supported `Permissions`/`VPreferences` inputs. It records changed/unchanged entry hashes, rejects uncertain fields, rechecks the output and preserves all non-target payloads byte for byte. Those checks prove the limited edit did not change known page/resource/attachment payload bytes; they do not establish full reachability, visual correctness, complete standard coverage or PKI validity. `AC-001/002` can be run on both original and copy without changing that copy contract. Signed inputs remain rejected rather than silently invalidated.

R18's `OfdDocumentPackage.RemoveAttachment(..., Physical)` and writer perform a separate authenticated deletion. The R17 copy does not invoke that operation or inherit deletion authority. After copying, a caller may read the new package and request R18 removal independently; its own ownership and source/final reference checks decide whether deletion is safe. Neither operation establishes an archival retention decision or complete package reachability for 6.2.1(e) and 6.15.

**Closure gate:** a future full-standard claim requires each Open/Partial/External item to gain named, versioned evidence or an explicit applicable exception. It also requires independent OFD renderer/reader comparison, attachment conversion corpus, font and image validation, source-to-output provenance, the cited standards, qualified PKI evidence where relevant, and archive-authority decisions. Until then the status is `Incomplete` even when every implemented rule passes.

## Concrete dependencies for the remaining work

| Clauses | Material or interface still needed | Why the current package cannot settle it |
| --- | --- | --- |
| 5.2, 6.2.1(b), 6.2.6, 6.3.1, 6.6(b), 6.13 | Applicable editions and amendments of the undated section 2 references: GB/T 13000, GB 18030, GB/T 18894, GB/T 33190, GB/T 35275, GB/T 38540, DA/T 46, DA/T 47 and DA/T 54. Pin the exact clauses and a permitted test corpus before adding predicates. | The GB/T 42133 text delegates character ranges, base OFD semantics, metadata names, archive properties and signature formats; matching an XML name alone is insufficient. |
| 6.2.1(d)-(e), 6.2.2(d)-(e), 6.2.3(a)-(b),(e), 6.3.2, 6.3.3(a)-(b), 6.9(b), 6.10.1(a)(1),(b), 6.12(b), 6.14 | A stable source OFD, chosen output, source-to-output object identity and an auditable transformation manifest. The output-changing handlers also need generated OFD→PDF pages, PNG comparison and macOS Preview review. | A final ZIP cannot prove that Versions were retained, removed content was summarized, links were migrated or appearance was frozen. The present API is read-only and changes no page. |
| 6.2.6(b)-(g), 6.3.1(b)-(c), 6.5(a), 6.7-6.8 | Authenticated resource ownership and media bytes, font-table/glyph validators, image decoders or signatures for the six permitted raster formats, profile readers and a corpus of page/resource sharing cases. GB/T 20090 applicable editions are needed for Annex A AVS rows. | Filename or XML declaration checks do not establish decoded format, used glyphs, color output, resource sharing or audio/video conversion fidelity. |
| 6.10.2, 6.11, 6.15 and Annex A | Archive retention policy and accession receipts; source attachments and their content classification; converters and independent output fixtures for each applicable Annex A row, including order, text, visual and information-loss checks. | R18 physical deletion proves only a requested owned payload removal. It neither chooses retention nor preserves removed content by itself. |
| 6.13 | Authenticated source signature/seal, applicable GB/T 35275 and GB/T 38540 criteria, qualified CA/registration evidence, certificate chain and revocation history, trusted signing time, and archive-authority acceptance of long-term status services. | Local inventory or a copied signature value cannot establish cryptographic validity, legal effect or permissible de-technicalization. |
| 7.1-7.2 | A named target writer/reader build and independent reader tests for full output, untreated-item reporting and OFD-A edit lock, plus archive-policy sign-off. | Passing this library's unit tests does not observe another reader's UI, rendering or archive behavior. |

Each row is an actual missing input or operation boundary. It is not converted to `Pass` by absence of the named XML element, and no publication or archival certification is implied by this ledger.
