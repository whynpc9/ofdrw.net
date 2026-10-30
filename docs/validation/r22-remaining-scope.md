# R22 remaining scope after PR #76

Reconciled against research v1.2 §§12.4–12.6 on 2026-10-30, source
`84a74b48a27a138bf16b85903408a62f0bde7a6a`. The accepted PR #76 matrix is an
immutable bounded delivery, not completion of stages B–E. Its twelve mixed
samples/77 original glyphs and reader decision keep their original identity.

| Stage / contract | Delivered evidence | Remaining scope and disposition |
| --- | --- | --- |
| B: font/run/position identity | Immutable static faces, UTF-16 clusters, original GIDs, design units and typed origins; bounded original-GID subset/readback | Other formats/instances and arbitrary mappings are not admitted by these samples. |
| B: OFD indexing and interoperability | Complete BMP source spans with explicit origins; known GroupedLocal profile | Non-BMP CGTransform counts, arbitrary imported mappings and cross-TextCode semantics remain in deferred R02. Native external OFD-reader/version evidence is still required for additional profiles. |
| B/D: repeated glyph occurrences | The paragraph query already shapes repeated Hebrew letters; PR #76 uses per-scalar PDF aliases | Writer/package/explicit-reader distinct-GID guards still reject these words at the baseline. This is the next executable unit, below. |
| C: paragraph and final-line shaping | Latin/Han candidate-line shaping and bounded LTR Latin/Hebrew word-break query with final run shaping | These are restricted query engines, not a general bidi paragraph or new Flow. Mixed writer rejects wrapped output; general context-sensitive final-line shaping, line-level semantic order and shared positioned-result pagination remain open. |
| C: Span and context | Existing Flow preserves compatibility measurement/pagination; single-run requests have source-ranged features | No new context-preserving cross-Span shaping/itemization API or acceptance. Do not silently change legacy Flow to per-grapheme HarfBuzz calls. |
| C/D: minimum script corpus | AV/ffi feature toggles, decomposed Latin marks, limited Hebrew/Latin and Han samples | Arabic joining with digits/punctuation and an Indic script still lack end-to-end font/cluster/layout/output/reader evidence. These are future executable work, not claims blocked solely on an external author. |
| D: PDF export | One unwrapped line, exact R22 face at 4 mm, Latin around one/two adjacent Hebrew words, 2–8 distinct letters per word; original contours/scalars/positions plus actual Preview copy/Find | Repeated letters, 3+ words, edge/separated RTL, wrapping and other font/size profiles remain rejected at the baseline. A guard or merged bounded PR does not close D. |
| D: SVG | Actual original-GID paths and logical source metadata; independently rendered actual SVG pages inspected | `aria-label`/`title`/source metadata are not proof of native search, selection, clipboard or accessibility. Those semantics require SVG-specific checks and do not inherit PDF acceptance. |
| D: reverse import | Existing PDF→OFD semantic/fallback paths still use PdfPig default word extraction | Separate open defect. The export reader decision does not accept RTL reverse import. |
| E: expanded profiles | Explicit rejection of unsupported formats/controls | Variable static instances, vertical text, CFF/CFF2, color glyphs, UVS/ZWJ and other profiles each need their own specification/font/output/interop evidence. |
| Deployment | Actual native shaping on the three current CI hosts, pinned versions and expected assets | Extra RIDs, single-file, trimming/AOT and restricted hosts are not proved. Per-process resolved native path/hash and a Linux uname receipt were not independently attested. |

R02's full imported-input scope stays deferred, PR #67 stays stopped, and R20
stays externally blocked. This ledger does not reopen their implementation.
The former PDFKit logical-character-index method remains falsified at 0/8;
all reference negatives stay negative. PdfPig default `GetWords`/`Page.Text`
remain versioned compatibility observations, separate from original-character,
font, GID, position, physical geometry and actual Preview gates.

## Next bounded candidate: repeated Hebrew letters

Retain the exact face/4 mm, single-line and one/two-word bounds. Admit a repeated
GID only when every occurrence matches the reversed source scalar's nominal
original glyph, has its own source interval and explicit origin, and retains
the complete logical source. Different source scalars sharing one semantic GID,
foreign GIDs, missing occurrences and ambiguous generic mappings still reject.

The new corpus adds one repeated-letter word, two such words, a repeated whole
word, a palindrome and an eight-letter identical-glyph word to all twelve prior
cases. Functional tests first reproduce the baseline rejection, then check OFD
write/read, both mapping strategies and per-occurrence PDF characters. Exact
packages, independent outline multiplicity/geometry, SVG rendering and actual
Preview copying/search remain required before acceptance.

A palindrome's reverse query is identical to its logical query. The reader
probe records that alias, executes each distinct query once and derives its
expected count from the declared source words. It does not demand a positive
query also have zero reverse results. Non-present reversed queries retain zero
expectations. This does not alter the earlier failed methods or their files.

Candidate status is not final-main acceptance; new results must bind their own
source, package manifest, outputs and visual receipts.

At the initial working candidate, four new typed roundtrip cases fail on the
unchanged baseline and pass with the occurrence checks. The explicit PDF suite
passes 117 cases; all seventeen source-generated mixed pages pass 109 original
scalar/GID/baseline/visible-path checks. Missing, duplicate and changed-outline
occurrence controls reject. The changed query probe separately preserves all
twelve frozen baseline selections and Find expectations. These are preliminary
functional/API results, not exact-package, GUI or final-main acceptance.

A later pre-merge negative control found that a font cmap alias could map two
different Hebrew scalars to the same GID and pass the low-level projection.
The repair keeps a per-span GID-to-source-scalar map in both Legacy and
GroupedLocal projection, rejecting that conflict while admitting repeated
occurrences of the same scalar. The initial candidate results remain historical;
the repaired head needs its own regression, package and visual acceptance.
