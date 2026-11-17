using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Medical;

/// <summary>Modal strength of one GB/T 48666-2026 matrix row.</summary>
public enum MedicalClauseStrength
{
    /// <summary>"应 / 不应": a requirement on the document.</summary>
    Shall,
    /// <summary>"宜 / 一般": a recommendation. Never yields a clause-level Fail.</summary>
    Should,
    /// <summary>"可": a permission. Never yields a clause-level Fail.</summary>
    May,
    /// <summary>"应支持 / 具备": a capability of the producing system. One file shows evidence of it but cannot violate it.</summary>
    Capability,
    /// <summary>No modal keyword: framework description or scope statement.</summary>
    Statement
}

/// <summary>How a matrix row can be decided.</summary>
public enum MedicalClauseCategory
{
    /// <summary>Decidable from the package bytes by the mapped rules.</summary>
    PackageDecidable,
    /// <summary>The mapped rules check only part of the clause; the rest needs other evidence.</summary>
    PartiallyDecidable,
    /// <summary>Needs workflow, source or institution evidence (generation, exchange, retention processes).</summary>
    ProcessEvidence,
    /// <summary>Needs law, PKI, referenced-standard, visual or industry-standard evidence.</summary>
    ExternalEvidence,
    /// <summary>A requirement on generating or reading software (chapter 9 and similar), not on one file.</summary>
    SoftwareRequirement
}

/// <summary>Whether the reading of the row's modal scope is settled.</summary>
public enum MedicalMatrixConfidence
{
    /// <summary>The modal keyword and its scope are unambiguous in the text.</summary>
    Verified,
    /// <summary>The text admits more than one reading, or conflicts with another clause; see the row note.</summary>
    Unverified
}

/// <summary>One row of the GB/T 48666-2026 clause matrix.</summary>
public sealed class MedicalClauseMatrixEntry
{
    internal MedicalClauseMatrixEntry(string clause, MedicalClauseStrength strength, MedicalClauseCategory category,
        string summary, string[] rules, MedicalMatrixConfidence confidence = MedicalMatrixConfidence.Verified, string? note = null)
    {
        Clause = clause; Strength = strength; Category = category; Summary = summary;
        Rules = Array.AsReadOnly(rules); Confidence = confidence; Note = note;
    }
    /// <summary>Clause identifier as printed, with list items in parentheses, for example <c>7.2(d)</c>.</summary>
    public string Clause { get; }
    /// <summary>Modal strength.</summary>
    public MedicalClauseStrength Strength { get; }
    /// <summary>How the row can be decided.</summary>
    public MedicalClauseCategory Category { get; }
    /// <summary>Short English paraphrase; not normative text.</summary>
    public string Summary { get; }
    /// <summary>Executable rule identifiers (<c>MH-0xx</c>) mapped to the row; empty when none exists.</summary>
    public IReadOnlyList<string> Rules { get; }
    /// <summary>Reading confidence.</summary>
    public MedicalMatrixConfidence Confidence { get; }
    /// <summary>Interpretation note, or null.</summary>
    public string? Note { get; }
}

/// <summary>
/// The reviewed clause matrix for GB/T 48666-2026 chapters 5 to 9. The human-readable baseline is
/// <c>docs/engineering/baselines/gbt48666-2026-clause-matrix.md</c>; a test keeps both in step.
/// </summary>
public static class MedicalClauseMatrix
{
    /// <summary>Assessed standard.</summary>
    public const string StandardId = "GB/T 48666-2026";
    /// <summary>Matrix revision; bump when any row changes.</summary>
    public const string Revision = "gbt48666-2026-matrix/1";
    /// <summary>SHA-256 of the reviewed standard copy (19 PDF pages). The copy is not committed.</summary>
    public const string StandardSourceSha256 = "3d2f59ea5f7d03c72d2f4257e5bf3e0f110ab87214792ea526eea31f7a979e7a";

    private const string ResourceTableConflict =
        "6.2 table (PublicRes.xml) places fonts in PublicRes; 7.2(a) registers fonts in DocumentRes. Fonts outside PageRes are Unknown until the conflict is resolved.";

    private static readonly MedicalClauseMatrixEntry[] Rows =
    {
        new("5.1", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Listed medical record scenarios conform or are converted", new string[0]),
        new("5.2", MedicalClauseStrength.Statement, MedicalClauseCategory.ProcessEvidence, "Reference application framework: institution file system collects, generates, converts and signs", new string[0]),
        new("5.3.1(a)", MedicalClauseStrength.May, MedicalClauseCategory.ProcessEvidence, "Business-system documents may be protected by e-seal or digital signature", new string[0]),
        new("5.3.1(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.ProcessEvidence, "Non-electronic materials become double-layer documents within the time limit", new string[0]),
        new("5.3.1(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.ProcessEvidence, "External records are audited, then digitized, checked or converted", new string[0]),
        new("5.3.2(a)", MedicalClauseStrength.Should, MedicalClauseCategory.ProcessEvidence, "Before exchange, combine per patient and linearize, or package as a document set", new string[0]),
        new("5.3.2(b)", MedicalClauseStrength.May, MedicalClauseCategory.ProcessEvidence, "Exchange may provide variants without structured data or as double-layer documents", new string[0]),
        new("5.3.2(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.ProcessEvidence, "Exchanged documents use a permission control mechanism", new string[0]),
        new("5.3.2(d)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Exchange of personal information follows the personal information protection law", new string[0]),
        new("5.3.2(e)", MedicalClauseStrength.May, MedicalClauseCategory.ProcessEvidence, "Exchange process information may be recorded as annotations", new string[0]),
        new("5.3.3", MedicalClauseStrength.Capability, MedicalClauseCategory.SoftwareRequirement, "Search and location of key content, version and annotation records", new string[0]),
        new("6.1.1", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Keep the OFD container + document framework", new[] { "MH-002" }),
        new("6.1.2(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "The carrier format is OFD with a single document", new[] { "MH-002" }),
        new("6.1.2(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Data and organization meet the relevant GB/T 33190 and GB/T 42133 requirements", new string[0],
            MedicalMatrixConfidence.Unverified, "\"Relevant\" is undefined; see the GB/T 42133 inheritance ledger."),
        new("6.1.3(a)", MedicalClauseStrength.Capability, MedicalClauseCategory.PartiallyDecidable, "Documents support linearization through the ZIP container", new[] { "MH-005" }),
        new("6.1.3(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "ZIP 6.2.0 or later; compressible content uses Deflate", new[] { "MH-004" },
            MedicalMatrixConfidence.Unverified, "Read as: stored (uncompressed) XML entries violate the Deflate requirement."),
        new("6.2(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Package entries use the canonical names and folder structure of figure 3", new[] { "MH-006" },
            MedicalMatrixConfidence.Unverified, "Table 1 calls the file names typical; only the folder structure is enforced."),
        new("6.2(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "Color spaces and fonts are defined in the public resources", new[] { "MH-007" },
            MedicalMatrixConfidence.Unverified, ResourceTableConflict),
        new("6.2(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "XML namespace matches the GB/T 33190 edition; every root declares the ofd prefix", new[] { "MH-003" }),
        new("6.2(d)", MedicalClauseStrength.Should, MedicalClauseCategory.PackageDecidable, "Elements use the ofd prefix; attributes are unqualified", new[] { "MH-003" }),
        new("6.3(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Signatures and seals comply with national law", new string[0]),
        new("6.3(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Cryptographic identifiers follow GB/T 33560", new[] { "MH-014" }),
        new("6.3(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Digital signatures follow GB/T 35275", new string[0]),
        new("6.3(d)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Seals and e-seal data follow GB/T 38540", new string[0]),
        new("6.3(e)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Signature or seal protects the page description content", new[] { "MH-013" }),
        new("6.3(f)", MedicalClauseStrength.Capability, MedicalClauseCategory.SoftwareRequirement, "Signed files accept annotations without changing protected content", new string[0]),
        new("6.3(g)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Traceable modifications are re-signed or counter-signed", new[] { "MH-013" }),
        new("6.3(h)", MedicalClauseStrength.Should, MedicalClauseCategory.ExternalEvidence, "Signature data includes a timestamp", new string[0]),
        new("7.1(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "Root DocType is OFD-H", new[] { "MH-001" }),
        new("7.1(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Fonts, images, audio, video and models are stored inside the container", new[] { "MH-006" }),
        new("7.1(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Description metadata is stored in DocInfo or DocInfo_N.xml", new[] { "MH-011" }),
        new("7.1(d)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "Custom tags are stored in the Tags folder", new[] { "MH-006" }),
        new("7.1(e)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "Attachments are stored in the Attachs folder", new[] { "MH-006" }),
        new("7.1(f)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "Annotations are stored in the Annots folder", new[] { "MH-006" }),
        new("7.2(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "Fonts, draw parameters and shared resources are registered in DocumentRes; single-page resources in PageRes", new[] { "MH-007" },
            MedicalMatrixConfidence.Unverified, ResourceTableConflict),
        new("7.2(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "All fonts needed for display and print are embedded unless exempt", new[] { "MH-008" }),
        new("7.2(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.PackageDecidable, "Image resources needed for display and print are in the file", new[] { "MH-009" }),
        new("7.2(d)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Images are BMP, JPEG, JPEG2000, TIFF, PNG or JBIG2; converted DICOM keeps clinical identifiers", new[] { "MH-009" }),
        new("7.2(e)", MedicalClauseStrength.Should, MedicalClauseCategory.PartiallyDecidable, "Associated objects within the institution limit (about 512 MB) are embedded, larger ones external", new[] { "MH-010" }),
        new("7.2(f)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Embedded audio uses MP3 or another common format", new[] { "MH-010" }),
        new("7.2(g)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Embedded video uses MP4 or another common format", new[] { "MH-010" }),
        new("7.2(h)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Embedded media clarity meets medical accuracy needs", new string[0]),
        new("7.3", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Rendering matches or approximates the traditional carrier", new string[0]),
        new("7.3(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Metadata identifies patient, responsible party, creation time and keywords", new[] { "MH-011" }),
        new("7.3(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "File identification gives file type, secrecy level and document identifier", new[] { "MH-011" }),
        new("7.3(c)", MedicalClauseStrength.Should, MedicalClauseCategory.ProcessEvidence, "Structured data is stored as custom tags or attachments consistent with the pages", new string[0]),
        new("7.3(d)", MedicalClauseStrength.Should, MedicalClauseCategory.PartiallyDecidable, "Document carries effectiveness information (signature or seal)", new[] { "MH-013" }),
        new("7.3(e)", MedicalClauseStrength.Should, MedicalClauseCategory.ProcessEvidence, "Scanned images on the background layer, OCR text on the body layer", new string[0]),
        new("7.3(f)", MedicalClauseStrength.Shall, MedicalClauseCategory.ProcessEvidence, "Embedded associated objects have internal jump actions; external ones have external links", new string[0]),
        new("7.3(g)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Metadata entry, annotation, combination and signing are not prohibited", new[] { "MH-015" }),
        new("7.4(a)", MedicalClauseStrength.Capability, MedicalClauseCategory.SoftwareRequirement, "Attachments can carry non-document materials such as DICOM, audio and video", new string[0]),
        new("7.4(b)", MedicalClauseStrength.Capability, MedicalClauseCategory.SoftwareRequirement, "Attachments can be internal or external", new string[0]),
        new("7.4(c)", MedicalClauseStrength.Should, MedicalClauseCategory.PartiallyDecidable, "Materials over the institution limit or needing special tools are external attachments", new[] { "MH-010" }),
        new("7.5(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Machine-readable information is semi-structured text such as XML or JSON", new[] { "MH-012" }),
        new("7.5(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Machine-readable files support GB 18030 characters and use UTF-8", new[] { "MH-012" }),
        new("7.5(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Multi-part machine-readable files list header, metadata and identification information", new string[0],
            MedicalMatrixConfidence.Verified, "The standard defines no schema; field mapping needs an industry specification."),
        new("7.5(d)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Organization supports fast retrieval and follows health-industry standards", new string[0]),
        new("7.6", MedicalClauseStrength.Shall, MedicalClauseCategory.ProcessEvidence, "Versioned records keep all historical versions or store changes as internal attachments", new string[0]),
        new("8(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Cryptographic technology and products meet national cryptography authority requirements", new string[0]),
        new("8(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Encryption only for temporary files; decrypt before archiving or migration", new[] { "MH-016" }),
        new("8(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Signature protection covers all content except annotation and signature lists", new[] { "MH-013" }),
        new("8(d)", MedicalClauseStrength.Shall, MedicalClauseCategory.PartiallyDecidable, "Signatures use national commercial cryptographic algorithms", new[] { "MH-014" }),
        new("8(e)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Verification does not depend on specific medical software", new string[0]),
        new("9.1(a)", MedicalClauseStrength.Capability, MedicalClauseCategory.SoftwareRequirement, "Generator creates OFD from data and files", new string[0]),
        new("9.1(b)", MedicalClauseStrength.Capability, MedicalClauseCategory.SoftwareRequirement, "Generator converts multiple formats to OFD", new string[0]),
        new("9.1(c)", MedicalClauseStrength.Capability, MedicalClauseCategory.SoftwareRequirement, "Generator processes OFD per GB/T 42133", new string[0]),
        new("9.1(d)", MedicalClauseStrength.Capability, MedicalClauseCategory.SoftwareRequirement, "Generation, conversion and processing follow this standard", new string[0]),
        new("9.2(a)", MedicalClauseStrength.Shall, MedicalClauseCategory.SoftwareRequirement, "Reader renders, verifies seals, controls permissions, locates annotations, shows metadata", new string[0]),
        new("9.2(b)", MedicalClauseStrength.Shall, MedicalClauseCategory.SoftwareRequirement, "Reader authorizes page editing and disables it for OFD-H files", new string[0],
            MedicalMatrixConfidence.Verified, "MedicalEditPolicy classifies operations for callers; reader behavior is not assessed."),
        new("9.2(c)", MedicalClauseStrength.Shall, MedicalClauseCategory.SoftwareRequirement, "Format check follows long-term preservation rules and reports unprocessed items", new string[0]),
        new("9.2(d)", MedicalClauseStrength.Should, MedicalClauseCategory.SoftwareRequirement, "Reader opens external attachments", new string[0]),
        new("9.2(e)", MedicalClauseStrength.Shall, MedicalClauseCategory.ExternalEvidence, "Reader verified by a third party for trusted viewing and audit", new string[0])
    };

    /// <summary>All rows in clause order.</summary>
    public static IReadOnlyList<MedicalClauseMatrixEntry> Entries { get; } = Array.AsReadOnly(Rows);

    /// <summary>Returns the row for an exact clause identifier.</summary>
    /// <param name="clause">Clause identifier, for example <c>7.1(a)</c>.</param>
    /// <returns>The matrix row.</returns>
    /// <exception cref="ArgumentException">The clause is not in the matrix.</exception>
    public static MedicalClauseMatrixEntry Get(string clause) =>
        Rows.FirstOrDefault(x => x.Clause == clause) ?? throw new ArgumentException("Unknown GB/T 48666 clause.", nameof(clause));
}
