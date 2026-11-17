using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Archival;

/// <summary>Modal strength observed on the rendered standard page for one ledger item.</summary>
public enum ArchivalClauseStrength
{
    /// <summary>"应 / 不应" (including "还应"): a requirement.</summary>
    Shall,
    /// <summary>No modal keyword: a treatment statement. Mapped rule failures keep their existing Fail semantics.</summary>
    Statement,
    /// <summary>"宜 / 不宜": a recommendation. Never yields a clause-level Fail in this package.</summary>
    Should,
    /// <summary>"可 / 不必 / 可不": a permission. Never yields a clause-level Fail in this package.</summary>
    May
}

/// <summary>How a ledger item can be decided for the first archival input profile.</summary>
public enum ArchivalProfileApplicability
{
    /// <summary>Decidable from package bytes alone, once an executable check exists.</summary>
    PackageDecidable,
    /// <summary>Object excluded by the profile; NotApplicable only after a complete-scope absence proof.</summary>
    AbsenceProvable,
    /// <summary>Needs source-to-result tracing; one result package cannot decide it.</summary>
    SourceEvidence,
    /// <summary>Needs policy, PKI, referenced-edition, independent reader or archival-authority material.</summary>
    ExternalEvidence,
    /// <summary>Applies only to inputs the first profile does not accept, such as signatures or attachments.</summary>
    OutOfProfile
}

/// <summary>Whether the scope of the modal keyword was unambiguous on the rendered page.</summary>
public enum ArchivalMatrixConfidence
{
    /// <summary>The modal keyword and its scope were confirmed.</summary>
    Verified,
    /// <summary>The keyword is visible but its scope admits more than one reading.</summary>
    Unverified
}

/// <summary>A secondary modal strength present in the same ledger item.</summary>
public sealed class ArchivalMixedStrength
{
    internal ArchivalMixedStrength(ArchivalClauseStrength strength, string part) { Strength = strength; Part = part; }
    /// <summary>The secondary strength.</summary>
    public ArchivalClauseStrength Strength { get; }
    /// <summary>Short paraphrase of the part carrying that strength; not normative text.</summary>
    public string Part { get; }
}

/// <summary>One row of the clause strength and first-profile applicability matrix.</summary>
public sealed class ArchivalClauseMatrixEntry
{
    internal ArchivalClauseMatrixEntry(string clause, ArchivalClauseStrength strength, string modalKeyword,
        ArchivalMixedStrength[] mixedStrengths, bool conditional, string? condition, int pdfPage, string printedPage,
        ArchivalMatrixConfidence confidence, ArchivalProfileApplicability firstProfile, string[] rules,
        string[] referencedStandards, bool editionFixed, string? note)
    {
        Clause = clause; Strength = strength; ModalKeyword = modalKeyword;
        MixedStrengths = Array.AsReadOnly(mixedStrengths); Conditional = conditional; Condition = condition;
        PdfPage = pdfPage; PrintedPage = printedPage; Confidence = confidence; FirstProfile = firstProfile;
        Rules = Array.AsReadOnly(rules); ReferencedStandards = Array.AsReadOnly(referencedStandards);
        EditionFixed = editionFixed; Note = note;
    }
    /// <summary>Exact clause identifier from <c>ArchivalClauseCatalog</c>.</summary>
    public string Clause { get; }
    /// <summary>Strongest modal strength in the item.</summary>
    public ArchivalClauseStrength Strength { get; }
    /// <summary>Modal keyword as printed ("none" when absent).</summary>
    public string ModalKeyword { get; }
    /// <summary>Other strengths present in the same item.</summary>
    public IReadOnlyList<ArchivalMixedStrength> MixedStrengths { get; }
    /// <summary>Whether the item applies only under a condition.</summary>
    public bool Conditional { get; }
    /// <summary>Paraphrased condition, or null.</summary>
    public string? Condition { get; }
    /// <summary>One-based page index in the inspected standard PDF.</summary>
    public int PdfPage { get; }
    /// <summary>Printed page label.</summary>
    public string PrintedPage { get; }
    /// <summary>Whether the modal scope is confirmed.</summary>
    public ArchivalMatrixConfidence Confidence { get; }
    /// <summary>Decidability for the first archival input profile.</summary>
    public ArchivalProfileApplicability FirstProfile { get; }
    /// <summary>Mapped rule identities, or planned work-package names such as A1b/A1c.</summary>
    public IReadOnlyList<string> Rules { get; }
    /// <summary>Referenced standards whose edition matters for this item.</summary>
    public IReadOnlyList<string> ReferencedStandards { get; }
    /// <summary>Whether the referenced edition has been fixed. Currently false for every item.</summary>
    public bool EditionFixed { get; }
    /// <summary>Reviewer note, or null.</summary>
    public string? Note { get; }
}

/// <summary>
/// Engineering classification of every <c>ArchivalClauseCatalog</c> item by modal strength and first-profile
/// applicability. This is not an interpretation issued by a standards body or archival authority.
/// </summary>
public static partial class ArchivalClauseMatrix
{
    /// <summary>Revision of this embedded matrix; changes whenever any row changes.</summary>
    public const string Revision = "ofdrw-archival-matrix/1";
    /// <summary>Ledger revision the matrix rows are aligned with.</summary>
    public const string LedgerRevision = "ofdrw-archival-ledger/1";
    /// <summary>Standard classified by this matrix.</summary>
    public const string StandardId = "GB/T 42133-2022";
    /// <summary>SHA-256 of the inspected standard copy.</summary>
    public const string StandardSourceSha256 = "7d76c086e1384ff200c0df8f5a40bf67660e841340c5e2630e582c81c22f1b82";
    /// <summary>Date the rendered pages were reviewed.</summary>
    public const string ReviewedOn = "2026-11-17";
    /// <summary>Repository source of the generated rows.</summary>
    public const string SourceDocument = "docs/engineering/baselines/gbt42133-2022-clause-applicability.json";

    private static readonly IReadOnlyList<ArchivalClauseMatrixEntry> EntriesValue = Array.AsReadOnly(CreateEntries());
    private static readonly Dictionary<string, ArchivalClauseMatrixEntry> ByClause =
        EntriesValue.ToDictionary(x => x.Clause, StringComparer.Ordinal);

    /// <summary>Strength precedence used to pick the item strength: Shall, Statement, Should, May.</summary>
    public static IReadOnlyList<ArchivalClauseStrength> StrengthOrder { get; } = Array.AsReadOnly(new[]
    { ArchivalClauseStrength.Shall, ArchivalClauseStrength.Statement, ArchivalClauseStrength.Should, ArchivalClauseStrength.May });

    /// <summary>All 105 rows in ledger order.</summary>
    public static IReadOnlyList<ArchivalClauseMatrixEntry> Entries => EntriesValue;

    /// <summary>Returns the row for an exact ledger clause identifier.</summary>
    /// <exception cref="ArgumentException">The clause is not in the ledger.</exception>
    public static ArchivalClauseMatrixEntry Get(string clause) =>
        clause != null && ByClause.TryGetValue(clause, out var entry) ? entry :
        throw new ArgumentException("Unknown ledger clause.", nameof(clause));
}
