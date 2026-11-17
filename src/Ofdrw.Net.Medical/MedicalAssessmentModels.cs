using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Archival;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Medical;

/// <summary>Whether the bounded, read-only assessment ran to the end. Cancellation throws instead.</summary>
public enum MedicalExecutionState
{
    /// <summary>Every inspection finished. Individual results can still be Unknown.</summary>
    Complete,
    /// <summary>A byte, entry, XML or reference budget stopped inspection; the report is not deliverable.</summary>
    BudgetExhausted,
    /// <summary>The input was rejected as malformed or unsafe; the report is not deliverable.</summary>
    Error
}

/// <summary>The raw root <c>DocType</c> declaration.</summary>
public enum MedicalDeclarationStatus
{
    /// <summary>The unqualified attribute is exactly the ASCII value <c>OFD-H</c>.</summary>
    Medical,
    /// <summary>The value is <c>OFD</c>, a non-ASCII dash and <c>H</c>, for example copied from the typeset standard.</summary>
    TypesetVariant,
    /// <summary>Another value. This is not an invalid ordinary OFD.</summary>
    NotMedical,
    /// <summary>No unqualified attribute exists.</summary>
    Missing,
    /// <summary>A namespace-qualified attribute with the same local name competes with the declaration.</summary>
    Ambiguous,
    /// <summary>OFD.xml could not be inspected.</summary>
    Unavailable
}

/// <summary>Package location and bounded metadata, without payload content.</summary>
public sealed class MedicalEvidence
{
    internal MedicalEvidence(string entry, string selector, string observed, string expected)
    {
        Entry = entry; Selector = selector;
        Observed = Bound(observed); Expected = Bound(expected);
    }
    /// <summary>Package entry, or <c>&lt;package&gt;</c> for package-level observations.</summary>
    public string Entry { get; }
    /// <summary>Element or attribute selector, or a package-level marker.</summary>
    public string Selector { get; }
    /// <summary>Sanitized observed metadata (at most 512 characters).</summary>
    public string Observed { get; }
    /// <summary>The bounded expectation.</summary>
    public string Expected { get; }

    private static string Bound(string value)
    {
        var clean = new string((value ?? string.Empty).Select(c => char.IsControl(c) ? ' ' : c).ToArray());
        return clean.Length <= 512 ? clean : clean.Substring(0, 509) + "...";
    }
}

/// <summary>Size and digest of the exact bytes inspected.</summary>
public sealed class MedicalInputIdentity
{
    internal MedicalInputIdentity(long length, string sha256) { Length = length; Sha256 = sha256; }
    /// <summary>Number of bytes inspected.</summary>
    public long Length { get; }
    /// <summary>Lowercase SHA-256 of those bytes.</summary>
    public string Sha256 { get; }
}

/// <summary>One executable rule observation on one GB/T 48666 clause.</summary>
public sealed class MedicalRuleObservation
{
    internal MedicalRuleObservation(string ruleId, string clause, PreflightOutcome outcome, string reasonCode, MedicalEvidence evidence)
    { RuleId = ruleId; Clause = clause; Outcome = outcome; ReasonCode = reasonCode; Evidence = evidence; }
    /// <summary>Rule identity, <c>MH-001</c> to <c>MH-016</c>.</summary>
    public string RuleId { get; }
    /// <summary>Matrix clause the observation applies to.</summary>
    public string Clause { get; }
    /// <summary>Rule outcome before strength-based aggregation.</summary>
    public PreflightOutcome Outcome { get; }
    /// <summary>Stable reason code, for example <c>NamespaceNotStandard</c>.</summary>
    public string ReasonCode { get; }
    /// <summary>Where the observation was made.</summary>
    public MedicalEvidence Evidence { get; }
}

/// <summary>Declaration dimension (7.1(a)).</summary>
public sealed class MedicalDeclarationAssessment
{
    internal MedicalDeclarationAssessment(MedicalDeclarationStatus status, string? rawDocType, MedicalRuleObservation? observation)
    { Status = status; RawDocType = rawDocType; Observation = observation; }
    /// <summary>Declaration status.</summary>
    public MedicalDeclarationStatus Status { get; }
    /// <summary>Raw unqualified <c>DocType</c>, or null.</summary>
    public string? RawDocType { get; }
    /// <summary>The MH-001 observation, or null when OFD.xml was not inspected.</summary>
    public MedicalRuleObservation? Observation { get; }
}

/// <summary>Clause dimension: one entry per matrix row.</summary>
public sealed class MedicalClauseAssessment
{
    internal MedicalClauseAssessment(MedicalClauseMatrixEntry matrix, IReadOnlyList<MedicalRuleObservation> observations,
        IReadOnlyList<MedicalRuleObservation> advisories, PreflightOutcome outcome)
    { Matrix = matrix; Observations = observations; Advisories = advisories; Outcome = outcome; }
    /// <summary>Clause identifier.</summary>
    public string Clause => Matrix.Clause;
    /// <summary>Matrix row: strength, category, confidence and note.</summary>
    public MedicalClauseMatrixEntry Matrix { get; }
    /// <summary>Whether at least one executable rule observed the clause.</summary>
    public bool HasExecutableCheck => Observations.Count > 0;
    /// <summary>All rule observations for the clause, unchanged.</summary>
    public IReadOnlyList<MedicalRuleObservation> Observations { get; }
    /// <summary>Unmet recommendations, permissions or capabilities. They never produce a clause Fail.</summary>
    public IReadOnlyList<MedicalRuleObservation> Advisories { get; }
    /// <summary>
    /// Fail, Unknown, Pass or NotApplicable over the mapped bounded checks only. A clause without a check is Unknown;
    /// Pass on a <see cref="MedicalClauseCategory.PartiallyDecidable"/> row covers only the checked part.
    /// </summary>
    public PreflightOutcome Outcome { get; }
}

/// <summary>One GB/T 42133 ledger item seen through the inheritance ledger.</summary>
public sealed class MedicalInheritedArchivalClause
{
    internal MedicalInheritedArchivalClause(MedicalArchivalInheritanceEntry inheritance, PreflightOutcome archivalOutcome)
    { Inheritance = inheritance; ArchivalOutcome = archivalOutcome; }
    /// <summary>GB/T 42133 clause.</summary>
    public string ArchivalClause => Inheritance.ArchivalClause;
    /// <summary>Disposition and basis.</summary>
    public MedicalArchivalInheritanceEntry Inheritance { get; }
    /// <summary>Unchanged clause outcome from the embedded <see cref="ArchivalAssessmentReport"/>.</summary>
    public PreflightOutcome ArchivalOutcome { get; }
    /// <summary>True when the clause is inherited and failed: it makes the medical standard outcome Fail.</summary>
    public bool ContributesFail => Inheritance.Disposition == MedicalArchivalDisposition.Inherited && ArchivalOutcome == PreflightOutcome.Fail;
}

/// <summary>Document dimension. There is no full-standard Pass.</summary>
public sealed class MedicalDocumentAssessment
{
    internal MedicalDocumentAssessment(IReadOnlyList<string> checkedClauses, IReadOnlyList<string> uncheckedClauses,
        IReadOnlyList<string> gaps, PreflightOutcome standardOutcome)
    { CheckedClauses = checkedClauses; UncheckedClauses = uncheckedClauses; Gaps = gaps; StandardOutcome = standardOutcome; }
    /// <summary>Assessed standard.</summary>
    public string StandardId => MedicalClauseMatrix.StandardId;
    /// <summary>SHA-256 of the reviewed standard copy.</summary>
    public string StandardSourceSha256 => MedicalClauseMatrix.StandardSourceSha256;
    /// <summary>Clauses with at least one executable check.</summary>
    public IReadOnlyList<string> CheckedClauses { get; }
    /// <summary>Clauses without an executable check; they stay Unknown.</summary>
    public IReadOnlyList<string> UncheckedClauses { get; }
    /// <summary>Stable named coverage gaps.</summary>
    public IReadOnlyList<string> Gaps { get; }
    /// <summary>Always <c>Incomplete</c>.</summary>
    public string CoverageStatus => "Incomplete";
    /// <summary>Always <c>NotAssessed</c>.</summary>
    public string Certification => "NotAssessed";
    /// <summary>Fail when a medical clause or an inherited GB/T 42133 clause failed, otherwise Unknown. Never Pass.</summary>
    public PreflightOutcome StandardOutcome { get; }
}

/// <summary>Read-only GB/T 48666-2026 assessment of one input snapshot. Not a certificate.</summary>
public sealed class MedicalAssessmentReport
{
    internal MedicalAssessmentReport(MedicalExecutionState state, string reason, MedicalInputIdentity? input,
        MedicalDeclarationAssessment declaration, IReadOnlyList<MedicalClauseAssessment> clauses,
        IReadOnlyList<MedicalInheritedArchivalClause> inherited, MedicalDocumentAssessment document, ArchivalAssessmentReport? archival)
    {
        ExecutionState = state; ExecutionReason = reason; Input = input; Declaration = declaration; Clauses = clauses;
        InheritedArchivalClauses = inherited; Document = document; Archival = archival;
    }
    /// <summary>Report contract revision.</summary>
    public string AssessmentRevision => MedicalAssessmentService.Revision;
    /// <summary>Clause matrix revision.</summary>
    public string MatrixRevision => MedicalClauseMatrix.Revision;
    /// <summary>GB/T 42133 inheritance ledger revision.</summary>
    public string InheritanceRevision => MedicalArchivalInheritance.Revision;
    /// <summary>Execution dimension. Only Complete reports describe a finished inspection.</summary>
    public MedicalExecutionState ExecutionState { get; }
    /// <summary>Stable reason for the execution state.</summary>
    public string ExecutionReason { get; }
    /// <summary>Identity of the inspected bytes; null when the input exceeded the byte budget.</summary>
    public MedicalInputIdentity? Input { get; }
    /// <summary>Declaration dimension.</summary>
    public MedicalDeclarationAssessment Declaration { get; }
    /// <summary>Clause dimension in matrix order.</summary>
    public IReadOnlyList<MedicalClauseAssessment> Clauses { get; }
    /// <summary>Every GB/T 42133 ledger item with its disposition and the embedded archival outcome.</summary>
    public IReadOnlyList<MedicalInheritedArchivalClause> InheritedArchivalClauses { get; }
    /// <summary>Document dimension.</summary>
    public MedicalDocumentAssessment Document { get; }
    /// <summary>The unchanged GB/T 42133 assessment of the same bytes, or null when it was not run.</summary>
    public ArchivalAssessmentReport? Archival { get; }

    /// <summary>Returns the entry for an exact GB/T 48666 clause.</summary>
    /// <param name="clause">Clause identifier.</param>
    /// <returns>The clause assessment.</returns>
    /// <exception cref="ArgumentException">The clause is not in the matrix.</exception>
    public MedicalClauseAssessment GetClause(string clause) =>
        Clauses.FirstOrDefault(x => x.Clause == clause) ?? throw new ArgumentException("Unknown GB/T 48666 clause.", nameof(clause));
}
