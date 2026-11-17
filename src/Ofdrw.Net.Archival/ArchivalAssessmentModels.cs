using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival;

/// <summary>Whether the bounded, read-only assessment ran to the end. Cancellation is never a state: it throws.</summary>
public enum ArchivalExecutionState
{
    /// <summary>Every requested inspection finished. Individual results can still be Unknown.</summary>
    Complete,
    /// <summary>An input, ZIP, XML, reference or result budget stopped inspection; the report is not deliverable.</summary>
    BudgetExhausted,
    /// <summary>The input was rejected as malformed or unsafe; the report is not deliverable.</summary>
    Error
}

/// <summary>The raw root <c>DocType</c> declaration as found in <c>OFD.xml</c>.</summary>
public enum ArchivalDeclarationStatus
{
    /// <summary>The unqualified attribute is exactly <c>OFD-A</c> and unambiguous.</summary>
    Archival,
    /// <summary>The unqualified attribute is present but is not <c>OFD-A</c>; this is not an invalid ordinary OFD.</summary>
    NotArchival,
    /// <summary>No unqualified attribute exists in the raw XML. No reader fallback value is substituted.</summary>
    Missing,
    /// <summary>A foreign-namespace attribute with the same local name competes with the declaration.</summary>
    Ambiguous,
    /// <summary>The root could not be inspected, for example because inspection stopped or the root is unsupported.</summary>
    Unavailable
}

/// <summary>Whether the input lies inside the selected archival input profile.</summary>
public enum ArchivalInputSupport
{
    /// <summary>Every inspected declaration lies inside the profile. This is not a standards verdict.</summary>
    Supported,
    /// <summary>The input contains content the profile excludes. This is not a standards violation by itself.</summary>
    OutOfScope,
    /// <summary>The scope could not be established.</summary>
    Unknown
}

/// <summary>Package location and bounded metadata, without payload content.</summary>
public sealed class ArchivalEvidenceLocation
{
    internal ArchivalEvidenceLocation(string entry, string selector, string observed, string expected)
    { Entry = entry; Selector = selector; Observed = observed; Expected = expected; }
    internal static ArchivalEvidenceLocation From(PreflightEvidence evidence) =>
        new ArchivalEvidenceLocation(evidence.Entry, evidence.Selector, evidence.Observed, evidence.Expected);
    /// <summary>Package entry, or <c>&lt;input&gt;</c> / <c>&lt;package&gt;</c> for input-level observations.</summary>
    public string Entry { get; }
    /// <summary>Expanded-name selector or a package-level marker.</summary>
    public string Selector { get; }
    /// <summary>Sanitized observed metadata.</summary>
    public string Observed { get; }
    /// <summary>The bounded expectation.</summary>
    public string Expected { get; }
}

/// <summary>Size and digest of the exact bytes that were inspected.</summary>
public sealed class ArchivalInputIdentity
{
    internal ArchivalInputIdentity(long length, string sha256) { Length = length; Sha256 = sha256; }
    /// <summary>Number of bytes read from the input and inspected.</summary>
    public long Length { get; }
    /// <summary>Lowercase SHA-256 of those bytes.</summary>
    public string Sha256 { get; }
}

/// <summary>Declaration dimension: the raw root DocType, mapped from the existing AC-001 observation.</summary>
public sealed class ArchivalDeclarationAssessment
{
    internal ArchivalDeclarationAssessment(ArchivalDeclarationStatus status, string? rawDocType, bool foreignDocType,
        PreflightOutcome outcome, string reasonCode, ArchivalEvidenceLocation evidence)
    {
        Status = status; RawDocType = rawDocType; ForeignDocTypePresent = foreignDocType;
        Outcome = outcome; ReasonCode = reasonCode; Evidence = evidence;
    }
    /// <summary>Declaration status derived from AC-001.</summary>
    public ArchivalDeclarationStatus Status { get; }
    /// <summary>Unqualified <c>DocType</c> value exactly as in the raw XML, or null when absent or unavailable.</summary>
    public string? RawDocType { get; }
    /// <summary>Whether a namespace-qualified attribute named <c>DocType</c> is present on the root.</summary>
    public bool ForeignDocTypePresent { get; }
    /// <summary>Rule that produced the outcome.</summary>
    public string RuleId => "AC-001";
    /// <summary>Unchanged AC-001 outcome. A declared OFD-A label proves nothing beyond this dimension.</summary>
    public PreflightOutcome Outcome { get; }
    /// <summary>Unchanged AC-001 reason code.</summary>
    public string ReasonCode { get; }
    /// <summary>AC-001 evidence location.</summary>
    public ArchivalEvidenceLocation Evidence { get; }
}

/// <summary>One reason the input is outside, or not provably inside, the selected profile.</summary>
public sealed class ArchivalScopeFinding
{
    internal ArchivalScopeFinding(string code, ArchivalInputSupport support, ArchivalEvidenceLocation evidence)
    { Code = code; Support = support; Evidence = evidence; }
    /// <summary>Stable reason code, for example <c>SignaturesDeclared</c> or <c>UnrecognizedDocumentContent</c>.</summary>
    public string Code { get; }
    /// <summary>OutOfScope or Unknown; never Supported.</summary>
    public ArchivalInputSupport Support { get; }
    /// <summary>Where the reason was observed.</summary>
    public ArchivalEvidenceLocation Evidence { get; }
}

/// <summary>Input-support dimension for the selected profile.</summary>
public sealed class ArchivalInputSupportAssessment
{
    internal ArchivalInputSupportAssessment(ArchivalInputProfile profile, ArchivalInputSupport support,
        IReadOnlyList<ArchivalScopeFinding> findings)
    { Profile = profile; Support = support; Findings = findings; }
    /// <summary>Selected profile.</summary>
    public ArchivalInputProfile Profile { get; }
    /// <summary>OutOfScope dominates Unknown, which dominates Supported.</summary>
    public ArchivalInputSupport Support { get; }
    /// <summary>Every out-of-scope or unknown-scope reason, in inspection order.</summary>
    public IReadOnlyList<ArchivalScopeFinding> Findings { get; }
    /// <summary>Always false: producer provenance is not trusted; support is decided from package content only.</summary>
    public bool ProvenanceVerified => false;
}

/// <summary>One mapped rule observation on one ledger clause.</summary>
public sealed class ArchivalClauseObservation
{
    internal ArchivalClauseObservation(string ruleId, string owner, int ruleVersion, string clause,
        PreflightOutcome outcome, string reasonCode, ArchivalEvidenceLocation evidence)
    {
        RuleId = ruleId; RuleOwner = owner; RuleVersion = ruleVersion; Clause = clause;
        Outcome = outcome; ReasonCode = reasonCode; Evidence = evidence;
    }
    /// <summary>Rule identity, such as <c>AR-001</c> or <c>AC-008</c>.</summary>
    public string RuleId { get; }
    /// <summary>Assembly that owns the rule identity: <c>Ofdrw.Net.Packaging</c> or <c>Ofdrw.Net.Archival</c>.</summary>
    public string RuleOwner { get; }
    /// <summary>Rule implementation revision.</summary>
    public int RuleVersion { get; }
    /// <summary>Ledger clause the rule is mapped to.</summary>
    public string Clause { get; }
    /// <summary>Unchanged rule outcome. Clause-level aggregation may treat it differently by strength.</summary>
    public PreflightOutcome Outcome { get; }
    /// <summary>Stable reason code from the rule.</summary>
    public string ReasonCode { get; }
    /// <summary>Rule evidence location.</summary>
    public ArchivalEvidenceLocation Evidence { get; }
}

/// <summary>Clause dimension: one entry per ledger item, in ledger order.</summary>
public sealed class ArchivalClauseAssessment
{
    internal ArchivalClauseAssessment(ArchivalClauseDefinition definition, ArchivalClauseMatrixEntry matrix,
        IReadOnlyList<ArchivalClauseObservation> observations, IReadOnlyList<ArchivalClauseObservation> advisories,
        PreflightOutcome outcome, IReadOnlyList<string> notes)
    {
        Definition = definition; Matrix = matrix; Observations = observations; Advisories = advisories;
        Outcome = outcome; Notes = notes;
    }
    /// <summary>Exact ledger clause identifier.</summary>
    public string Clause => Definition.Clause;
    /// <summary>Unchanged ledger definition, including condition and mapped rule IDs.</summary>
    public ArchivalClauseDefinition Definition { get; }
    /// <summary>Strength and first-profile applicability row.</summary>
    public ArchivalClauseMatrixEntry Matrix { get; }
    /// <summary>Normative strength, exposed so callers can apply a stricter or looser policy.</summary>
    public ArchivalClauseStrength Strength => Matrix.Strength;
    /// <summary>Whether the item applies only under a condition.</summary>
    public bool Conditional => Matrix.Conditional;
    /// <summary>First-profile applicability category.</summary>
    public ArchivalProfileApplicability Applicability => Matrix.FirstProfile;
    /// <summary>Whether the modal scope is confirmed; 6.11 and 6.13.1(a)-(c) are Unverified.</summary>
    public ArchivalMatrixConfidence Confidence => Matrix.Confidence;
    /// <summary>Whether at least one executable rule observed this clause.</summary>
    public bool HasExecutableCheck => Observations.Count > 0;
    /// <summary>All mapped rule observations, unchanged.</summary>
    public IReadOnlyList<ArchivalClauseObservation> Observations { get; }
    /// <summary>Unmet recommendations or permissions (Should/May clauses). They never produce a clause Fail.</summary>
    public IReadOnlyList<ArchivalClauseObservation> Advisories { get; }
    /// <summary>
    /// Fail, Unknown, Pass or NotApplicable over the mapped bounded checks only. A clause without a check is Unknown.
    /// Pass never closes <see cref="RemainingEvidence"/>.
    /// </summary>
    public PreflightOutcome Outcome { get; }
    /// <summary>Evidence locations of all observations.</summary>
    public IReadOnlyList<ArchivalEvidenceLocation> Evidence => Observations.Select(x => x.Evidence).ToList().AsReadOnly();
    /// <summary>Remaining evidence still required by the ledger.</summary>
    public string RemainingEvidence => Definition.RemainingEvidence;
    /// <summary>Stable notes about assumptions behind the mapped checks.</summary>
    public IReadOnlyList<string> Notes { get; }
}

/// <summary>Document dimension: covered scope and gaps. There is no full-standard Pass.</summary>
public sealed class ArchivalDocumentAssessment
{
    internal ArchivalDocumentAssessment(IReadOnlyList<string> checkedClauses, IReadOnlyList<string> uncheckedClauses,
        IReadOnlyList<string> gaps, PreflightOutcome standardOutcome)
    { CheckedClauses = checkedClauses; UncheckedClauses = uncheckedClauses; Gaps = gaps; StandardOutcome = standardOutcome; }
    /// <summary>Assessed standard.</summary>
    public string StandardId => ArchivalClauseMatrix.StandardId;
    /// <summary>Edition year of the assessed standard.</summary>
    public string StandardVersion => "2022";
    /// <summary>SHA-256 of the inspected standard copy.</summary>
    public string StandardSourceSha256 => ArchivalClauseMatrix.StandardSourceSha256;
    /// <summary>Clauses with at least one executable bounded check.</summary>
    public IReadOnlyList<string> CheckedClauses { get; }
    /// <summary>Clauses with no executable check; they remain Unknown.</summary>
    public IReadOnlyList<string> UncheckedClauses { get; }
    /// <summary>Stable named coverage gaps.</summary>
    public IReadOnlyList<string> Gaps { get; }
    /// <summary>Always Incomplete.</summary>
    public string CoverageStatus => "Incomplete";
    /// <summary>Always NotAssessed.</summary>
    public string Certification => "NotAssessed";
    /// <summary>Fail when any clause failed, otherwise Unknown. Never Pass.</summary>
    public PreflightOutcome StandardOutcome { get; }
}

/// <summary>A dimension that cannot be inferred from a single file and is reported with named gaps.</summary>
public sealed class ArchivalUnassessedDimension
{
    internal ArchivalUnassessedDimension(string status, IReadOnlyList<string> clauses, IReadOnlyList<string> gaps)
    { Status = status; Clauses = clauses; Gaps = gaps; }
    /// <summary><c>NotAssessed</c> or <c>NotAuthenticated</c>.</summary>
    public string Status { get; }
    /// <summary>Ledger clauses that need this kind of evidence.</summary>
    public IReadOnlyList<string> Clauses { get; }
    /// <summary>Named missing assessments.</summary>
    public IReadOnlyList<string> Gaps { get; }
}

/// <summary>Read-only, multi-dimensional archival assessment of one input snapshot. Not a certificate.</summary>
public sealed class ArchivalAssessmentReport
{
    internal ArchivalAssessmentReport(ArchivalExecutionState state, string reason, ArchivalInputIdentity? input,
        ArchivalDeclarationAssessment declaration, ArchivalInputSupportAssessment support,
        IReadOnlyList<ArchivalClauseAssessment> clauses, ArchivalDocumentAssessment document,
        ArchivalUnassessedDimension software, ArchivalUnassessedDimension external, ArchivalCoverageReport? coverage)
    {
        ExecutionState = state; ExecutionReason = reason; Input = input; Declaration = declaration;
        InputSupport = support; Clauses = clauses; Document = document; SoftwareRequirements = software;
        ExternalRecognition = external; Coverage = coverage;
    }
    /// <summary>Serialization contract and assessment revision.</summary>
    public string AssessmentRevision => ArchivalAssessmentService.Revision;
    /// <summary>Clause ledger revision.</summary>
    public string LedgerRevision => ArchivalClauseCatalog.Revision;
    /// <summary>Strength/applicability matrix revision.</summary>
    public string MatrixRevision => ArchivalClauseMatrix.Revision;
    /// <summary>Execution dimension. Only Complete reports describe a finished inspection.</summary>
    public ArchivalExecutionState ExecutionState { get; }
    /// <summary>Stable reason for the execution state, such as <c>InspectionBudgetExceeded</c>.</summary>
    public string ExecutionReason { get; }
    /// <summary>Identity of the inspected bytes; null only when the input exceeded the byte budget before hashing.</summary>
    public ArchivalInputIdentity? Input { get; }
    /// <summary>Declaration dimension.</summary>
    public ArchivalDeclarationAssessment Declaration { get; }
    /// <summary>Input-support dimension.</summary>
    public ArchivalInputSupportAssessment InputSupport { get; }
    /// <summary>Clause dimension: 105 entries in ledger order.</summary>
    public IReadOnlyList<ArchivalClauseAssessment> Clauses { get; }
    /// <summary>Document dimension.</summary>
    public ArchivalDocumentAssessment Document { get; }
    /// <summary>Chapter 7 writer/reader requirements: not assessed from a single file.</summary>
    public ArchivalUnassessedDimension SoftwareRequirements { get; }
    /// <summary>Policy, PKI and independent approval: not authenticated.</summary>
    public ArchivalUnassessedDimension ExternalRecognition { get; }
    /// <summary>Unchanged underlying APF/AR/AC report, or null when inspection never started.</summary>
    public ArchivalCoverageReport? Coverage { get; }
    /// <summary>Returns the entry for an exact ledger clause identifier.</summary>
    /// <exception cref="ArgumentException">The clause is not in the ledger.</exception>
    public ArchivalClauseAssessment GetClause(string clause) =>
        Clauses.FirstOrDefault(x => x.Clause == clause) ?? throw new ArgumentException("Unknown ledger clause.", nameof(clause));
}
