using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Archival.Internal;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival;

/// <summary>
/// Read-only GB/T 42133-2022 assessment that composes the existing Preflight services and reports declaration,
/// input support, clauses, document coverage, software requirements, external recognition and execution separately.
/// It never writes, repairs or labels OFD-A, and never reports a full-standard Pass.
/// </summary>
public sealed class ArchivalAssessmentService
{
    /// <summary>Report contract and assessment logic revision.</summary>
    public const string Revision = "ofdrw-archival-assessment/1";

    private static readonly string[] SoftwareClauses = { "7.1(1)", "7.1(2)", "7.2(1)", "7.2(2)" };
    private static readonly string[] AnnotationOptionAClauses = { "6.10.1(a)(2)", "6.10.1(a)(3)" };
    private readonly IReadOnlyList<IArchivalClauseCheck> checks;

    /// <summary>Creates a service with the package's registered clause checks.</summary>
    public ArchivalAssessmentService() : this(ArchivalOwnedRules.Registered) { }

    internal ArchivalAssessmentService(IEnumerable<IArchivalClauseCheck> checks)
    { this.checks = ArchivalOwnedRules.Validate(checks); }

    /// <summary>
    /// Assesses a borrowed, stable, seekable input at position zero. The bytes are read once within
    /// <c>LoadOptions.MaxInputBytes</c>, hashed, and every inspection runs on that same snapshot. The input is never
    /// written, is left open and is restored to position zero on every exit.
    /// </summary>
    /// <param name="input">Readable, seekable OFD stream at position zero.</param>
    /// <param name="options">Profile and Preflight budgets; copied before reading.</param>
    /// <param name="cancellationToken">Cancellation throws <see cref="OperationCanceledException"/>, as the Preflight services do; no partial report is returned.</param>
    /// <returns>A report whose <see cref="ArchivalAssessmentReport.ExecutionState"/> records budget exhaustion or input rejection.</returns>
    /// <exception cref="ArgumentException">The input or options violate the borrowing contract.</exception>
    /// <exception cref="IOException">Reading the input or restoring its position fails.</exception>
    public async Task<ArchivalAssessmentReport> AssessAsync(Stream input, ArchivalAssessmentOptions options,
        CancellationToken cancellationToken = default)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (!input.CanRead || !input.CanSeek || input.Position != 0)
            throw new ArgumentException("A readable, seekable input at position zero is required.", nameof(input));
        var profile = options.ResolveProfile();
        var preflight = options.SnapshotPreflight();
        var fonts = options.SnapshotFonts();
        cancellationToken.ThrowIfCancellationRequested();

        byte[]? bytes;
        ArchivalInputIdentity? identity;
        try { (bytes, identity) = await Snapshot(input, preflight.LoadOptions.MaxInputBytes, cancellationToken).ConfigureAwait(false); }
        finally { input.Position = 0; }
        if (bytes == null)
            return Build(profile, ArchivalExecutionState.BudgetExhausted, "InputBudgetExceeded", null, null,
                new ArchivalScopeResult(), Array.Empty<ArchivalClauseObservation>(), new HashSet<string>());

        using var copy = new MemoryStream(bytes, false);
        var coverage = await new ArchivalCoverageService().CheckAsync(copy, preflight, cancellationToken).ConfigureAwait(false);
        copy.Position = 0;
        OfdPackageArchive? archive = null;
        string? loadFailure = null;
        try { archive = await new OfdPackageLoader().LoadAsync(copy, preflight.LoadOptions, cancellationToken).ConfigureAwait(false); }
        catch (InvalidDataException ex)
        {
            // The public loader reports configured limits and malformed ZIPs with the same exception type; limit
            // failures carry an internal marker so the classification does not depend on message text.
            loadFailure = OfdPackageLoader.IsLimitFailure(ex) ? "InputBudgetExceeded" : "InputRejected";
        }

        var scope = new ArchivalScopeResult();
        var owned = new List<ArchivalClauseObservation>();
        var decided = new HashSet<string>(StringComparer.Ordinal);
        // Several checks may share one gap (AC-017..AC-020 share FontPayloadNotAssessed): it closes only when all of them decided.
        var undecided = new HashSet<string>(StringComparer.Ordinal);
        if (archive != null)
        {
            var xml = new BoundedPackageXml(archive, preflight, cancellationToken);
            scope = new ArchivalScopeInspector(archive, xml).Inspect();
            var context = new ArchivalCheckContext(archive, xml, preflight, scope, cancellationToken, fonts);
            foreach (var check in checks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var produced = (check.Check(context) ?? Enumerable.Empty<ArchivalClauseObservation>()).ToList();
                    if (produced.Count == 0 || produced.Any(x => x.RuleId != check.Rule.RuleId || !check.Rule.Clauses.Contains(x.Clause)))
                        throw new InvalidOperationException("Clause check " + check.Rule.RuleId + " returned an invalid observation set.");
                    owned.AddRange(produced);
                    if (check.GapCode != null) (check.Decided(produced.AsReadOnly()) ? decided : undecided).Add(check.GapCode);
                }
                catch (PackageBudgetExceededException ex)
                {
                    scope.BudgetExceeded ??= ex.Budget;
                    if (check.GapCode != null) undecided.Add(check.GapCode);
                    owned.Add(ArchivalOwnedRules.Observation(check.Rule, PreflightOutcome.Unknown, "InspectionBudgetExceeded",
                        "<input>", "<package>", ex.Budget + " budget"));
                }
            }
        }
        else
            foreach (var check in checks)
                owned.Add(ArchivalOwnedRules.Observation(check.Rule, PreflightOutcome.Unknown, loadFailure ?? "InputRejected",
                    "<input>", "<package>", "Package not loaded"));

        var local = coverage.FoundationalRules.LocalPreflight;
        var (state, reason) =
            loadFailure != null ? (loadFailure == "InputBudgetExceeded" ? ArchivalExecutionState.BudgetExhausted : ArchivalExecutionState.Error, loadFailure) :
            local.RunState != PreflightRunState.Complete
                ? (local.CoverageGaps.Contains("InspectionBudgetExceeded") ? (ArchivalExecutionState.BudgetExhausted, "InspectionBudgetExceeded") : (ArchivalExecutionState.Error, "InputRejected")) :
            scope.BudgetExceeded != null ? (ArchivalExecutionState.BudgetExhausted, "ScopeInspectionBudgetExceeded") :
            (ArchivalExecutionState.Complete, "Complete");
        // A gap closes only for a complete run whose checks actually decided the question for this input.
        if (state != ArchivalExecutionState.Complete) decided.Clear();
        decided.ExceptWith(undecided);
        return Build(profile, state, reason, identity, coverage, scope, owned, decided);
    }

    private static async Task<(byte[]?, ArchivalInputIdentity?)> Snapshot(Stream input, long limit, CancellationToken token)
    {
        if (input.Length > limit) return (null, null);
        using var copy = new MemoryStream();
        using var hash = SHA256.Create();
        var buffer = new byte[81920];
        int count;
        while ((count = await input.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false)) != 0)
        {
            token.ThrowIfCancellationRequested();
            if (count > limit - copy.Length) return (null, null);
            hash.TransformBlock(buffer, 0, count, buffer, 0);
            copy.Write(buffer, 0, count);
        }
        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        var digest = BitConverter.ToString(hash.Hash!).Replace("-", string.Empty).ToLowerInvariant();
        return (copy.ToArray(), new ArchivalInputIdentity(copy.Length, digest));
    }

    private static ArchivalAssessmentReport Build(ArchivalInputProfile profile, ArchivalExecutionState state, string reason,
        ArchivalInputIdentity? identity, ArchivalCoverageReport? coverage, ArchivalScopeResult scope,
        IReadOnlyList<ArchivalClauseObservation> owned, ISet<string> decidedGaps)
    {
        var complete = state == ArchivalExecutionState.Complete;
        var declaration = Declaration(coverage, scope);

        var findings = scope.Findings.ToList();
        if (!complete && findings.All(x => x.Code != "AssessmentIncomplete"))
            findings.Add(new ArchivalScopeFinding("AssessmentIncomplete", ArchivalInputSupport.Unknown,
                new ArchivalEvidenceLocation("<input>", "<package>", reason, "Complete bounded inspection")));
        var support = findings.Any(x => x.Support == ArchivalInputSupport.OutOfScope) ? ArchivalInputSupport.OutOfScope :
            findings.Count > 0 ? ArchivalInputSupport.Unknown : ArchivalInputSupport.Supported;
        var inputSupport = new ArchivalInputSupportAssessment(profile, support, findings.AsReadOnly());

        var mapped = new List<ArchivalClauseObservation>();
        if (coverage != null)
        {
            mapped.AddRange(coverage.FoundationalRules.RuleResults.Select(x => new ArchivalClauseObservation(x.Rule.RuleId,
                ArchivalOwnedRules.PackagingOwner, x.Rule.RuleVersion, x.Rule.Clause, x.Outcome, x.ReasonCode,
                ArchivalEvidenceLocation.From(x.Evidence))));
            mapped.AddRange(coverage.RuleResults.Select(x => new ArchivalClauseObservation(x.Rule.RuleId,
                ArchivalOwnedRules.PackagingOwner, x.Rule.RuleVersion, x.Rule.Clause, x.Outcome, x.ReasonCode,
                ArchivalEvidenceLocation.From(x.Evidence))));
        }
        mapped.AddRange(owned);

        var clauses = ArchivalClauseCatalog.Clauses.Select(definition =>
        {
            var matrix = ArchivalClauseMatrix.Get(definition.Clause);
            var observations = mapped.Where(x => x.Clause == definition.Clause).ToList().AsReadOnly();
            var notes = new List<string>();
            if (AnnotationOptionAClauses.Contains(definition.Clause))
                notes.Add("AnnotationObjectRetentionAssumed: the mapped rule applies only when annotations are retained as annotation objects (6.10.1 option a), not when appearances are merged into pages (option b).");
            if (matrix.Confidence == ArchivalMatrixConfidence.Unverified)
                notes.Add("ModalScopeUnverified: the scope of the modal keyword admits more than one reading.");
            if (!matrix.EditionFixed && matrix.ReferencedStandards.Count > 0)
                notes.Add("ReferencedEditionNotFixed: " + string.Join(", ", matrix.ReferencedStandards) + ".");
            return new ArchivalClauseAssessment(definition, matrix, observations,
                ArchivalClauseAggregation.Advisories(matrix.Strength, observations),
                ArchivalClauseAggregation.Outcome(matrix.Strength, observations, complete), notes.AsReadOnly());
        }).ToList().AsReadOnly();

        // CustomTagClosureNotAssessed and FontPayloadNotAssessed are removed only where the owned checks decided them.
        var gaps = new List<string>
        {
            "NoFullStandardAssessment", "ReferencedEditionsNotFixed", "MatrixScopeUnverified:6.11,6.13.1(a),6.13.1(b),6.13.1(c)",
            "LocalPreflightRulesNotMappedToClauses", "CustomTagClosureNotAssessed", "FontPayloadNotAssessed"
        }.Where(x => !decidedGaps.Contains(x)).ToList();
        if (support == ArchivalInputSupport.OutOfScope) gaps.Add("InputOutsideProfile");
        if (support == ArchivalInputSupport.Unknown) gaps.Add("InputScopeUnknown");
        if (!complete) gaps.Add("AssessmentIncomplete:" + reason);
        var document = new ArchivalDocumentAssessment(
            clauses.Where(x => x.HasExecutableCheck).Select(x => x.Clause).ToList().AsReadOnly(),
            clauses.Where(x => !x.HasExecutableCheck).Select(x => x.Clause).ToList().AsReadOnly(),
            gaps.AsReadOnly(),
            clauses.Any(x => x.Outcome == PreflightOutcome.Fail) ? PreflightOutcome.Fail : PreflightOutcome.Unknown);

        var software = new ArchivalUnassessedDimension("NotAssessed", Array.AsReadOnly(SoftwareClauses), Array.AsReadOnly(new[]
        {
            "WriterCapabilityNotAssessed", "WriterProcessingNotAssessed",
            "ReaderOutputAndUnprocessedItemReportingNotAssessed", "ReaderEditProtectionNotAssessed"
        }));
        var external = new ArchivalUnassessedDimension("NotAuthenticated",
            ArchivalClauseMatrix.Entries.Where(x => x.FirstProfile == ArchivalProfileApplicability.ExternalEvidence &&
                !SoftwareClauses.Contains(x.Clause)).Select(x => x.Clause).ToList().AsReadOnly(),
            Array.AsReadOnly(new[]
            {
                "ArchivalAuthorityUnspecified", "ArchivalPolicyNotAssessed", "ReferencedEditionsNotAssessed",
                "PkiNotAssessed", "IndependentApprovalNotAuthenticated"
            }));
        return new ArchivalAssessmentReport(state, reason, identity, declaration, inputSupport, clauses, document,
            software, external, coverage);
    }

    private static ArchivalDeclarationAssessment Declaration(ArchivalCoverageReport? coverage, ArchivalScopeResult scope)
    {
        var rule = coverage?.RuleResults.FirstOrDefault(x => x.Rule.RuleId == "AC-001");
        if (rule == null)
            return new ArchivalDeclarationAssessment(ArchivalDeclarationStatus.Unavailable, scope.RawDocType, scope.ForeignDocType,
                PreflightOutcome.Unknown, "InspectionNotStarted",
                new ArchivalEvidenceLocation("<input>", "<package>", "Inspection not started", "Recognized OFD root declares DocType=OFD-A"));
        var status = rule.Outcome == PreflightOutcome.Pass ? ArchivalDeclarationStatus.Archival :
            rule.ReasonCode == "DocTypeMissing" ? ArchivalDeclarationStatus.Missing :
            rule.ReasonCode == "NonArchivalDocType" ? ArchivalDeclarationStatus.NotArchival :
            rule.ReasonCode == "AmbiguousDocType" ? ArchivalDeclarationStatus.Ambiguous : ArchivalDeclarationStatus.Unavailable;
        return new ArchivalDeclarationAssessment(status, scope.RawDocType, scope.ForeignDocType, rule.Outcome, rule.ReasonCode,
            ArchivalEvidenceLocation.From(rule.Evidence));
    }
}
