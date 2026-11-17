using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Archival;
using Ofdrw.Net.Medical.Internal;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Medical;

/// <summary>Description of one package-owned executable rule.</summary>
public sealed class MedicalRuleDefinition
{
    internal MedicalRuleDefinition(string ruleId, IReadOnlyList<string> clauses, string description)
    { RuleId = ruleId; Clauses = clauses; Description = description; }
    /// <summary>Rule identity, <c>MH-001</c> to <c>MH-016</c>.</summary>
    public string RuleId { get; }
    /// <summary>Matrix clauses the rule observes.</summary>
    public IReadOnlyList<string> Clauses { get; }
    /// <summary>Bounded condition the rule checks.</summary>
    public string Description { get; }
}

/// <summary>
/// Read-only GB/T 48666-2026 (OFD-H) assessment. It runs the GB/T 42133 assessment of <c>Ofdrw.Net.Archival</c> on the
/// same bytes, applies the inheritance ledger, and adds the medical rules MH-001 to MH-016. It never writes, repairs,
/// verifies signatures cryptographically or reports a full-standard Pass.
/// </summary>
public sealed class MedicalAssessmentService
{
    /// <summary>Report contract and assessment logic revision.</summary>
    public const string Revision = "ofdrw-medical-assessment/1";

    /// <summary>The executable rules, in evaluation order.</summary>
    public static IReadOnlyList<MedicalRuleDefinition> Rules { get; } = Array.AsReadOnly(MedicalRules.All
        .Select(x => new MedicalRuleDefinition(x.Id, Array.AsReadOnly(x.Clauses), x.Description)).ToArray());

    /// <summary>
    /// Assesses a borrowed, stable, seekable input at position zero. The bytes are read once within
    /// <c>options.Archival.Preflight.LoadOptions.MaxInputBytes</c>, hashed, and every inspection uses that snapshot.
    /// The input is never written and is restored to position zero.
    /// </summary>
    /// <param name="input">Readable, seekable OFD stream at position zero.</param>
    /// <param name="options">Assessment options; null uses defaults.</param>
    /// <param name="cancellationToken">Cancellation throws <see cref="OperationCanceledException"/>.</param>
    /// <returns>The report; its <see cref="MedicalAssessmentReport.ExecutionState"/> records budget exhaustion or rejection.</returns>
    /// <exception cref="ArgumentException">The input violates the borrowing contract, or an option is invalid.</exception>
    public async Task<MedicalAssessmentReport> AssessAsync(Stream input, MedicalAssessmentOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (!input.CanRead || !input.CanSeek || input.Position != 0)
            throw new ArgumentException("A readable, seekable input at position zero is required.", nameof(input));
        options ??= new MedicalAssessmentOptions();
        if (options.Archival == null) throw new ArgumentException("Archival options are required.", nameof(options));
        if (options.MaxEmbeddedObjectBytes <= 0) throw new ArgumentException("MaxEmbeddedObjectBytes must be positive.", nameof(options));
        var mapping = options.Metadata?.Snapshot();
        var budgets = options.Archival.Preflight ?? throw new ArgumentException("Preflight options are required.", nameof(options));
        cancellationToken.ThrowIfCancellationRequested();

        byte[]? bytes;
        MedicalInputIdentity? identity;
        try { (bytes, identity) = await Snapshot(input, budgets.LoadOptions.MaxInputBytes, cancellationToken).ConfigureAwait(false); }
        finally { input.Position = 0; }
        if (bytes == null)
            return Build(MedicalExecutionState.BudgetExhausted, "InputBudgetExceeded", null, new List<MedicalRuleObservation>(), null, null);

        ArchivalAssessmentReport archival;
        using (var copy = new MemoryStream(bytes, false))
            archival = await new ArchivalAssessmentService().AssessAsync(copy, options.Archival, cancellationToken).ConfigureAwait(false);

        var observations = new List<MedicalRuleObservation>();
        string? rootDocType = null;
        var state = MedicalExecutionState.Complete;
        var reason = "Complete";
        try
        {
            using var package = new MedicalPackage(bytes, budgets, cancellationToken);
            var structure = MedicalStructure.Walk(package);
            rootDocType = (string?)structure.Root?.Root?.Attribute("DocType");
            var context = new MedicalRuleContext(package, structure, options, mapping, archival);
            foreach (var rule in MedicalRules.All)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var produced = rule.Check(context).ToList();
                if (rule.Clauses.Any(clause => produced.All(x => x.Clause != clause)) || produced.Any(x => x.RuleId != rule.Id || !rule.Clauses.Contains(x.Clause)))
                    throw new InvalidOperationException("Rule " + rule.Id + " returned an invalid observation set.");
                observations.AddRange(produced);
            }
        }
        catch (MedicalBudgetExceededException ex) { state = MedicalExecutionState.BudgetExhausted; reason = ex.Budget + "BudgetExceeded"; }
        catch (MedicalInputRejectedException ex) { state = MedicalExecutionState.Error; reason = ex.Message; }
        catch (InvalidDataException) { state = MedicalExecutionState.Error; reason = "MalformedZip"; }
        return Build(state, reason, identity, observations, archival, rootDocType);
    }

    private static async Task<(byte[]?, MedicalInputIdentity?)> Snapshot(Stream input, long limit, CancellationToken token)
    {
        if (input.Length > limit) return (null, null);
        using var copy = new MemoryStream();
        using var hash = SHA256.Create();
        var buffer = new byte[81920];
        int count;
        while ((count = await input.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false)) != 0)
        {
            if (count > limit - copy.Length) return (null, null);
            hash.TransformBlock(buffer, 0, count, buffer, 0);
            copy.Write(buffer, 0, count);
        }
        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return (copy.ToArray(), new MedicalInputIdentity(copy.Length, BitConverter.ToString(hash.Hash!).Replace("-", string.Empty).ToLowerInvariant()));
    }

    private static MedicalAssessmentReport Build(MedicalExecutionState state, string reason, MedicalInputIdentity? identity,
        List<MedicalRuleObservation> observations, ArchivalAssessmentReport? archival, string? rawDocType)
    {
        var complete = state == MedicalExecutionState.Complete;
        if (!complete) observations.Clear();
        var clauses = MedicalClauseMatrix.Entries.Select(entry =>
        {
            var mine = observations.Where(x => x.Clause == entry.Clause).ToList().AsReadOnly();
            var advisory = entry.Strength != MedicalClauseStrength.Shall && entry.Strength != MedicalClauseStrength.Statement;
            var advisories = advisory ? mine.Where(x => x.Outcome == PreflightOutcome.Fail).ToList().AsReadOnly() : new List<MedicalRuleObservation>().AsReadOnly();
            return new MedicalClauseAssessment(entry, mine, advisories, Aggregate(mine, advisory));
        }).ToList().AsReadOnly();

        var inherited = MedicalArchivalInheritance.Entries.Select(entry =>
        {
            var outcome = archival != null && archival.ExecutionState == ArchivalExecutionState.Complete
                ? archival.GetClause(entry.ArchivalClause).Outcome : PreflightOutcome.Unknown;
            return new MedicalInheritedArchivalClause(entry, outcome);
        }).ToList().AsReadOnly();

        var gaps = new List<string>
        {
            "NoFullStandardAssessment", "ReferencedStandardsNotAssessed:GB/T 33190,GB/T 33560,GB/T 35275,GB/T 38540,GB 18030",
            "CryptographicVerificationNotPerformed", "ProcessAndSoftwareRequirementsNotAssessed", "MachineReadableSchemaUndefined",
            "ArchivalInheritanceUnverified", "ResourceTableClauseConflict:6.2(b),7.2(a)"
        };
        if (!complete) gaps.Add("AssessmentIncomplete:" + reason);
        if (archival == null || archival.ExecutionState != ArchivalExecutionState.Complete) gaps.Add("ArchivalAssessmentIncomplete");
        var failed = clauses.Any(x => x.Outcome == PreflightOutcome.Fail) || inherited.Any(x => x.ContributesFail);
        var document = new MedicalDocumentAssessment(
            clauses.Where(x => x.HasExecutableCheck).Select(x => x.Clause).ToList().AsReadOnly(),
            clauses.Where(x => !x.HasExecutableCheck).Select(x => x.Clause).ToList().AsReadOnly(),
            gaps.AsReadOnly(), failed ? PreflightOutcome.Fail : PreflightOutcome.Unknown);

        var declarationObservation = observations.FirstOrDefault(x => x.RuleId == "MH-001");
        var status = declarationObservation == null ? MedicalDeclarationStatus.Unavailable : declarationObservation.ReasonCode switch
        {
            "MedicalDocType" => MedicalDeclarationStatus.Medical,
            "TypesetDashVariant" => MedicalDeclarationStatus.TypesetVariant,
            "NonMedicalDocType" => MedicalDeclarationStatus.NotMedical,
            "DocTypeMissing" => MedicalDeclarationStatus.Missing,
            "AmbiguousDocType" => MedicalDeclarationStatus.Ambiguous,
            _ => MedicalDeclarationStatus.Unavailable
        };
        return new MedicalAssessmentReport(state, reason, identity,
            new MedicalDeclarationAssessment(status, complete ? rawDocType : null, declarationObservation),
            clauses, inherited, document, archival);
    }

    /// <summary>Shall/Statement: Fail dominates Unknown, Unknown dominates Pass. Advisory rows never Fail.</summary>
    private static PreflightOutcome Aggregate(IReadOnlyList<MedicalRuleObservation> observations, bool advisory)
    {
        if (observations.Count == 0) return PreflightOutcome.Unknown;
        if (observations.Any(x => x.Outcome == PreflightOutcome.Fail)) return advisory ? PreflightOutcome.Unknown : PreflightOutcome.Fail;
        if (observations.Any(x => x.Outcome == PreflightOutcome.Unknown)) return PreflightOutcome.Unknown;
        if (observations.All(x => x.Outcome == PreflightOutcome.NotApplicable)) return PreflightOutcome.NotApplicable;
        return PreflightOutcome.Pass;
    }
}
