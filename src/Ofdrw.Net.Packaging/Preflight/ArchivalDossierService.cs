using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>Evidence roles are caller declarations, never independently authenticated by this API.</summary>
public enum ArchivalEvidenceKind
{
    /// <summary>Declared positive control material.</summary>
    PositiveExample,
    /// <summary>Declared negative control material.</summary>
    NegativeExample,
    /// <summary>Declared uncertain control material.</summary>
    UncertainExample,
    /// <summary>Source-to-result transformation history.</summary>
    SourceTrace,
    /// <summary>Resource, font or media validation material.</summary>
    ResourceValidation,
    /// <summary>Conversion treatment and information-loss record.</summary>
    Conversion,
    /// <summary>Attachment classification and retention material.</summary>
    Attachment,
    /// <summary>Visual inspection material, with viewer and page scope.</summary>
    VisualReview,
    /// <summary>Independent reader observations.</summary>
    IndependentReader,
    /// <summary>Target archive policy material.</summary>
    ArchivalPolicy,
    /// <summary>Certificate, revocation or trusted-time material.</summary>
    Pki,
    /// <summary>Claimed independent approval; authority is not authenticated.</summary>
    IndependentApproval,
    /// <summary>Reference edition, amendments and applicability material.</summary>
    NormativeReference
}

/// <summary>A borrowed, stable evidence stream bound to one ledger clause and an expected digest.</summary>
public sealed class ArchivalEvidenceInput
{
    /// <summary>Declares evidence identity, role, provenance and the exact source/result pair it concerns.</summary>
    public ArchivalEvidenceInput(string id, string clause, ArchivalEvidenceKind kind, Stream content,
        string expectedSha256, string provenance, string sourceSha256, string resultSha256)
    {
        Id = Required(id); Clause = Required(clause); Provenance = Required(provenance);
        if (!Enum.IsDefined(typeof(ArchivalEvidenceKind), kind)) throw new ArgumentOutOfRangeException(nameof(kind));
        Kind = kind; Content = content ?? throw new ArgumentNullException(nameof(content));
        ExpectedSha256 = Digest(expectedSha256); SourceSha256 = Digest(sourceSha256); ResultSha256 = Digest(resultSha256);
    }
    /// <summary>Caller-supplied stable identity within this profile or dossier.</summary>
    public string Id { get; }
    /// <summary>Exact clause identifier from ArchivalClauseCatalog.</summary>
    public string Clause { get; }
    /// <summary>Caller-declared material role, not a validation verdict.</summary>
    public ArchivalEvidenceKind Kind { get; }
    /// <summary>Borrowed stable stream, read from zero and left open.</summary>
    public Stream Content { get; }
    /// <summary>Declared SHA-256 of the source artifact this evidence concerns.</summary>
    public string SourceSha256 { get; }
    /// <summary>Declared SHA-256 of the result artifact this evidence concerns.</summary>
    public string ResultSha256 { get; }
    /// <summary>Expected lowercase SHA-256 of the evidence bytes.</summary>
    public string ExpectedSha256 { get; }
    /// <summary>Caller-supplied material origin; never fetched or authenticated.</summary>
    public string Provenance { get; }
    private static string Digest(string value) => value != null && value.Length == 64 && value.All(Uri.IsHexDigit)
        ? value.ToLowerInvariant() : throw new ArgumentException("A SHA-256 hex digest is required.");
    internal static string Required(string value) => !string.IsNullOrWhiteSpace(value) ? value :
        throw new ArgumentException("A nonempty identity or provenance is required.");
}

/// <summary>Explicit target context. Version values must include amendments or an explanation of non-applicability.</summary>
public sealed class ArchivalDossierProfile
{
    /// <summary>Copies target context and explicit referenced editions; no edition is inferred.</summary>
    public ArchivalDossierProfile(string id, string? archivalAuthority, string scope,
        IEnumerable<KeyValuePair<string, string>> referencedEditions)
    {
        Id = ArchivalEvidenceInput.Required(id);
        ArchivalAuthority = string.IsNullOrWhiteSpace(archivalAuthority) ? null : archivalAuthority;
        Scope = ArchivalEvidenceInput.Required(scope);
        if (referencedEditions == null) throw new ArgumentNullException(nameof(referencedEditions));
        var editions = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var item in referencedEditions)
        {
            if (!ArchivalClauseCatalog.ReferencedStandards.Contains(item.Key))
                throw new ArgumentException("Unknown referenced standard: " + item.Key, nameof(referencedEditions));
            editions.Add(item.Key, ArchivalEvidenceInput.Required(item.Value));
        }
        ReferencedEditions = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(editions);
    }
    /// <summary>Caller-supplied stable identity within this profile or dossier.</summary>
    public string Id { get; }
    /// <summary>Declared target institution, or null when unspecified; not an authenticated approval.</summary>
    public string? ArchivalAuthority { get; }
    /// <summary>Declared archival scope and exclusions.</summary>
    public string Scope { get; }
    /// <summary>Immutable edition and amendment declarations keyed by referenced standard.</summary>
    public IReadOnlyDictionary<string, string> ReferencedEditions { get; }
}

/// <summary>Byte identity of a supplied artifact. This is not proof of its truth or approval authority.</summary>
public sealed class ArchivalArtifactIdentity
{
    internal ArchivalArtifactIdentity(string id, long size, string digest)
    { Id = id; Size = size; Sha256 = digest; }
    /// <summary>Caller-supplied stable identity within this profile or dossier.</summary>
    public string Id { get; }
    /// <summary>Number of bytes actually read.</summary>
    public long Size { get; }
    /// <summary>Lowercase SHA-256 of the bytes actually read.</summary>
    public string Sha256 { get; }
}

/// <summary>Collected evidence identity and declared role, without retaining payload bytes.</summary>
public sealed class ArchivalEvidenceRecord
{
    internal ArchivalEvidenceRecord(ArchivalEvidenceInput input, ArchivalArtifactIdentity identity, string sourceDigest, string resultDigest)
    { Id = input.Id; Clause = input.Clause; Kind = input.Kind; Provenance = input.Provenance;
      ExpectedSha256 = input.ExpectedSha256; Artifact = identity;
      SourceSha256 = input.SourceSha256; ResultSha256 = input.ResultSha256;
      SubjectMatches = SourceSha256 == sourceDigest && ResultSha256 == resultDigest; }
    /// <summary>Caller-supplied stable identity within this profile or dossier.</summary>
    public string Id { get; }
    /// <summary>Exact clause identifier from ArchivalClauseCatalog.</summary>
    public string Clause { get; }
    /// <summary>Caller-declared material role, not a validation verdict.</summary>
    public ArchivalEvidenceKind Kind { get; }
    /// <summary>Caller-supplied material origin; never fetched or authenticated.</summary>
    public string Provenance { get; }
    /// <summary>Expected lowercase SHA-256 of the evidence bytes.</summary>
    public string ExpectedSha256 { get; }
    /// <summary>Observed evidence byte identity.</summary>
    public ArchivalArtifactIdentity Artifact { get; }
    /// <summary>Declared SHA-256 of the source artifact this evidence concerns.</summary>
    public string SourceSha256 { get; }
    /// <summary>Declared SHA-256 of the result artifact this evidence concerns.</summary>
    public string ResultSha256 { get; }
    /// <summary>Whether declared source/result digests match this dossier pair.</summary>
    public bool SubjectMatches { get; }
    /// <summary>Whether observed bytes match the expected evidence digest.</summary>
    public bool DigestMatches => Artifact.Sha256 == ExpectedSha256;
    /// <summary>Always NotAuthenticated; hash matching does not validate authority.</summary>
    public string ApprovalStatus => "NotAuthenticated";
}

/// <summary>All obligations remain visible, including those with no object or supplied evidence.</summary>
public sealed class ArchivalClauseRecord
{
    internal ArchivalClauseRecord(ArchivalClauseDefinition clause, ArchivalCoverageReport report,
        IReadOnlyList<ArchivalEvidenceRecord> evidence)
    {
        Definition = clause;
        Evidence = Array.AsReadOnly(evidence.Where(x => x.Clause == clause.Clause).ToArray());
        SubsetOutcome = report.RuleResults.Where(x => x.Rule.Clause == clause.Clause).Select(x => x.Outcome)
            .Concat(report.FoundationalRules.RuleResults.Where(x => x.Rule.Clause == clause.Clause).Select(x => x.Outcome))
            .Aggregate(PreflightOutcome.NotApplicable, Merge);
    }
    internal static PreflightOutcome Merge(PreflightOutcome a, PreflightOutcome b) =>
        a == PreflightOutcome.Fail || b == PreflightOutcome.Fail ? PreflightOutcome.Fail :
        a == PreflightOutcome.Unknown || b == PreflightOutcome.Unknown ? PreflightOutcome.Unknown :
        a == PreflightOutcome.Pass || b == PreflightOutcome.Pass ? PreflightOutcome.Pass : PreflightOutcome.NotApplicable;
    /// <summary>Versioned standard obligation, mapped subset rules and remaining requirements.</summary>
    public ArchivalClauseDefinition Definition { get; }
    /// <summary>Existing mapped software conditions only. Unmapped clauses report Unknown below.</summary>
    public PreflightOutcome SubsetOutcome { get; }
    /// <summary>Immutable list of collected evidence records.</summary>
    public IReadOnlyList<ArchivalEvidenceRecord> Evidence { get; }
    /// <summary>Unvalidated external statements and subset Pass cannot close a full clause.</summary>
    public PreflightOutcome ClauseOutcome => SubsetOutcome == PreflightOutcome.Fail ? PreflightOutcome.Fail : PreflightOutcome.Unknown;
    /// <summary>Basic control roles lacking digest-matched, subject-matched material; not an exhaustive sufficiency check.</summary>
    public IReadOnlyList<ArchivalEvidenceKind> MissingControlEvidence => Array.AsReadOnly(new[]
    { ArchivalEvidenceKind.PositiveExample, ArchivalEvidenceKind.NegativeExample, ArchivalEvidenceKind.UncertainExample,
      ArchivalEvidenceKind.IndependentApproval }.Where(kind => !Evidence.Any(x => x.Kind == kind && x.DigestMatches && x.SubjectMatches && x.Artifact.Size > 0)).ToArray());
}

/// <summary>A source/result-bound evidence index, not an archival certificate.</summary>
public sealed class ArchivalDossierReport
{
    internal ArchivalDossierReport(ArchivalDossierProfile profile, ArchivalArtifactIdentity source,
        ArchivalArtifactIdentity result, ArchivalCoverageReport sourceCoverage, ArchivalCoverageReport resultCoverage,
        IReadOnlyList<ArchivalEvidenceRecord> evidence)
    {
        Profile = profile; Source = source; Result = result; SourceCoverage = sourceCoverage; ResultCoverage = resultCoverage;
        Evidence = evidence;
        Clauses = Array.AsReadOnly(ArchivalClauseCatalog.Clauses.Select(x => new ArchivalClauseRecord(x, resultCoverage, evidence)).ToArray());
        MissingReferencedStandards = Array.AsReadOnly(ArchivalClauseCatalog.ReferencedStandards
            .Where(x => !profile.ReferencedEditions.ContainsKey(x)).ToArray());
    }
    /// <summary>Serialization contract identifier.</summary>
    public string Schema => "ofdrw-archival-dossier/1";
    /// <summary>Revision of the complete obligation inventory.</summary>
    public string LedgerRevision => ArchivalClauseCatalog.Revision;
    /// <summary>Immutable declared archival target context.</summary>
    public ArchivalDossierProfile Profile { get; }
    /// <summary>Source package byte identity.</summary>
    public ArchivalArtifactIdentity Source { get; }
    /// <summary>Result package byte identity.</summary>
    public ArchivalArtifactIdentity Result { get; }
    /// <summary>Existing software checks on the source snapshot.</summary>
    public ArchivalCoverageReport SourceCoverage { get; }
    /// <summary>Existing software checks on the result snapshot.</summary>
    public ArchivalCoverageReport ResultCoverage { get; }
    /// <summary>Immutable list of collected evidence records.</summary>
    public IReadOnlyList<ArchivalEvidenceRecord> Evidence { get; }
    /// <summary>Every ledger obligation, including unimplemented and externally assessed items.</summary>
    public IReadOnlyList<ArchivalClauseRecord> Clauses { get; }
    /// <summary>Referenced standards without an explicit edition declaration.</summary>
    public IReadOnlyList<string> MissingReferencedStandards { get; }
    /// <summary>External baseline material roles not supplied with nonempty, matching bytes for this pair.
    /// This does not enumerate every clause-specific requirement and does not authenticate supplied material.</summary>
    public IReadOnlyList<ArchivalEvidenceKind> MissingExternalEvidence => Array.AsReadOnly(new[]
    { ArchivalEvidenceKind.NormativeReference, ArchivalEvidenceKind.ArchivalPolicy, ArchivalEvidenceKind.IndependentApproval }
        .Where(kind => !Evidence.Any(x => x.Kind == kind && x.DigestMatches && x.SubjectMatches && x.Artifact.Size > 0)).ToArray());
    /// <summary>Stable external assessment gaps. Declaring an edition or supplying a file cannot resolve these assessments.</summary>
    public IReadOnlyList<string> ExternalAssessmentGaps => Array.AsReadOnly(
        (Profile.ArchivalAuthority == null ? new[] { "ArchivalAuthorityUnspecified" } : Array.Empty<string>())
        .Concat(new[] { "ReferencedEditionsNotAssessed", "ArchivalPolicyNotAssessed", "IndependentApprovalNotAuthenticated" }).ToArray());
    /// <summary>Always Unknown: this service cannot independently assess archive acceptance or authenticate approvals.</summary>
    public PreflightOutcome ExternalAssessmentOutcome => PreflightOutcome.Unknown;
    /// <summary>Only artifact integrity. A matching digest does not approve the artifact's content.</summary>
    public PreflightOutcome EvidenceIntegrity => Evidence.Any(x => !x.DigestMatches || !x.SubjectMatches) ? PreflightOutcome.Fail :
        Evidence.Count == 0 ? PreflightOutcome.Unknown : PreflightOutcome.Pass;
    /// <summary>Result-only standard status: known Fail or Unknown; never full-standard Pass.</summary>
    public PreflightOutcome StandardOutcome => ResultCoverage.StandardOutcome;
    /// <summary>Always Incomplete until outstanding standard obligations have been independently resolved.</summary>
    public string CoverageStatus => "Incomplete";
    /// <summary>Always NotAssessed; no certification is performed.</summary>
    public string Certification => "NotAssessed";
}

/// <summary>Read-only, bounded evidence collection; never converts, deletes, fetches URLs or validates PKI.</summary>
public sealed class ArchivalDossierService
{
    /// <summary>Snapshots packages before hashing/checking. All borrowed streams start and finish at zero.
    /// Evidence is hashed without retaining its content. Budget/cancellation/I/O errors throw; no partial report is returned.
    /// Caller must supply stable streams and archive the original artifacts separately.</summary>
    public async Task<ArchivalDossierReport> CheckAsync(Stream source, Stream result, ArchivalDossierProfile profile,
        IEnumerable<ArchivalEvidenceInput> evidence, ArchivalPreflightOptions options,
        long maxTotalBytes = 1024L * 1024 * 1024, int maxEvidenceCount = 256, CancellationToken cancellationToken = default)
    {
        if (profile == null) throw new ArgumentNullException(nameof(profile));
        if (evidence == null) throw new ArgumentNullException(nameof(evidence));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (maxTotalBytes <= 0 || maxEvidenceCount < 0 || maxEvidenceCount > 10000) throw new ArgumentOutOfRangeException(nameof(maxTotalBytes));
        var snapshot = options.Snapshot();
        cancellationToken.ThrowIfCancellationRequested();
        var inputs = new List<ArchivalEvidenceInput>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in evidence)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (inputs.Count >= maxEvidenceCount) throw new InvalidDataException("Evidence count budget exceeded.");
            if (item == null || !ids.Add(item.Id) || !ArchivalClauseCatalog.Clauses.Any(x => x.Clause == item.Clause))
                throw new ArgumentException("Evidence needs unique IDs and recognized ledger clauses.", nameof(evidence));
            Validate(item.Content);
            inputs.Add(item);
        }
        Validate(source); Validate(result);
        var remaining = maxTotalBytes;
        using var sourceCopy = new MemoryStream();
        using var resultCopy = new MemoryStream();
        var sourceId = await Read(source, "source", sourceCopy, snapshot.LoadOptions.MaxInputBytes).ConfigureAwait(false);
        var resultId = await Read(result, "result", resultCopy, snapshot.LoadOptions.MaxInputBytes).ConfigureAwait(false);
        var records = new List<ArchivalEvidenceRecord>();
        foreach (var item in inputs)
            records.Add(new ArchivalEvidenceRecord(item, await Read(item.Content, item.Id, null, maxTotalBytes).ConfigureAwait(false), sourceId.Sha256, resultId.Sha256));
        var checker = new ArchivalCoverageService();
        var before = await checker.CheckAsync(sourceCopy, snapshot, cancellationToken).ConfigureAwait(false);
        var after = await checker.CheckAsync(resultCopy, snapshot, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return new ArchivalDossierReport(profile, sourceId, resultId, before, after, records.AsReadOnly());

        async Task<ArchivalArtifactIdentity> Read(Stream stream, string id, MemoryStream? copy, long limit)
        {
            long size = 0;
            var buffer = new byte[81920];
            using var hash = SHA256.Create();
            try
            {
                int count;
                while ((count = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) != 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (count > remaining || count > limit - size) throw new InvalidDataException("Dossier byte budget exceeded.");
                    remaining -= count; size += count;
                    hash.TransformBlock(buffer, 0, count, buffer, 0);
                    if (copy != null) copy.Write(buffer, 0, count);
                }
                hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                if (copy != null) copy.Position = 0;
                return new ArchivalArtifactIdentity(id, size, BitConverter.ToString(hash.Hash!).Replace("-", "").ToLowerInvariant());
            }
            finally { stream.Position = 0; }
        }
    }

    private static void Validate(Stream stream)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead || !stream.CanSeek || stream.Position != 0)
            throw new ArgumentException("Stable, readable, seekable streams at position zero are required.", nameof(stream));
    }
}
