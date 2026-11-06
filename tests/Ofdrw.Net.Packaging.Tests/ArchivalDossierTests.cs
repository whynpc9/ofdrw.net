using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Packaging.Tests;

public class ArchivalDossierTests
{
    private static MemoryStream Package(string docType = "OFD-A", string extra = "")
    {
        var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
        {
            Write("OFD.xml", $"<OFD xmlns='http://www.ofdspec.org' DocType='{docType}'><DocBody><DocRoot>Doc/Document.xml</DocRoot></DocBody></OFD>");
            Write("Doc/Document.xml", $"<Document xmlns='http://www.ofdspec.org'><Pages><Page ID='1' BaseLoc='Page.xml'/></Pages>{extra}</Document>");
            Write("Doc/Page.xml", "<Page xmlns='http://www.ofdspec.org'><Content><Layer><TextObject Size='3'/></Layer></Content></Page>");
            void Write(string name, string content)
            { using var stream = zip.CreateEntry(name).Open(); stream.Write(Encoding.UTF8.GetBytes(content)); }
        }
        output.Position = 0;
        return output;
    }
    private static ArchivalDossierProfile Profile() => new("test-profile/1", "Test archive (not approved)",
        "Unsigned single-document fixture", Array.Empty<KeyValuePair<string, string>>());
    private static string Hash(byte[] value) => Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();

    [Fact]
    public async Task EveryLedgerObligationIsPresentAndMissingEvidenceNeverPasses()
    {
        using var source = Package(extra: "<Permissions Edit='false'/>");
        using var result = Package();
        var sourceBytes = source.ToArray(); var resultBytes = result.ToArray();
        var report = await new ArchivalDossierService().CheckAsync(source, result, Profile(), [], new());
        Assert.Equal(105, report.Clauses.Count);
        Assert.Equal(105, report.Clauses.Select(x => x.Definition.Clause).Distinct().Count());
        Assert.Equal(13, report.Clauses.Count(x => x.Definition.Clause.StartsWith("A.1/")));
        Assert.Equal(17, report.Clauses.Sum(x => x.Definition.RuleIds.Count));
        Assert.Equal(2, report.Clauses.Count(x => x.Definition.Handler != null));
        Assert.All(report.Clauses, x => { Assert.NotEqual(PreflightOutcome.Pass, x.ClauseOutcome); Assert.Equal(4, x.MissingControlEvidence.Count); });
        Assert.Equal(10, report.MissingReferencedStandards.Count);
        Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
        Assert.Equal(PreflightOutcome.Unknown, report.EvidenceIntegrity);
        Assert.Equal(PreflightOutcome.Fail, report.SourceCoverage.StandardOutcome);
        Assert.Equal(Hash(sourceBytes), report.Source.Sha256);
        Assert.Equal(Hash(resultBytes), report.Result.Sha256);
        Assert.Equal(resultBytes.Length, report.Result.Size);
        Assert.Equal(0, source.Position); Assert.Equal(0, result.Position);
        Assert.Equal(sourceBytes, source.ToArray()); Assert.Equal(resultBytes, result.ToArray());
        var json = JsonSerializer.Serialize(report);
        Assert.Contains("NotAssessed", json);
        Assert.Contains("ofdrw-archival-dossier/1", json);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task MatchingOrForgedApprovalCannotPromoteCompliance(bool validHash)
    {
        using var source = Package(); using var result = Package();
        using var statement = new MemoryStream(Encoding.UTF8.GetBytes("claimed independent approval"));
        var input = new ArchivalEvidenceInput("approval", "6.2.1(a)", ArchivalEvidenceKind.IndependentApproval,
            statement, validHash ? Hash(statement.ToArray()) : new string('0', 64), "test authority claim", Hash(source.ToArray()), Hash(result.ToArray()));
        var report = await new ArchivalDossierService().CheckAsync(source, result, Profile(), [input], new());
        var clause = Assert.Single(report.Clauses, x => x.Definition.Clause == "6.2.1(a)");
        Assert.Equal(PreflightOutcome.Pass, clause.SubsetOutcome);
        Assert.Equal(PreflightOutcome.Unknown, clause.ClauseOutcome);
        Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
        Assert.Equal(validHash ? PreflightOutcome.Pass : PreflightOutcome.Fail, report.EvidenceIntegrity);
        Assert.Equal(validHash ? 3 : 4, clause.MissingControlEvidence.Count);
        Assert.Equal("NotAuthenticated", Assert.Single(report.Evidence).ApprovalStatus);
        Assert.True(Assert.Single(report.Evidence).SubjectMatches);
        Assert.Equal(0, statement.Position);
    }

    [Fact]
    public async Task EvidenceFromAnotherSourceCannotFillCurrentControls()
    {
        using var source = Package(); using var result = Package(); using var statement = new MemoryStream(new byte[] { 1 });
        var e = new ArchivalEvidenceInput("e", "6.14", ArchivalEvidenceKind.PositiveExample, statement,
            Hash(statement.ToArray()), "stale-source", new string('0', 64), Hash(result.ToArray()));
        var report = await new ArchivalDossierService().CheckAsync(source, result, Profile(), [e], new());
        Assert.False(Assert.Single(report.Evidence).SubjectMatches);
        Assert.Equal(PreflightOutcome.Fail, report.EvidenceIntegrity);
        Assert.Equal(4, Assert.Single(report.Clauses, x => x.Definition.Clause == "6.14").MissingControlEvidence.Count);
        Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
    }

    [Fact]
    public async Task ResultFailureIsNotHiddenByMissingExternalMaterial()
    {
        using var source = Package(); using var result = Package("OFD");
        var report = await new ArchivalDossierService().CheckAsync(source, result, Profile(), [], new());
        Assert.Equal(PreflightOutcome.Fail, report.StandardOutcome);
        Assert.Equal(PreflightOutcome.Fail, Assert.Single(report.Clauses, x => x.Definition.Clause == "6.2.1(a)").ClauseOutcome);
    }

    [Theory]
    [InlineData(1)] // source alone exceeds the cumulative budget
    [InlineData(2)] // source+result exceeds it
    [InlineData(3)] // evidence exceeds remaining budget
    public async Task CumulativeBudgetThrowsWithoutPartialReportAndRestoresStreams(int stage)
    {
        using var source = Package(); using var result = Package();
        using var content = new MemoryStream(new byte[400]);
        var input = new ArchivalEvidenceInput("e", "6.14", ArchivalEvidenceKind.SourceTrace, content, Hash(content.ToArray()), "fixture", Hash(source.ToArray()), Hash(result.ToArray()));
        var budget = stage == 1 ? 1 : stage == 2 ? source.Length + 1 : source.Length + result.Length + 399;
        await Assert.ThrowsAsync<InvalidDataException>(() => new ArchivalDossierService().CheckAsync(source, result, Profile(), [input], new(), budget));
        Assert.Equal(0, source.Position); Assert.Equal(0, result.Position); Assert.Equal(0, content.Position);
        Assert.True(source.CanRead); Assert.True(result.CanRead); Assert.True(content.CanRead);
    }

    [Fact]
    public async Task CountDuplicateAndUnknownClauseAreRejectedBeforeReadingPackages()
    {
        using var source = Package(); using var result = Package(); using var content = new MemoryStream();
        var e = new ArchivalEvidenceInput("e", "6.14", ArchivalEvidenceKind.SourceTrace, content, Hash([]), "fixture", Hash(source.ToArray()), Hash(result.ToArray()));
        var service = new ArchivalDossierService();
        await Assert.ThrowsAsync<InvalidDataException>(() => service.CheckAsync(source, result, Profile(), [e], new(), maxEvidenceCount: 0));
        await Assert.ThrowsAsync<ArgumentException>(() => service.CheckAsync(source, result, Profile(), [e, e], new()));
        var unknown = new ArchivalEvidenceInput("e", "imaginary", ArchivalEvidenceKind.SourceTrace, content, Hash([]), "fixture", Hash(source.ToArray()), Hash(result.ToArray()));
        await Assert.ThrowsAsync<ArgumentException>(() => service.CheckAsync(source, result, Profile(), [unknown], new()));
        Assert.Equal(0, source.Position); Assert.Equal(0, result.Position);
    }

    [Fact]
    public async Task CancellationDuringEvidenceReadRestoresBorrowedStream()
    {
        using var cts = new CancellationTokenSource();
        using var source = Package(); using var result = Package(); using var content = new CancelStream(cts);
        var e = new ArchivalEvidenceInput("e", "6.14", ArchivalEvidenceKind.SourceTrace, content, Hash([1]), "fixture", Hash(source.ToArray()), Hash(result.ToArray()));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new ArchivalDossierService().CheckAsync(source, result, Profile(), [e], new(), cancellationToken: cts.Token));
        Assert.Equal(0, source.Position); Assert.Equal(0, result.Position); Assert.Equal(0, content.Position);
    }
    private sealed class CancelStream(CancellationTokenSource cancellation) : MemoryStream(new byte[] { 1 })
    {
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        { cancellation.Cancel(); return base.ReadAsync(buffer, offset, count, CancellationToken.None); }
    }

    [Fact]
    public void CatalogMatchesPublishedLedgerIncludingAnnexAndAllMappedRules()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !File.Exists(Path.Combine(root.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        Assert.NotNull(root);
        var text = File.ReadAllText(Path.Combine(root.FullName, "docs/content/reference/archival-standard-coverage.md"));
        var rows = text.Split("## 结果与转换边界")[0].Split('\n')
            .Where(x => x.StartsWith("| ")).Select(x => x.Trim('|').Split('|').Select(v => v.Trim()).ToArray())
            .Where(x => x.Length == 3 && (System.Text.RegularExpressions.Regex.IsMatch(x[0], @"^[567]\.") ||
                new[] { "6.7", "6.8", "6.11", "6.14", "6.16" }.Contains(x[0]) || int.TryParse(x[0], out _)))
            .Select(x => (Clause: int.TryParse(x[0], out _) ? "A.1/" + x[0] : x[0], Condition: x[1], Gap: x[2])).ToArray();
        Assert.Equal(rows, ArchivalClauseCatalog.Clauses.Select(x => (x.Clause, x.Condition, x.RemainingEvidence)).ToArray());
        var mapped = ArchivalClauseCatalog.Clauses.SelectMany(x => x.RuleIds).Order().ToArray();
        Assert.Equal(ArchivalRuleCatalog.Rules.Select(x => x.RuleId).Concat(ArchivalCoverageRuleCatalog.Rules.Select(x => x.RuleId)).Order(), mapped);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public async Task UnspecifiedArchiveIsAnExplicitGapWithoutInventingAnInstitution(string? authority)
    {
        using var source = Package(); using var result = Package();
        var profile = new ArchivalDossierProfile("official-standard-only/1", authority, "Software checks only", []);
        var report = await new ArchivalDossierService().CheckAsync(source, result, profile, [], new());
        Assert.Null(report.Profile.ArchivalAuthority);
        Assert.Contains("ArchivalAuthorityUnspecified", report.ExternalAssessmentGaps);
        Assert.Equal(3, report.MissingExternalEvidence.Count);
        Assert.Equal(PreflightOutcome.Unknown, report.ExternalAssessmentOutcome);
        Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
    }

    [Fact]
    public async Task AllExternalRolesAndEditionDeclarationsStillRequireIndependentAssessment()
    {
        using var source = Package(); using var result = Package();
        using var statement = new MemoryStream(new byte[] { 1 });
        var profile = new ArchivalDossierProfile("claims/1", "Claimed authority", "Claims only",
            ArchivalClauseCatalog.ReferencedStandards.Select(x => KeyValuePair.Create(x, "caller edition declaration")));
        var roles = new[] { ArchivalEvidenceKind.NormativeReference, ArchivalEvidenceKind.ArchivalPolicy, ArchivalEvidenceKind.IndependentApproval };
        var evidence = roles.Select(kind => new ArchivalEvidenceInput(kind.ToString(), "5.1(a)", kind, statement,
            Hash(statement.ToArray()), "unverified claim", Hash(source.ToArray()), Hash(result.ToArray()))).ToArray();
        var service = new ArchivalDossierService();
        var report = await service.CheckAsync(source, result, profile, evidence, new());
        Assert.Empty(report.MissingExternalEvidence);
        Assert.Empty(report.MissingReferencedStandards);
        Assert.DoesNotContain("ArchivalAuthorityUnspecified", report.ExternalAssessmentGaps);
        Assert.Equal(3, report.ExternalAssessmentGaps.Count);
        Assert.Equal(PreflightOutcome.Pass, report.EvidenceIntegrity);
        Assert.Equal(PreflightOutcome.Unknown, report.ExternalAssessmentOutcome);
        Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
        statement.SetLength(0);
        var empty = new ArchivalEvidenceInput("empty-policy", "5.1(a)", ArchivalEvidenceKind.ArchivalPolicy, statement,
            Hash([]), "empty claim", Hash(source.ToArray()), Hash(result.ToArray()));
        var emptyReport = await service.CheckAsync(source, result, profile, [empty], new());
        Assert.Contains(ArchivalEvidenceKind.ArchivalPolicy, emptyReport.MissingExternalEvidence);
    }

    [Fact]
    public void ProfileCopiesEditionsAndRejectsSilentUnknownStandard()
    {
        var editions = new Dictionary<string, string> { ["GB/T 33190"] = "2016; amendments not evaluated" };
        var profile = new ArchivalDossierProfile("p", "a", "s", editions);
        editions.Clear();
        Assert.Single(profile.ReferencedEditions);
        Assert.Throws<ArgumentException>(() => new ArchivalDossierProfile("p", "a", "s", new Dictionary<string, string> { ["unknown"] = "1" }));
    }
}
