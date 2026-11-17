using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Ofdrw.Net.Archival.Internal;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Archival.Tests;

public class ArchivalAssessmentServiceTests
{
    private const string Ns = "http://www.ofdspec.org";
    private static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    // Same in-memory package shape as the Packaging archival tests.
    private static Dictionary<string, byte[]> Seed(string docType = "DocType='OFD-A'", string bodyExtra = "", string documentExtra = "") =>
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' {docType}><DocBody><DocRoot>Doc_0/Document.xml</DocRoot>{bodyExtra}</DocBody></OFD>"),
            ["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'><Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages>{documentExtra}</Document>"),
            ["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer><TextObject Size='3'/></Layer></Content></Page>")
        };

    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
                using (var stream = zip.CreateEntry(entry.Key).Open()) stream.Write(entry.Value);
        return output.ToArray();
    }

    private static Dictionary<string, byte[]> Unzip(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package, false), ZipArchiveMode.Read);
        var result = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries.Where(x => x.Name.Length > 0))
        {
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            result[entry.FullName] = copy.ToArray();
        }
        return result;
    }

    private static async Task<byte[]> LayoutOfd()
    {
        var document = new FlowDocument();
        document.Blocks.Add(new Paragraph("归档评估 Archival assessment"));
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(document.Render(), output);
        return output.ToArray();
    }

    private static async Task<ArchivalAssessmentReport> Assess(byte[] package, ArchivalAssessmentOptions? options = null,
        ArchivalAssessmentService? service = null)
    {
        using var input = new MemoryStream(package, false);
        var report = await (service ?? new ArchivalAssessmentService()).AssessAsync(input, options ?? new ArchivalAssessmentOptions());
        Assert.Equal(0, input.Position);
        Assert.Equal(package, input.ToArray());
        Assert.Equal(ArchivalClauseCatalog.Clauses.Select(x => x.Clause), report.Clauses.Select(x => x.Clause));
        Assert.All(report.Clauses, x => Assert.Same(ArchivalClauseMatrix.Get(x.Clause), x.Matrix));
        Assert.Equal(ArchivalInputProfile.SelfGeneratedSimpleId, report.InputSupport.Profile.Id);
        Assert.NotEqual(PreflightOutcome.Pass, report.Document.StandardOutcome);
        Assert.Equal("NotAssessed", report.SoftwareRequirements.Status);
        Assert.Equal("NotAuthenticated", report.ExternalRecognition.Status);
        if (report.ExecutionState != ArchivalExecutionState.Complete)
            Assert.All(report.Clauses, x => Assert.True(x.Outcome is PreflightOutcome.Unknown or PreflightOutcome.Fail));
        return report;
    }

    private static IEnumerable<string> FailedClauses(ArchivalAssessmentReport report) =>
        report.Clauses.Where(x => x.Outcome == PreflightOutcome.Fail).Select(x => x.Clause);

    [Fact]
    public async Task OrdinaryLayoutOutputKeepsDeclarationFailAndBindsInputIdentity()
    {
        var package = await LayoutOfd();
        var report = await Assess(package);
        Assert.Equal(ArchivalExecutionState.Complete, report.ExecutionState);
        Assert.Equal(Hash(package), report.Input!.Sha256);
        Assert.Equal(package.Length, report.Input.Length);
        Assert.Equal(ArchivalDeclarationStatus.NotArchival, report.Declaration.Status);
        Assert.Equal("OFD", report.Declaration.RawDocType);
        Assert.Equal(PreflightOutcome.Fail, report.Declaration.Outcome);
        Assert.Equal("NonArchivalDocType", report.Declaration.ReasonCode);
        Assert.Equal(ArchivalInputSupport.Supported, report.InputSupport.Support);
        Assert.Empty(report.InputSupport.Findings);
        Assert.Equal(PreflightOutcome.Fail, report.GetClause("6.2.1(a)").Outcome);
        Assert.Equal(new[] { "6.2.1(a)" }, FailedClauses(report));
        Assert.Equal(PreflightOutcome.Fail, report.Document.StandardOutcome);
        Assert.Equal(report.Coverage!.StandardOutcome, report.Document.StandardOutcome);
        // 17 Packaging observations plus AC-016 on 6.11: Layout output declares no CustomTags.
        Assert.Equal(18, report.Clauses.Sum(x => x.Observations.Count));
        Assert.Equal("NoCustomTagsInResult", Assert.Single(report.GetClause("6.11").Observations).ReasonCode);
        Assert.DoesNotContain("CustomTagClosureNotAssessed", report.Document.Gaps);
        Assert.Equal(87, report.Document.UncheckedClauses.Count);
        Assert.All(report.Document.UncheckedClauses, x => Assert.Equal(PreflightOutcome.Unknown, report.GetClause(x).Outcome));
        Assert.Equal(new[] { "7.1(1)", "7.1(2)", "7.2(1)", "7.2(2)" }, report.SoftwareRequirements.Clauses);
        var json = JsonSerializer.Serialize(report, new JsonSerializerOptions { Converters = { new JsonStringEnumConverter() } });
        Assert.Contains("ofdrw-archival-assessment/1", json);
        Assert.Contains("ofdrw-archival-matrix/1", json);
        Assert.DoesNotContain("IsCompliant", json);
    }

    [Fact]
    public async Task ArchivalLabelAloneChangesOnlyTheDeclaration()
    {
        var ordinary = await LayoutOfd();
        var entries = Unzip(ordinary);
        var root = Encoding.UTF8.GetString(entries["OFD.xml"]);
        Assert.Contains("DocType=\"OFD\"", root);
        entries["OFD.xml"] = Utf8(root.Replace("DocType=\"OFD\"", "DocType=\"OFD-A\"", StringComparison.Ordinal));
        var labelled = Zip(entries);
        Assert.All(Unzip(labelled).Where(x => x.Key != "OFD.xml"), x => Assert.Equal(Unzip(ordinary)[x.Key], x.Value));

        var before = await Assess(ordinary);
        var after = await Assess(labelled);
        Assert.Equal(ArchivalDeclarationStatus.Archival, after.Declaration.Status);
        Assert.Equal("OFD-A", after.Declaration.RawDocType);
        Assert.Equal(PreflightOutcome.Pass, after.Declaration.Outcome);
        Assert.Equal(PreflightOutcome.Pass, after.GetClause("6.2.1(a)").Outcome);
        Assert.Equal(before.InputSupport.Support, after.InputSupport.Support);
        foreach (var clause in before.Clauses.Where(x => x.Clause != "6.2.1(a)"))
        {
            var other = after.GetClause(clause.Clause);
            Assert.Equal(clause.Outcome, other.Outcome);
            Assert.Equal(clause.Observations.Select(x => (x.RuleId, x.Outcome, x.ReasonCode)),
                other.Observations.Select(x => (x.RuleId, x.Outcome, x.ReasonCode)));
        }
        Assert.Equal(PreflightOutcome.Unknown, after.Document.StandardOutcome);
        Assert.Equal(before.Document.UncheckedClauses, after.Document.UncheckedClauses);
        Assert.NotEqual(before.Input!.Sha256, after.Input!.Sha256);
    }

    [Fact]
    public async Task MissingDocTypeIsReadFromRawXml()
    {
        var report = await Assess(Zip(Seed(docType: "")));
        Assert.Equal(ArchivalDeclarationStatus.Missing, report.Declaration.Status);
        Assert.Null(report.Declaration.RawDocType);
        Assert.Equal("DocTypeMissing", report.Declaration.ReasonCode);
        Assert.Equal(PreflightOutcome.Fail, report.GetClause("6.2.1(a)").Outcome);
    }

    [Theory]
    [InlineData("DocType='OFD-A' xmlns:x='urn:vendor' x:DocType='OFD'", ArchivalDeclarationStatus.Ambiguous, "OFD-A", PreflightOutcome.Unknown)]
    [InlineData("xmlns:x='urn:vendor' x:DocType='OFD-A'", ArchivalDeclarationStatus.Missing, null, PreflightOutcome.Fail)]
    public async Task ForeignDocTypeAttributeIsNotTrusted(string docType, ArchivalDeclarationStatus status, string? raw,
        PreflightOutcome outcome)
    {
        var report = await Assess(Zip(Seed(docType: docType)));
        Assert.Equal(status, report.Declaration.Status);
        Assert.Equal(raw, report.Declaration.RawDocType);
        Assert.True(report.Declaration.ForeignDocTypePresent);
        Assert.Equal(outcome, report.Declaration.Outcome);
    }

    public static TheoryData<string, string> OutOfScopeInputs => new()
    {
        { "multiple", "MultipleDocuments" },
        { "signatures", "SignaturesDeclared" },
        { "encryption", "EncryptionMarkerPresent" },
        { "attachments", "AttachmentsDeclared" },
        { "versions", "VersionsDeclared" }
    };

    [Theory]
    [MemberData(nameof(OutOfScopeInputs))]
    public async Task OutOfScopeInputIsReportedWithoutFabricatedClauseFailures(string kind, string code)
    {
        var entries = Seed();
        switch (kind)
        {
            case "multiple":
                entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
                break;
            case "signatures":
                entries = Seed(bodyExtra: "<Signatures>Doc_0/Signs/Signatures.xml</Signatures>");
                entries["Doc_0/Signs/Signatures.xml"] = Utf8($"<Signatures xmlns='{Ns}'><Signature ID='1' BaseLoc='Sign_0/Signature.xml'/></Signatures>");
                break;
            case "encryption":
                entries["Encryptions.xml"] = Utf8($"<Encryptions xmlns='{Ns}'/>");
                break;
            case "attachments":
                entries = Seed(documentExtra: "<Attachments>Attachs/Attachments.xml</Attachments>");
                entries["Doc_0/Attachs/Attachments.xml"] = Utf8($"<Attachments xmlns='{Ns}'><Attachment ID='A1' Name='a'><FileLoc>a.txt</FileLoc></Attachment></Attachments>");
                entries["Doc_0/Attachs/a.txt"] = Utf8("attachment");
                break;
            case "versions":
                entries = Seed(bodyExtra: "<Versions><Version ID='v1' Index='1' BaseLoc='Doc_0/Version.xml'/></Versions>");
                break;
        }
        var report = await Assess(Zip(entries));
        Assert.Equal(ArchivalExecutionState.Complete, report.ExecutionState);
        Assert.Equal(ArchivalInputSupport.OutOfScope, report.InputSupport.Support);
        var finding = Assert.Single(report.InputSupport.Findings, x => x.Code == code);
        Assert.Equal(ArchivalInputSupport.OutOfScope, finding.Support);
        Assert.Contains("InputOutsideProfile", report.Document.Gaps);
        // Every clause Fail must come from an existing rule's definite violation, never from the scope decision.
        foreach (var clause in report.Clauses.Where(x => x.Outcome == PreflightOutcome.Fail))
            Assert.Contains(clause.Observations, x => x.Outcome == PreflightOutcome.Fail && x.RuleOwner == "Ofdrw.Net.Packaging");
        var expectedFailures = kind switch
        {
            "multiple" => new[] { "6.2.1(c)" },
            "encryption" => new[] { "6.16" },
            _ => Array.Empty<string>()
        };
        Assert.Equal(expectedFailures, FailedClauses(report));
        Assert.Equal(expectedFailures.Length > 0 ? PreflightOutcome.Fail : PreflightOutcome.Unknown, report.Document.StandardOutcome);
    }

    [Theory]
    [InlineData("<Future/>", "UnrecognizedDocumentContent")]
    [InlineData("<x:Vendor xmlns:x='urn:vendor'/>", "UnrecognizedDocumentContent")]
    public async Task UnrecognizedScopeIsUnknownNotOutOfScope(string extra, string code)
    {
        var report = await Assess(Zip(Seed(documentExtra: extra)));
        Assert.Equal(ArchivalInputSupport.Unknown, report.InputSupport.Support);
        Assert.Equal(code, Assert.Single(report.InputSupport.Findings).Code);
        Assert.Contains("InputScopeUnknown", report.Document.Gaps);
        Assert.Empty(FailedClauses(report));
    }

    [Fact]
    public async Task MissingPageContentIsUnknownScope()
    {
        var entries = Seed();
        entries.Remove("Doc_0/Pages/P0/Content.xml");
        var report = await Assess(Zip(entries));
        Assert.Equal(ArchivalInputSupport.Unknown, report.InputSupport.Support);
        Assert.Contains(report.InputSupport.Findings, x => x.Code == "PageContentUnavailable");
    }

    [Fact]
    public async Task ByteBudgetBeforeHashingIsReportedAsExhausted()
    {
        var package = Zip(Seed());
        var options = new ArchivalAssessmentOptions();
        options.Preflight.LoadOptions.MaxInputBytes = package.Length - 1;
        var report = await Assess(package, options);
        Assert.Equal(ArchivalExecutionState.BudgetExhausted, report.ExecutionState);
        Assert.Equal("InputBudgetExceeded", report.ExecutionReason);
        Assert.Null(report.Input);
        Assert.Null(report.Coverage);
        Assert.Equal(ArchivalDeclarationStatus.Unavailable, report.Declaration.Status);
        Assert.Equal(ArchivalInputSupport.Unknown, report.InputSupport.Support);
        Assert.Contains("AssessmentIncomplete:InputBudgetExceeded", report.Document.Gaps);
    }

    [Fact]
    public async Task ZipAndXmlBudgetsMarkTheRunIncomplete()
    {
        var package = Zip(Seed());
        var entryBudget = new ArchivalAssessmentOptions();
        entryBudget.Preflight.LoadOptions.MaxEntryCount = 1;
        var report = await Assess(package, entryBudget);
        Assert.Equal(ArchivalExecutionState.BudgetExhausted, report.ExecutionState);
        Assert.Equal("InputBudgetExceeded", report.ExecutionReason);
        Assert.Equal(Hash(package), report.Input!.Sha256);

        var xmlBudget = new ArchivalAssessmentOptions();
        xmlBudget.Preflight.MaxXmlNodes = 3;
        report = await Assess(package, xmlBudget);
        Assert.Equal(ArchivalExecutionState.BudgetExhausted, report.ExecutionState);
        Assert.Equal("InspectionBudgetExceeded", report.ExecutionReason);
        Assert.Equal(ArchivalInputSupport.Unknown, report.InputSupport.Support);
        Assert.Contains(report.InputSupport.Findings, x => x.Code == "AssessmentIncomplete");
    }

    [Fact]
    public async Task MalformedInputIsAnErrorState()
    {
        var report = await Assess(Utf8("not a zip package"));
        Assert.Equal(ArchivalExecutionState.Error, report.ExecutionState);
        Assert.Equal("InputRejected", report.ExecutionReason);
        Assert.Equal(ArchivalDeclarationStatus.Unavailable, report.Declaration.Status);
    }

    [Fact]
    public async Task CancellationThrowsAndRestoresTheInput()
    {
        var package = Zip(Seed());
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        using var input = new MemoryStream(package, false);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new ArchivalAssessmentService().AssessAsync(input, new ArchivalAssessmentOptions(), cancelled.Token));
        Assert.Equal(0, input.Position);

        using var source = new CancellationTokenSource();
        using var cancelling = new CancelOnRead(package, source);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new ArchivalAssessmentService().AssessAsync(cancelling, new ArchivalAssessmentOptions(), source.Token));
        Assert.Equal(0, cancelling.Position);
        Assert.Equal(package, cancelling.ToArray());
    }

    [Fact]
    public async Task BorrowingContractIsEnforced()
    {
        var service = new ArchivalAssessmentService();
        var package = Zip(Seed());
        using var moved = new MemoryStream(package, false) { Position = 1 };
        await Assert.ThrowsAsync<ArgumentException>(() => service.AssessAsync(moved, new ArchivalAssessmentOptions()));
        Assert.Equal(1, moved.Position);
        using var unseekable = new NonSeekableInput(package);
        await Assert.ThrowsAsync<ArgumentException>(() => service.AssessAsync(unseekable, new ArchivalAssessmentOptions()));
        Assert.Equal(0, unseekable.BytesRead);
        using var input = new MemoryStream(package, false);
        await Assert.ThrowsAsync<ArgumentException>(() =>
            service.AssessAsync(input, new ArchivalAssessmentOptions { ProfileId = "ofdrw-archival-profile/unknown/1" }));
    }

    [Fact]
    public void StrengthPolicyNeverFailsRecommendations()
    {
        var fail = Observation("6.11", PreflightOutcome.Fail);
        var pass = Observation("6.11", PreflightOutcome.Pass);
        foreach (var strength in new[] { ArchivalClauseStrength.Should, ArchivalClauseStrength.May })
        {
            Assert.Equal(PreflightOutcome.Unknown, ArchivalClauseAggregation.Outcome(strength, [fail, pass], true));
            Assert.Same(fail, Assert.Single(ArchivalClauseAggregation.Advisories(strength, [fail, pass])));
        }
        foreach (var strength in new[] { ArchivalClauseStrength.Statement, ArchivalClauseStrength.Shall })
        {
            Assert.Equal(PreflightOutcome.Fail, ArchivalClauseAggregation.Outcome(strength, [fail, pass], true));
            Assert.Empty(ArchivalClauseAggregation.Advisories(strength, [fail]));
        }
        Assert.Equal(PreflightOutcome.Unknown, ArchivalClauseAggregation.Outcome(ArchivalClauseStrength.Statement, [], true));
        Assert.Equal(PreflightOutcome.Pass, ArchivalClauseAggregation.Outcome(ArchivalClauseStrength.Statement, [pass], true));
        Assert.Equal(PreflightOutcome.Unknown, ArchivalClauseAggregation.Outcome(ArchivalClauseStrength.Statement, [pass], false));
        Assert.Equal(PreflightOutcome.Fail, ArchivalClauseAggregation.Outcome(ArchivalClauseStrength.Statement, [fail], false));
    }

    [Fact]
    public async Task OwnedCheckSeamAppliesStrengthPolicyToTheReport()
    {
        var service = new ArchivalAssessmentService(new IArchivalClauseCheck[]
        {
            new FixedCheck("AC-901", "6.11", PreflightOutcome.Fail),        // Should, matrix Unverified
            new FixedCheck("AC-902", "6.2.6(b)(1)", PreflightOutcome.Pass)   // Statement
        });
        var report = await Assess(Zip(Seed()), service: service);
        var recommendation = report.GetClause("6.11");
        Assert.Equal(ArchivalClauseStrength.Should, recommendation.Strength);
        Assert.Equal(ArchivalMatrixConfidence.Unverified, recommendation.Confidence);
        Assert.Equal(PreflightOutcome.Unknown, recommendation.Outcome);
        Assert.Equal("AC-901", Assert.Single(recommendation.Advisories).RuleId);
        Assert.Equal("Ofdrw.Net.Archival", recommendation.Observations[0].RuleOwner);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause("6.2.6(b)(1)").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, report.Document.StandardOutcome);

        report = await Assess(Zip(Seed()), service: new ArchivalAssessmentService([new FixedCheck("AC-903", "6.2.6(b)(1)", PreflightOutcome.Fail)]));
        Assert.Equal(PreflightOutcome.Fail, report.GetClause("6.2.6(b)(1)").Outcome);
        Assert.Equal(PreflightOutcome.Fail, report.Document.StandardOutcome);
    }

    [Theory]
    [InlineData("AC-015", "6.11")]
    [InlineData("AR-016", "6.11")]
    [InlineData("AC-016", "9.9")]
    public void OwnedCheckIdentitiesAreValidated(string id, string clause) =>
        Assert.Throws<ArgumentException>(() => new ArchivalAssessmentService([new FixedCheck(id, clause, PreflightOutcome.Pass)]));

    [Fact]
    public async Task ClauseNotesStateAssumptionsAndUnverifiedScope()
    {
        var report = await Assess(Zip(Seed()));
        Assert.Contains(report.GetClause("6.10.1(a)(3)").Notes, x => x.StartsWith("AnnotationObjectRetentionAssumed", StringComparison.Ordinal));
        Assert.Contains(report.GetClause("6.10.1(a)(2)").Notes, x => x.StartsWith("AnnotationObjectRetentionAssumed", StringComparison.Ordinal));
        Assert.Contains(report.GetClause("6.13.1(a)").Notes, x => x.StartsWith("ModalScopeUnverified", StringComparison.Ordinal));
    }

    private static ArchivalClauseObservation Observation(string clause, PreflightOutcome outcome) =>
        ArchivalOwnedRules.Observation(new ArchivalOwnedRuleDefinition("AC-900", clause, "synthetic"), outcome, "Synthetic",
            "<input>", "<package>", "synthetic");

    private sealed class FixedCheck(string id, string clause, PreflightOutcome outcome) : IArchivalClauseCheck
    {
        public ArchivalOwnedRuleDefinition Rule { get; } = new(id, clause, "synthetic condition");
        public string? GapCode => null;
        public bool Decided(IReadOnlyList<ArchivalClauseObservation> observations) => false;
        public IEnumerable<ArchivalClauseObservation> Check(ArchivalCheckContext context)
        {
            Assert.NotNull(context.Scope.Document);
            yield return ArchivalOwnedRules.Observation(Rule, outcome, "Synthetic", context.Scope.DocumentEntry!, "<document>", "synthetic");
        }
    }

    private sealed class CancelOnRead(byte[] bytes, CancellationTokenSource source) : MemoryStream(bytes, false)
    {
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            source.Cancel();
            return base.ReadAsync(buffer, offset, count, cancellationToken);
        }
    }
}
