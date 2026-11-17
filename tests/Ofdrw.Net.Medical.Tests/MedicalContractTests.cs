using System.Text.RegularExpressions;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Preflight;
using Ofdrw.Net.Reader.Readers;
using static Ofdrw.Net.Medical.Tests.MedicalTestSupport;

namespace Ofdrw.Net.Medical.Tests;

public class MedicalContractTests
{
    [Fact]
    public void MatrixRulesAndBaselineDocumentAgree()
    {
        var rows = MedicalClauseMatrix.Entries;
        Assert.Equal(rows.Count, rows.Select(x => x.Clause).Distinct().Count());
        foreach (var rule in MedicalAssessmentService.Rules)
            foreach (var clause in rule.Clauses)
                Assert.Contains(rule.RuleId, MedicalClauseMatrix.Get(clause).Rules);
        foreach (var row in rows)
            foreach (var rule in row.Rules)
                Assert.Contains(row.Clause, MedicalAssessmentService.Rules.Single(x => x.RuleId == rule).Clauses);
        Assert.Equal(Enumerable.Range(1, 16).Select(i => $"MH-{i:000}"), MedicalAssessmentService.Rules.Select(x => x.RuleId));

        // The reviewed baseline lists exactly the same clauses, strengths and rules.
        var baseline = File.ReadAllText(Path.Combine(RepositoryRoot(), "docs/engineering/baselines/gbt48666-2026-clause-matrix.md"));
        var table = Regex.Matches(baseline, @"^\| `(?<clause>[^`]+)` \| (?<strength>\w+) \| (?<category>\w+) \| (?<rules>[^|]*) \|", RegexOptions.Multiline)
            .Select(m => (Clause: m.Groups["clause"].Value, Strength: m.Groups["strength"].Value, Category: m.Groups["category"].Value,
                Rules: m.Groups["rules"].Value.Trim() == "—" ? "" : m.Groups["rules"].Value.Trim())).ToList();
        Assert.Equal(rows.Select(x => x.Clause), table.Select(x => x.Clause));
        foreach (var (row, line) in rows.Zip(table))
        {
            Assert.Equal(row.Strength.ToString(), line.Strength);
            Assert.Equal(row.Category.ToString(), line.Category);
            Assert.Equal(string.Join(", ", row.Rules), line.Rules);
        }
        Assert.Contains(MedicalClauseMatrix.StandardSourceSha256, baseline);
        Assert.Contains(MedicalClauseMatrix.Revision, baseline);
    }

    [Fact]
    public void InheritanceLedgerCoversEveryArchivalClause()
    {
        Assert.Equal(ArchivalClauseCatalog.Clauses.Select(x => x.Clause), MedicalArchivalInheritance.Entries.Select(x => x.ArchivalClause));
        Assert.All(MedicalArchivalInheritance.ExplicitClauses, x => Assert.Contains(x, ArchivalClauseCatalog.Clauses.Select(c => c.Clause)));
        Assert.Equal(MedicalArchivalDisposition.Superseded, MedicalArchivalInheritance.Get("6.2.1(a)").Disposition);
        Assert.Equal("7.1(a)", MedicalArchivalInheritance.Get("6.2.1(a)").MedicalBasis);
        Assert.Equal(MedicalArchivalDisposition.Superseded, MedicalArchivalInheritance.Get("6.16").Disposition);
        Assert.Equal(MedicalArchivalDisposition.Inherited, MedicalArchivalInheritance.Get("6.5(b)").Disposition);
        Assert.Equal(MedicalMatrixConfidence.Unverified, MedicalArchivalInheritance.Get("6.5(b)").Confidence);
        Assert.All(MedicalArchivalInheritance.Entries.Where(x => x.Disposition != MedicalArchivalDisposition.Inherited),
            x => Assert.Contains(MedicalClauseMatrix.Entries, m => m.Clause == x.MedicalBasis || m.Clause.StartsWith(x.MedicalBasis + "(", StringComparison.Ordinal)));
    }

    [Theory]
    [InlineData("OFD-H", true, false)]
    [InlineData("OFD⁃H", false, true)]
    [InlineData("OFD‐H", false, true)]
    [InlineData("OFD－H", false, true)]
    [InlineData("ofd-h", false, false)]
    [InlineData("OFD-A", false, false)]
    [InlineData(null, false, false)]
    public void DocTypeIsExactAscii(string? value, bool medical, bool variant)
    {
        Assert.Equal(medical, MedicalDocTypes.IsMedical(value));
        Assert.Equal(variant, MedicalDocTypes.IsTypesetVariant(value));
    }

    [Fact]
    public async Task ProfileConfiguresOptionsAndDocIdRoundTrips()
    {
        var options = MedicalDocumentProfile.Apply(new OfdDocumentOptions(), Metadata());
        Assert.Equal("OFD-H", options.DocType);
        Assert.Equal(OfdConstants.StandardNamespace, options.Namespace);
        Assert.Equal("Doc_0", options.DocumentId);
        Assert.True(options.EnableDeflateCompression && options.LinearizedEntryOrder);
        Assert.Equal(new[] { "EMR.PatientID", "EMR.FileType", "EMR.SecretLevel" }, options.Metadata.CustomDatas!.Items.Select(x => x.Name));
        // Applying twice with the same values is idempotent.
        MedicalDocumentProfile.Apply(options, Metadata());
        Assert.Equal(3, options.Metadata.CustomDatas!.Items.Count);

        var package = await new OfdReader().ReadAsync(new MemoryStream(await LayoutOfd(true)));
        Assert.Equal("EMR-2026-0001", package.Options.Metadata.DocId);
        Assert.Equal("OFD-H", package.Options.DocType);
    }

    [Fact]
    public void ProfileRejectsIncompleteOrConflictingMetadata()
    {
        var missing = Metadata(); missing.PatientId = " ";
        Assert.Throws<ArgumentException>(() => MedicalDocumentProfile.Apply(new OfdDocumentOptions(), missing));
        var undated = Metadata(); undated.CreationTime = default;
        Assert.Throws<ArgumentException>(() => MedicalDocumentProfile.Apply(new OfdDocumentOptions(), undated));
        Assert.Throws<ArgumentException>(() => MedicalDocumentProfile.Apply(new OfdDocumentOptions(), Metadata(),
            new MedicalMetadataMapping { PatientField = "A", FileTypeField = "A", SecretLevelField = "B" }));
        var conflicting = new OfdDocumentOptions();
        conflicting.Metadata.CustomDatas = new OfdCustomDatas();
        conflicting.Metadata.CustomDatas.Add("EMR.PatientID", "someone else");
        Assert.Throws<ArgumentException>(() => MedicalDocumentProfile.Apply(conflicting, Metadata()));
        Assert.Equal("OFD", conflicting.DocType);
    }

    [Fact]
    public async Task DefaultWriterOrderIsUnchangedAndLinearizedOrderIsOptIn()
    {
        var ordinary = Order(await LayoutOfd(false));
        Assert.Equal(ordinary.OrderBy(x => x, StringComparer.OrdinalIgnoreCase), ordinary);
        var linear = Order(await LayoutOfd(true));
        Assert.Equal("OFD.xml", linear[0]);
        Assert.Equal("Doc_0/Document.xml", linear[1]);
        Assert.True(linear.IndexOf("Doc_0/Pages/Page_0/Content.xml") < linear.FindIndex(x => x.StartsWith("Doc_0/Res/")));
        Assert.True(linear.FindIndex(x => x.StartsWith("Doc_0/Res/")) < linear.IndexOf("Doc_0/Pages/Page_1/Content.xml"));
        Assert.Equal(ordinary.Count, linear.Count);
    }

    [Theory]
    [InlineData(MedicalEditOperation.InsertPage, MedicalEditDecision.RequiresAuthorization)]
    [InlineData(MedicalEditOperation.RemovePage, MedicalEditDecision.RequiresAuthorization)]
    [InlineData(MedicalEditOperation.AdjustPages, MedicalEditDecision.RequiresAuthorization)]
    [InlineData(MedicalEditOperation.EditPageContent, MedicalEditDecision.RequiresAuthorization)]
    [InlineData(MedicalEditOperation.AddAnnotation, MedicalEditDecision.Allowed)]
    [InlineData(MedicalEditOperation.Sign, MedicalEditDecision.Allowed)]
    [InlineData(MedicalEditOperation.EditMetadata, MedicalEditDecision.AllowedWithResign)]
    [InlineData(MedicalEditOperation.Combine, MedicalEditDecision.AllowedWithResign)]
    public void EditPolicyFollows92bAnd73g(MedicalEditOperation operation, MedicalEditDecision decision)
    {
        Assert.Equal(decision, MedicalEditPolicy.Classify("OFD-H", operation));
        Assert.Equal(MedicalEditDecision.Allowed, MedicalEditPolicy.Classify("OFD", operation));
        Assert.Equal(MedicalEditDecision.Allowed, MedicalEditPolicy.Classify("OFD⁃H", operation));
        if (decision == MedicalEditDecision.RequiresAuthorization)
        {
            Assert.Equal(operation, Assert.Throws<MedicalEditNotAuthorizedException>(() => MedicalEditPolicy.EnsureAllowed("OFD-H", operation, false)).Operation);
            MedicalEditPolicy.EnsureAllowed("OFD-H", operation, true);
        }
        else MedicalEditPolicy.EnsureAllowed("OFD-H", operation, false);
    }
}
