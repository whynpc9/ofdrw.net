using System.Text.Json;
using Ofdrw.Net.Archival.Internal;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival.Tests;

public class ArchivalClauseMatrixTests
{
    private static string RepositoryRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !File.Exists(Path.Combine(root.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        Assert.NotNull(root);
        return root!.FullName;
    }

    private static JsonElement Baseline()
    {
        var path = Path.Combine(RepositoryRoot(), "docs/engineering/baselines/gbt42133-2022-clause-applicability.json");
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return document.RootElement.Clone();
    }

    private static string? Text(JsonElement item, string name) =>
        item.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null ? value.GetString() : null;

    private static string[] Strings(JsonElement item, string name) =>
        item.GetProperty(name).EnumerateArray().Select(x => x.GetString()!).ToArray();

    [Fact]
    public void EmbeddedMatrixEqualsRepositoryBaselineItemForItem()
    {
        var json = Baseline();
        Assert.Equal(ArchivalClauseMatrix.LedgerRevision, json.GetProperty("ledgerRevision").GetString());
        Assert.Equal(ArchivalClauseMatrix.StandardSourceSha256, json.GetProperty("sourceSha256").GetString());
        Assert.Equal(ArchivalClauseMatrix.ReviewedOn, json.GetProperty("reviewedOn").GetString());
        Assert.StartsWith(ArchivalClauseMatrix.StandardId, json.GetProperty("standard").GetString());
        Assert.Equal(Strings(json, "strengthOrder"), ArchivalClauseMatrix.StrengthOrder.Select(x => x.ToString()).ToArray());
        var items = json.GetProperty("items").EnumerateArray().ToList();
        Assert.Equal(items.Count, ArchivalClauseMatrix.Entries.Count);
        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            var entry = ArchivalClauseMatrix.Entries[i];
            Assert.Equal(item.GetProperty("clause").GetString(), entry.Clause);
            Assert.Equal(item.GetProperty("strength").GetString(), entry.Strength.ToString());
            Assert.Equal(item.GetProperty("modalKeyword").GetString(), entry.ModalKeyword);
            var mixed = item.TryGetProperty("mixedStrengths", out var values)
                ? values.EnumerateArray().Select(x => (x.GetProperty("strength").GetString(), x.GetProperty("part").GetString())).ToArray()
                : [];
            Assert.Equal(mixed, entry.MixedStrengths.Select(x => ((string?)x.Strength.ToString(), (string?)x.Part)).ToArray());
            Assert.Equal(item.GetProperty("conditional").GetBoolean(), entry.Conditional);
            Assert.Equal(Text(item, "condition"), entry.Condition);
            Assert.Equal(item.GetProperty("pdfPage").GetInt32(), entry.PdfPage);
            Assert.Equal(item.GetProperty("printedPage").GetString(), entry.PrintedPage);
            Assert.Equal(item.GetProperty("confidence").GetString(), entry.Confidence.ToString());
            Assert.Equal(item.GetProperty("firstProfile").GetString(), entry.FirstProfile.ToString());
            Assert.Equal(Strings(item, "rules"), entry.Rules);
            Assert.Equal(Strings(item, "referencedStandards"), entry.ReferencedStandards);
            Assert.Equal(item.GetProperty("editionFixed").GetBoolean(), entry.EditionFixed);
            Assert.Equal(Text(item, "note"), entry.Note);
            Assert.Equal(
                item.EnumerateObject().Select(x => x.Name).Order(StringComparer.Ordinal),
                new[] { "clause", "strength", "modalKeyword", "conditional", "condition", "pdfPage", "printedPage", "confidence",
                        "firstProfile", "rules", "referencedStandards", "editionFixed" }
                    .Concat(item.TryGetProperty("mixedStrengths", out _) ? ["mixedStrengths"] : Array.Empty<string>())
                    .Concat(item.TryGetProperty("note", out _) ? ["note"] : Array.Empty<string>())
                    .Order(StringComparer.Ordinal));
        }
    }

    [Fact]
    public void MatrixFollowsLedgerOrderAndMappedRules()
    {
        Assert.Equal(105, ArchivalClauseCatalog.Clauses.Count);
        Assert.Equal(ArchivalClauseCatalog.Revision, ArchivalClauseMatrix.LedgerRevision);
        Assert.Equal(ArchivalClauseCatalog.Clauses.Select(x => x.Clause), ArchivalClauseMatrix.Entries.Select(x => x.Clause));
        foreach (var clause in ArchivalClauseCatalog.Clauses)
            Assert.Equal(clause.RuleIds.Order(StringComparer.Ordinal),
                ArchivalClauseMatrix.Get(clause.Clause).Rules.Where(x => x.StartsWith("A", StringComparison.Ordinal) && x.Contains('-'))
                    .Order(StringComparer.Ordinal));
        var strengths = ArchivalClauseMatrix.Entries.GroupBy(x => x.Strength).ToDictionary(x => x.Key, x => x.Count());
        Assert.Equal(68, strengths[ArchivalClauseStrength.Statement]);
        Assert.Equal(24, strengths[ArchivalClauseStrength.Should]);
        Assert.Equal(9, strengths[ArchivalClauseStrength.May]);
        Assert.Equal(4, strengths[ArchivalClauseStrength.Shall]);
        Assert.Equal(new[] { "6.11", "6.13.1(a)", "6.13.1(b)", "6.13.1(c)" },
            ArchivalClauseMatrix.Entries.Where(x => x.Confidence == ArchivalMatrixConfidence.Unverified).Select(x => x.Clause));
    }

    [Fact]
    public void ArchivalStaysOptInAtSourceLevel()
    {
        var projects = Directory.GetFiles(Path.Combine(RepositoryRoot(), "src"), "*.csproj", SearchOption.AllDirectories);
        var references = projects.ToDictionary(Path.GetFileNameWithoutExtension, path =>
            System.Xml.Linq.XDocument.Load(path).Descendants("ProjectReference")
                .Select(x => Path.GetFileNameWithoutExtension((string)x.Attribute("Include")!)).ToArray());
        Assert.Equal(new[] { "Ofdrw.Net.Packaging" }, references["Ofdrw.Net.Archival"]);
        Assert.DoesNotContain(references, x => x.Key != "Ofdrw.Net.Archival" && x.Value.Contains("Ofdrw.Net.Archival"));
    }

    [Fact]
    public void OwnedRuleNumberingContinuesAfterPackagingCatalog()
    {
        Assert.Equal(15, ArchivalCoverageRuleCatalog.Rules.Count);
        Assert.Equal("AC-015", ArchivalCoverageRuleCatalog.Rules[^1].RuleId);
        Assert.Equal(new[] { "AC-016" }, ArchivalOwnedRules.Registered.Select(x => x.Rule.RuleId));
        Assert.False(ArchivalOwnedRules.IsOwnedId("AC-015"));
        Assert.False(ArchivalOwnedRules.IsOwnedId("AR-016"));
        Assert.True(ArchivalOwnedRules.IsOwnedId("AC-016"));
    }
}
