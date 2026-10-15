using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TextReplacementIntegrityTests
{
    [Theory]
    [InlineData("escape")]
    [InlineData("no-delta")]
    [InlineData("rotation")]
    [InlineData("path-neighbor")]
    public async Task QuerySupportDoesNotAuthorizeUnsafeRenderingOrNeighborProfiles(string kind)
    {
        var source = await EditArchive(document =>
        {
            var text = document.Descendants().Single(value => value.Name.LocalName == "TextObject");
            var code = text.Elements().Single(value => value.Name.LocalName == "TextCode");
            if (kind == "escape") code.Value = "A\\0042C";
            if (kind == "no-delta") code.Attribute("DeltaX")!.Remove();
            if (kind == "rotation") text.SetAttributeValue("CTM", "0 1 -1 0 20 0");
            if (kind == "path-neighbor") text.AddAfterSelf(new XElement(text.Name.Namespace + "PathObject",
                new XAttribute("ID", "31"), new XAttribute("Boundary", "80 80 10 10"),
                new XElement(text.Name.Namespace + "AbbreviatedData", "M 0 0 L 1 1")));
        });
        var hit = Assert.Single(new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, "B").Hits);
        Assert.True(hit.CanReplace); var original = TextReplacementTests.Text(source).SourceXml;
        Assert.Throws<NotSupportedException>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source,
            new[] { new OfdTextReplacement(hit, "-") }));
        Assert.Equal(original, TextReplacementTests.Text(source).SourceXml);
    }

    [Fact]
    public async Task TranslationAndNonzeroPageOriginApplyOnce()
    {
        var source = await EditArchive(document => document.Descendants().Single(value => value.Name.LocalName == "TextObject")
            .SetAttributeValue("CTM", "1 0 0 1 1 2"));
        using var result = new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, "B", "--");
        var saved = await TextReplacementTests.Roundtrip(result.CreatePackageCopy());
        var suffix = Assert.Single(new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(saved, "C").Hits).Fragments[0];
        Assert.Equal(24, suffix.BaselineMillimeters.X); Assert.Equal(33, suffix.BaselineMillimeters.Y);
        Assert.Equal(new[] { 1d, 0, 0, 1, 1, 2 }, TextReplacementTests.Text(saved).Transform);
    }

    [Fact]
    public async Task InheritedStartsAreFrozenEvenWhenPredecessorIsDeleted()
    {
        var source = await EditArchive(document =>
        {
            var text = document.Descendants().Single(value => value.Name.LocalName == "TextObject");
            text.Elements().Where(value => value.Name.LocalName == "TextCode").Remove();
            var ns = text.Name.Namespace;
            text.Add(new XElement(ns + "TextCode", new XAttribute("X", "2"), new XAttribute("Y", "8"), "A"),
                new XElement(ns + "TextCode", new XAttribute("X", "5"), "B"),
                new XElement(ns + "TextCode", new XAttribute("X", "8"), "C"));
        });
        Assert.Equal(4, TextReplacementTests.Text(source).Runs[1].YMillimeters);
        using var result = new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, "AB", "-");
        var output = result.CreatePackageCopy(); var runs = TextReplacementTests.Text(output).Runs;
        Assert.Equal(new[] { "-", "", "C" }, runs.Select(value => value.Text));
        Assert.All(runs, value => Assert.Equal(8, value.YMillimeters));
        var saved = await TextReplacementTests.Roundtrip(output);
        var suffix = Assert.Single(new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(saved, "C").Hits).Fragments[0];
        Assert.Equal(23, suffix.BaselineMillimeters.X); Assert.Equal(31, suffix.BaselineMillimeters.Y);
    }

    [Fact]
    public async Task LeadingInteriorTrailingEmptySlotsKeepTheirOriginalOrdinals()
    {
        var source = await TextReplacementTests.Read([("", 0d, ""), ("A", 2d, ""), ("", 5d, ""), ("B", 5d, ""), ("", 8d, "")]);
        using var result = new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, "AB", "-");
        var output = result.CreatePackageCopy();
        Assert.Equal(new[] { "", "-", "", "", "" }, TextReplacementTests.Text(output).Runs.Select(value => value.Text));
        var saved = await TextReplacementTests.Roundtrip(output);
        Assert.Equal(5, TextReplacementTests.Text(saved).Runs.Count);
        Assert.Single(new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(saved, "-").Hits);
    }

    [Fact]
    public async Task ConcurrentConsumersHaveExactlyOneWinner()
    {
        var source = await TextReplacementTests.Read(); using var result = new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, "B", "-");
        using var ready = new Barrier(3); var winners = 0; var rejected = 0;
        Task Consume() => Task.Run(() =>
        {
            ready.SignalAndWait();
            try { var package = result.CreatePackageCopy(); Assert.Equal("A-C", TextReplacementTests.Text(package).Text); Interlocked.Increment(ref winners); }
            catch (InvalidOperationException) { Interlocked.Increment(ref rejected); }
        });
        var first = Consume(); var second = Consume(); ready.SignalAndWait(); await Task.WhenAll(first, second);
        Assert.Equal(1, winners); Assert.Equal(1, rejected);
    }

    [Fact]
    public async Task SourceFreeAndAliasedOwnerTargetsRefuseBeforeClone()
    {
        var sourceFree = KeywordLocatorTests.Model("ABC"); KeywordLocatorTests.Text(sourceFree).Runs[0].DeltaX = "3 3";
        var hit = Assert.Single(new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(sourceFree, "B").Hits);
        Assert.Throws<NotSupportedException>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(sourceFree,
            new[] { new OfdTextReplacement(hit, "-") }));
        var aliased = await TextReplacementTests.Read(); aliased.Pages[0].Elements.Add(aliased.Pages[0].Elements[0]);
        Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(aliased, "B", "-"));
    }

    private static async Task<OfdDocumentPackage> EditArchive(Action<XDocument> edit)
    {
        var package = await TextReplacementTests.Read(); using var original = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, original); original.Position = 0;
        using var result = new MemoryStream();
        using (var input = new ZipArchive(original, ZipArchiveMode.Read, true))
        using (var output = new ZipArchive(result, ZipArchiveMode.Create, true))
            foreach (var entry in input.Entries)
            {
                using var from = entry.Open(); using var to = output.CreateEntry(entry.FullName).Open();
                if (entry.FullName == "Doc_0/Pages/Page_0/Content.xml")
                { var xml = XDocument.Load(from, LoadOptions.PreserveWhitespace); edit(xml); xml.Save(to); }
                else await from.CopyToAsync(to);
            }
        result.Position = 0; return await new OfdReader().ReadAsync(result);
    }
}
