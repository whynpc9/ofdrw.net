using System.Globalization;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TextReplacementTests
{
    [Theory]
    [InlineData("--", "A--C", "3 1.25 1.75")]
    [InlineData("-", "A-C", "3 3")]
    [InlineData("", "AC", "6")]
    public async Task ActualReaderHitReplacementWriterReaderPreservesSuffix(string replacement, string expected, string delta)
    {
        var source = await Read(); var originalXml = Text(source).SourceXml; var originalBytes = source.Fonts[0].Data.ToArray();
        var provider = new Metrics(); var hit = Assert.Single(new OfdKeywordLocator(provider).Find(source, "B").Hits);
        using var result = new OfdTextReplacer(provider).Replace(source, new[] { new OfdTextReplacement(hit, replacement) });
        Assert.True(result.Changed); Assert.Equal(1, result.RequestedCount); Assert.Equal(1, result.ChangedCount);
        var output = result.CreatePackageCopy(); Assert.Equal(expected, Text(output).Text);
        Assert.Equal(delta, Text(output).Runs[0].DeltaX); Assert.Equal(originalXml, Text(source).SourceXml);
        Assert.NotSame(source, output); Assert.NotSame(source.Fonts[0].Data, output.Fonts[0].Data);
        Assert.Equal(originalBytes, source.Fonts[0].Data);
        Assert.Throws<InvalidOperationException>(() => result.CreatePackageCopy()); result.Dispose();
        var saved = await Roundtrip(output);
        Assert.Equal(expected, Text(saved).Text); Assert.Equal("30", Text(saved).ObjectId); Assert.Equal("10", Text(saved).FontResourceId);
        var suffix = Assert.Single(new OfdKeywordLocator(new Metrics()).Find(saved, "C").Hits).Fragments[0];
        Assert.Equal(23, suffix.BaselineMillimeters.X); Assert.Equal(31, suffix.BaselineMillimeters.Y);
    }

    [Fact]
    public async Task KeywordOverloadCompletesSearchAndFitInsideOneOwnedProviderSession()
    {
        var source = await Read(); var provider = new Metrics();
        using var result = new OfdTextReplacer(provider).Replace(source, "B", "--");
        Assert.Contains(result.Statistics.Phases, value => value.Name == "KeywordMatchesCaptured");
        Assert.Single(result.Statistics.Analyses); Assert.Equal(1, provider.Opens);
        Assert.Equal(1, provider.StatisticsCalls); Assert.Equal(1, provider.Disposals); Assert.Equal(2, provider.Measures);
        Assert.Equal("A--C", Text(result.CreatePackageCopy()).Text);
    }

    [Fact]
    public async Task FullDeletionRetainsRootAllOriginalSlotsAndExplicitEmptyAnchors()
    {
        var source = await Read([("A", 2d, ""), ("B", 5d, ""), ("C", 8d, "")]);
        using var result = new OfdTextReplacer(new Metrics()).Replace(source, "ABC", "");
        var output = result.CreatePackageCopy(); var text = Text(output);
        Assert.Equal("30", text.ObjectId); Assert.Equal("", text.Text); Assert.Equal(3, text.Runs.Count);
        Assert.Equal(new[] { 2d, 5, 8 }, text.Runs.Select(value => value.XMillimeters));
        Assert.All(text.Runs, value => { Assert.Empty(value.Text); Assert.Equal(8, value.YMillimeters); Assert.Null(value.DeltaX); Assert.Null(value.DeltaY); });
        var saved = await Roundtrip(output); var provider = new Metrics();
        var found = new OfdKeywordLocator(provider).Find(saved, "ABC");
        Assert.True(found.IsComplete); Assert.Empty(found.Hits); Assert.Equal(0, provider.Opens);
        Assert.Equal(3, Text(saved).Runs.Count);
    }

    [Fact]
    public async Task CrossSlotEditKeepsConsumedSlotsEmptyAndFreezesSuffix()
    {
        var source = await Read([("A", 2d, ""), ("B", 5d, ""), ("C", 8d, ""), ("D", 11d, "")]);
        using var result = new OfdTextReplacer(new Metrics()).Replace(source, "BC", "--");
        var output = result.CreatePackageCopy(); var text = Text(output);
        Assert.Equal(new[] { "A", "--", "", "D" }, text.Runs.Select(value => value.Text));
        Assert.Equal(new[] { 2d, 5, 8, 11 }, text.Runs.Select(value => value.XMillimeters));
        var saved = await Roundtrip(output);
        var suffix = Assert.Single(new OfdKeywordLocator(new Metrics()).Find(saved, "D").Hits).Fragments[0];
        Assert.Equal(26, suffix.BaselineMillimeters.X); Assert.Equal(31, suffix.BaselineMillimeters.Y);
    }

    [Fact]
    public async Task NoOpIsLazyIndependentAndSingleUse()
    {
        var source = await Read(); using var result = new OfdTextReplacer(new Metrics()).Replace(source, "B", "B");
        Assert.False(result.Changed); Assert.DoesNotContain(result.Statistics.Phases, value => value.Name == "PreparedCandidate");
        var output = result.CreatePackageCopy(); Text(output).Runs[0].Text = "changed";
        Assert.Equal("ABC", Text(source).Runs[0].Text); Assert.Equal("ABC", Text(source).Text);
        using var noHit = new OfdTextReplacer(new Metrics()).Replace(source, "missing", "-");
        Assert.False(noHit.Changed); Assert.Equal(0, noHit.RequestedCount);
    }

    [Theory]
    [InlineData("foreign")]
    [InlineData("mixed")]
    [InlineData("stale-text")]
    [InlineData("stale-font")]
    [InlineData("duplicate")]
    public async Task StaleForeignMixedAndOverlappingHitsReject(string kind)
    {
        var source = await Read(); var locator = new OfdKeywordLocator(new Metrics());
        var hit = Assert.Single(locator.Find(source, "B").Hits); var requests = new List<OfdTextReplacement> { new(hit, "-") };
        if (kind == "foreign") source = await Read();
        if (kind == "mixed") requests.Add(new(Assert.Single(locator.Find(source, "C").Hits), "-"));
        if (kind == "stale-text") Text(source).Text = "changed";
        if (kind == "stale-font") source.Fonts[0].Data[0]++;
        if (kind == "duplicate") requests.Add(new(hit, "B"));
        Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(new Metrics()).Replace(source, requests));
    }

    [Theory]
    [InlineData("open")]
    [InlineData("measure")]
    [InlineData("statistics")]
    [InlineData("dispose")]
    public async Task WholeSerializedSourceGuardFollowsEveryProviderCallback(string phase)
    {
        var source = await Read(); var hit = Assert.Single(new OfdKeywordLocator(new Metrics()).Find(source, "B").Hits);
        var provider = new Metrics { Callback = value => { if (value == phase) source.Options.Metadata.Title = "mutated"; } };
        var error = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(provider).Replace(source, new[] { new OfdTextReplacement(hit, "-") }));
        Assert.Equal(phase switch { "open" => "OpenProvider", "measure" => "MeasureSource", "statistics" => "ProviderStatistics", _ => "DisposeProvider" },
            error.Data["Ofdrw.TextReplacement.Phase"]);
        Assert.Equal(1, provider.Disposals);
    }

    [Theory]
    [InlineData("WW")]
    [InlineData(" ")]
    [InlineData("\\")]
    [InlineData("\n")]
    [InlineData("a\u0301")]
    [InlineData("مرحبا")]
    public async Task OverflowAndUnsafeLogicalProfilesRejectWithoutChangingSource(string replacement)
    {
        var source = await Read(); var before = Text(source).SourceXml;
        Assert.Throws<NotSupportedException>(() => new OfdTextReplacer(new Metrics()).Replace(source, "B", replacement));
        Assert.Equal(before, Text(source).SourceXml);
    }

    [Fact]
    public async Task LowBudgetFailsBeforeProviderAndIncludesPhaseCounters()
    {
        var source = await Read(); var provider = new Metrics();
        var options = new OfdTextReplacementOptions(new(maxMaterializedBytes: 1_000_000));
        var error = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(provider).Replace(source, "B", "-", options));
        Assert.Equal(0, provider.Opens); Assert.Equal("KeywordQuery", error.Data["Ofdrw.TextReplacement.Phase"]);
        Assert.IsType<OfdTextReplacementPhase>(error.Data["Ofdrw.TextReplacement.Counters"]);
    }

    [Fact]
    public async Task CancelledOrStaleConsumptionIsTerminal()
    {
        var source = await Read(); using var result = new OfdTextReplacer(new Metrics()).Replace(source, "B", "-");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => result.CreatePackageCopy(cancelled.Token));
        Assert.Throws<InvalidOperationException>(() => result.CreatePackageCopy());
        using var stale = new OfdTextReplacer(new Metrics()).Replace(source, "B", "B");
        source.Options.Metadata.Title = "after preparation";
        Assert.Throws<InvalidDataException>(() => stale.CreatePackageCopy());
        Assert.Throws<InvalidOperationException>(() => stale.CreatePackageCopy());
    }

    internal static OfdTextElement Text(OfdDocumentPackage package) => (OfdTextElement)package.Pages[0].Elements[0];
    internal static async Task<OfdDocumentPackage> Read((string Text, double X, string Delta)[]? slots = null,
        double textWidth = 100, double pageWidth = 200)
    {
        var package = new OfdDocumentPackage(); package.Fonts.Add(new() { Id = "10", FontName = "Fixture", FileName = "font.ttf", Data = [1, 2, 3] });
        var page = new OfdPage { Id = "1", XMillimeters = 5, YMillimeters = 7, WidthMillimeters = pageWidth, HeightMillimeters = 200 };
        var text = new OfdTextElement { ObjectId = "30", LayerId = "20", FontResourceId = "10", FontName = "Fixture",
            XMillimeters = 20, YMillimeters = 30, WidthMillimeters = textWidth, HeightMillimeters = 50, FontSizeMillimeters = 4 };
        foreach (var slot in slots ?? [("ABC", 2d, "3 3")]) text.Runs.Add(new() { Text = slot.Text, XMillimeters = slot.X, YMillimeters = 8, DeltaX = slot.Delta.Length == 0 ? null : slot.Delta });
        text.Text = string.Concat(text.Runs.Select(value => value.Text)); page.Elements.Add(text); package.Pages.Add(page);
        return await Roundtrip(package);
    }
    internal static async Task<OfdDocumentPackage> Roundtrip(OfdDocumentPackage package)
    {
        using var stream = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, stream); stream.Position = 0;
        return await new OfdReader().ReadAsync(stream);
    }
    internal sealed class Metrics : IOfdGlyphGeometryProvider
    {
        internal Action<string>? Callback; internal int Opens, Disposals, StatisticsCalls, Measures;
        public IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits limits, CancellationToken cancellationToken = default)
        { Opens++; Callback?.Invoke("open"); return new Session(this); }
        private sealed class Session(Metrics owner) : IOfdGlyphGeometrySession
        {
            private long glyphs;
            public OfdGlyphGeometryStatistics Statistics { get { owner.StatisticsCalls++; owner.Callback?.Invoke("statistics"); return new(glyphRequests: glyphs); } }
            public OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken cancellationToken = default)
            {
                owner.Measures++; owner.Callback?.Invoke("measure"); var result = new List<OfdGlyphClusterGeometry>();
                for (var i = 0; i < request.Text.Length;)
                {
                    var scalar = char.ConvertToUtf32(request.Text, i); var length = scalar > 0xffff ? 2 : 1;
                    var hyphen = scalar == '-'; var advance = hyphen ? 1.25 : scalar == 'W' ? 4 : 3;
                    result.Add(new(i, length, (uint)scalar, hyphen ? new(.25, -1.5, 1, -.5) : new(.5, -2, 2.5, 0), advance));
                    i += length; glyphs++;
                }
                return new(result);
            }
            public void Dispose() { owner.Disposals++; owner.Callback?.Invoke("dispose"); }
        }
    }
}
