using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Audit025RegressionTests
{
    private readonly Xunit.Abstractions.ITestOutputHelper output;
    public Audit025RegressionTests(Xunit.Abstractions.ITestOutputHelper output) => this.output = output;
    private static readonly Lazy<OpenTypeCmap> Coverage = new(() => new OpenTypeCmap(new OpenTypeFace(File.ReadAllBytes(FontPath()))));
    private static OfdTextElement Element(string value, string? deltaX = null, string? deltaY = null)
        => new() { Text = value, Runs = { new OfdTextRun { Text = value, DeltaX = deltaX, DeltaY = deltaY } } };

    [Fact]
    public void Audit025_NormalControl()
        => EmbeddedFontCoverage.Validate(Element("AV e\u0301"), Coverage.Value, "control");

    [Fact]
    public async Task Audit025_ElementOverloadNoDelta()
    {
        // A separate process prevents other tests from invalidating the bounded
        // no-GC intervals. Every sample must meet the original 4096-byte budget;
        // the probe never retries an interval or selects a minimum sample.
        var receipt = await AllocationProbe.RunBudgetChild(FontPath(), output);
        Assert.True(receipt.Passed);
        Assert.Equal(16, receipt.Cases.Count);
        Assert.All(receipt.Cases, item =>
        {
            Assert.True(item.ValidNoGcInterval, item.Error);
            Assert.Equal(9, item.Samples.Length);
            Assert.All(item.Samples.Skip(1), sample =>
            {
                if (item.Operation == "target") Assert.True(sample.Allocated < 4096,
                    $"{item.DeltaRole}/{item.Length}/{sample.Ordinal}: {sample.Allocated} bytes.");
                else if (item.Operation == "empty") Assert.Equal(0, sample.Allocated);
                else Assert.True(sample.Allocated >= item.MinimumExpectedAllocation);
            });
        });
        // Restore the former per-grapheme work only in a separate negative
        // control process. The same budget evaluator must reject every long run.
        var regression = await AllocationProbe.RunRegressionChild(FontPath(), output);
        Assert.False(regression.Passed);
        Assert.Null(regression.Error);
        Assert.Equal(16, regression.Cases.Count);
        Assert.All(regression.Cases, item => Assert.True(item.ValidNoGcInterval, item.Error));
        Assert.All(regression.Cases.Where(item => item.Operation == "target" && item.Length >= 4096), item =>
            Assert.All(item.Samples.Skip(1), sample => Assert.True(sample.Allocated >= 4096)));
    }

    [Theory]
    [InlineData("bad")] [InlineData("NaN")] [InlineData("Infinity")]
    [InlineData("g -1 2")] [InlineData("g 2")] [InlineData("g 0 bad")]
    public void Audit025_InvalidDelta(string delta)
    {
        Assert.Throws<InvalidDataException>(() => EmbeddedFontCoverage.Validate(Element("AB", delta), Coverage.Value, "invalid"));
        // Y must be parsed even when X already establishes positioning.
        Assert.Throws<InvalidDataException>(() => EmbeddedFontCoverage.Validate(Element("AB", "3", delta), Coverage.Value, "invalid-y"));
    }

    [Fact]
    public void Audit025_PositionedWhitespace()
    {
        var face = new OpenTypeFace(File.ReadAllBytes(FontPath()));
        face.Tables["cmap"] = OpenTypeCmap.Build(new Dictionary<int, int> { ['A'] = Coverage.Value.Glyph('A') });
        var cmap = new OpenTypeCmap(face);
        EmbeddedFontCoverage.Validate(Element("A\u202FA", "3 4"), cmap, "positioned");
        EmbeddedFontCoverage.Validate(Element("A\u202FA", deltaY: "3 4"), cmap, "positioned-y");
        EmbeddedFontCoverage.Validate(Element("\u202F\u202F"), cmap, "all-blank");
        foreach (var delta in new string?[] { null, "", " ", "g 0 3" })
            Assert.Throws<InvalidDataException>(() => EmbeddedFontCoverage.Validate(Element("A\u202FA", delta), cmap, "mixed"));
    }

    [Fact]
    public void Audit025_CombiningAndVariation()
    {
        var value = "e\u0301A\uFE0FB";
        Assert.Equal(new[] { "e\u0301", "A\uFE0F", "B" }, OfdTextGeometry.EnumerateGlyphs(value, default));
        var deltas = OfdTextGeometry.PositioningDeltas(Element(value, "3 5 99").Runs[0], default);
        Assert.Equal(new[] { 3d, 5 }, deltas.X);
        EmbeddedFontCoverage.Validate(Element("e\u0301B", "3"), Coverage.Value, "combining");
        foreach (var delta in new[] { null, "3" })
            Assert.Throws<InvalidDataException>(() => EmbeddedFontCoverage.Validate(Element("A\uFE0FB", delta), Coverage.Value, "unsupported-variation"));
        // Existing bounded semantics: no adjacent origin exists for one grapheme,
        // and unused trailing delta tokens are not interpreted.
        Assert.Empty(OfdTextGeometry.PositioningDeltas(Element("e\u0301", "bad").Runs[0], default).X);
        Assert.Equal(new[] { 3d }, OfdTextGeometry.PositioningDeltas(Element("AB", "3 bad").Runs[0], default).X);
    }

    [Fact]
    public void Audit025_CancelLongRun()
    {
        using var source = new CancellationTokenSource();
        using var glyphs = OfdTextGeometry.EnumerateGlyphs(new string('A', 8192), source.Token).GetEnumerator();
        Assert.True(glyphs.MoveNext()); Assert.Equal("A", glyphs.Current);
        source.Cancel();
        Assert.Throws<OperationCanceledException>(() => glyphs.MoveNext());
        Assert.Throws<OperationCanceledException>(() => OfdTextGeometry.PositioningDeltas(Element("AB", "g 1 2").Runs[0], source.Token));
        Assert.Throws<OperationCanceledException>(() => OfdTextGeometry.ExpandDeltas("g 8192 3", 8192, source.Token));
        Assert.Throws<OperationCanceledException>(() => EmbeddedFontCoverage.Validate(Element(new string('A', 8192)), Coverage.Value, "cancel", source.Token));
    }

    [Fact]
    public void Audit025_ElementNoDeltaMidScanCancellation()
    {
        using var source = new CancellationTokenSource();
        var text = Element(new string('A', 4096)); var completedScalars = 0;
        Assert.Throws<OperationCanceledException>(() => EmbeddedFontCoverage.Validate(text, Coverage.Value,
            "element-mid-scan", source.Token, scalar =>
            {
                Assert.Equal('A', scalar);
                completedScalars++;
                source.Cancel();
            }));
        // The actual element overload reached coverage, validated one scalar,
        // then observed cancellation before completing another scalar.
        Assert.Equal(1, completedScalars);
    }

    private static string FontPath()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return Path.Combine(root!.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf");
    }
}
