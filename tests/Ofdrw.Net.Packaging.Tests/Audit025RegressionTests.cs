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

    [Theory]
    [InlineData(null)] [InlineData("")] [InlineData(" \t\r\n")]
    public void Audit025_ElementOverloadNoDelta(string? delta)
    {
        var text = Element(new string('A', 4096), delta); var cmap = Coverage.Value;
        EmbeddedFontCoverage.Validate(text, cmap, "warmup");
        var before = GC.GetAllocatedBytesForCurrentThread();
        EmbeddedFontCoverage.Validate(text, cmap, "measured");
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        output.WriteLine($"Measured allocation: {allocated} bytes for 4096 characters.");
        // 4096 text-element strings alone cost over 96 KiB. Allow fixed iterator overhead.
        Assert.True(allocated < 4096, $"Coverage allocated {allocated} bytes for a no-delta run.");
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
