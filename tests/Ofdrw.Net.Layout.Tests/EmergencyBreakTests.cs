using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class EmergencyBreakTests
{
    private static readonly Lazy<TextStyle> Noto = new(() => One("Noto Sans", "NotoSans-Regular.ttf"));
    private static readonly Lazy<TextStyle> DejaVu = new(() => One("Ofdrw Test DejaVu Sans", "DejaVuSans-TextFixture.ttf"));
    private static readonly Lazy<TextStyle> Thai = new(() => One("Noto Sans Thai", "NotoSansThai-Regular.ttf"));

    private static TextStyle One(string family, string file) =>
        new(new TextFontFamily(family, ParagraphLayoutTests.Face(family, file)), 4);
    private static string Repeat(string text, int count) => string.Concat(Enumerable.Repeat(text, count));

    private static ParagraphLayoutResult Layout(string text, TextStyle style, double width, double indent = 0) =>
        HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(text, style,
            new ParagraphLayoutOptions(width, firstLineIndentMillimeters: indent, overflow: TextOverflowPolicy.BreakGraphemes)));

    private static string LineRanges(ParagraphLayoutResult result) => string.Join(",", result.Lines.Select(line =>
        string.Create(CultureInfo.InvariantCulture, $"{line.SourceStart}:{line.SourceLength}")));

    [Theory]
    [InlineData("latin", "25b9f7f931229f06bd85db95cea8ca105675fb47e0d32132e661645ade1f2649")]
    [InlineData("mixed-widths", "c884c756c118fe193349bc37560770a63033b00c049571e5e867b57d9f34cce4")]
    [InlineData("ligatures", "5e36773bb09832d294993266030eec73b4caf510587d7840de8d8e681c3f2c1e")]
    [InlineData("arabic", "21f2dd354287da4e6b3cfdf2c0b666a7c223f5931a264935eac7bad151a76fe4")]
    [InlineData("spaced", "f116c00976a1ddfa57098b56dc8c84e82a5ab8e4363c59c8576304e3d8c31dfe")]
    [InlineData("indent", "59894bc919041c5fa6e4f4fae46670b8e2673c395cc7672823c84dab29e79dfc")]
    [InlineData("thai", "4fdf5d18305498f9fc4b3342de4af26900fbd7de7f5325460b0748f91981d181")]
    [InlineData("after-words", "ccdfd36ed2725c3d50caa0eae32dbd1123df5fe8ef2b13cdce1560d6325394f5")]
    public void LineChoicesAreUnchanged(string caseName, string expected)
    {
        var result = caseName switch
        {
            "latin" => Layout(new string('a', 600), Noto.Value, 30),
            "mixed-widths" => Layout(Repeat("WiMl", 150), Noto.Value, 30),
            "ligatures" => Layout(Repeat("ffi", 120), DejaVu.Value, 20),
            "arabic" => Layout(Repeat("لا", 120), DejaVu.Value, 20),
            "spaced" => Layout(new string('a', 400), Noto.Value.WithLetterSpacing(0.5), 30),
            "indent" => Layout(new string('a', 400), Noto.Value, 30, indent: 10),
            "thai" => Layout(Repeat("ภาษาไทย", 60), Thai.Value, 30),
            "after-words" => Layout("word word " + new string('a', 300) + " end", Noto.Value, 30),
            _ => throw new ArgumentOutOfRangeException(nameof(caseName)),
        };
        var ranges = LineRanges(result);
        var actual = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(ranges))).ToLowerInvariant();
        Assert.True(expected == actual, $"Expected {expected}; actual {actual}. Line ranges: {ranges}");
    }
    [Fact]
    public void NegativeTrackingKeepsTheLaterFittingCut()
    {
        var text = new string('W', 32) + new string('i', 64) + new string('W', 32);
        var result = Layout(text, Noto.Value.WithLetterSpacing(-1.99), 30);
        Assert.Equal("0:115,115:13", LineRanges(result));
    }

    [Fact]
    public void EmergencySearchPreservesProbesForANonmonotonicWidth()
    {
        var firstLineLengths = new List<int>();
        var face = Noto.Value.FontFallback[0].Faces[0];
        var style = new TextStyle(Noto.Value.FontFallback[0], 2.5);
        var shaper = new FixtureShaper(request =>
        {
            if (request.SourceStart == 0) firstLineLengths.Add(request.Text.Length);
            // Every character measures 1 mm. At the exceptional cut, reshaping
            // makes the 20 characters fit together even though cuts 11–19 overflow.
            var advance = request.SourceStart == 0 && request.Text.Length == 20 ? 200 : 400;
            return Enumerable.Range(0, request.Text.Length).Select(i => new TextShapedGlyph(
                (uint)face.NominalGlyph('a'), request.SourceStart + i, 1, advance, 0, 0, 0, null, false)).ToArray();
        });
        var result = new ParagraphLayoutEngine(shaper).Layout(new ParagraphLayoutRequest(new string('a', 80), style,
            new ParagraphLayoutOptions(10, overflow: TextOverflowPolicy.BreakGraphemes)));
        Assert.Equal(20, result.Lines[0].SourceLength);
        // The first call measures the full item; the rest are the original trial order.
        Assert.Equal(new[] { 80, 40, 20, 30, 25, 22, 21, 21, 22, 23, 24 }, firstLineLengths);
    }

    private sealed class FixtureShaper(Func<TextShapeRequest, IReadOnlyList<TextShapedGlyph>> shape) : ITextShaper
    {
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default) =>
            new(request, shape(request), "EmergencyBreakTests");
    }
}
