using System;
using System.IO;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Fonts;

namespace Ofdrw.Net.Core.Tests;

// A1c: byte-level font validation for archival checks. Variants are derived from the open fixtures at test time.
public sealed class OfdFontPayloadInspectorTests
{
    [Fact]
    public void CompleteTrueTypeCoversRequestedTextAndGlyphs()
    {
        var report = Inspect(NotoSans(), "Ofdrw é", 36);
        Assert.Equal((OfdFontPayloadStatus.Parsed, OfdFontPayloadFormat.TrueType, "Parsed"), (report.Status, report.Format, report.Reason));
        Assert.True(report.CoverageChecked);
        Assert.True(report.CanRenderRequest);
        Assert.Empty(report.MissingRequiredTables);
        Assert.Empty(report.MissingOutlineTables);
        Assert.Contains("glyf", report.Tables);
        Assert.True(report.GlyphCount > 36);
        Assert.NotNull(report.EmbeddingPermissions);
    }

    [Fact]
    public void CffOutlinesAreSupported()
    {
        var report = Inspect(File.ReadAllBytes(Repository("tests/Ofdrw.Net.Graphics.SkiaSharp.Tests/testdata/cff/SourceCodePro-Regular.otf")), "Archive");
        Assert.Equal((OfdFontPayloadStatus.Parsed, OfdFontPayloadFormat.OpenTypeCff), (report.Status, report.Format));
        Assert.True(report.CanRenderRequest);
        Assert.Contains("CFF ", report.Tables);
    }

    [Fact]
    public void MissingCharactersAndGlyphIdsAreListedNotThrown()
    {
        var report = Inspect(NotoSans(), "A中", 0, 65000);
        Assert.Equal(OfdFontPayloadStatus.Parsed, report.Status);
        Assert.False(report.CanRenderRequest);
        Assert.Equal(new[] { 0x4E2D }, report.MissingScalars);
        Assert.Equal(new[] { 0, 65000 }, report.MissingGlyphIds);
    }

    [Fact]
    public void MissingRequiredTableIsReportedFromTheBytes()
    {
        var face = new OpenTypeFace(NotoSans());
        face.Tables.Remove("OS/2"); face.Tables.Remove("post");
        var report = Inspect(face.Build(), "A");
        Assert.Equal(OfdFontPayloadStatus.Parsed, report.Status);
        Assert.Equal(new[] { "OS/2", "post" }, report.MissingRequiredTables);
        Assert.Null(report.EmbeddingPermissions);
        Assert.False(report.CanRenderRequest);
    }

    [Fact]
    public void MissingOutlineTableIsReportedWithoutCoverage()
    {
        var face = new OpenTypeFace(NotoSans());
        face.Tables.Remove("glyf");
        var report = Inspect(face.Build(), "A");
        Assert.Equal(("OutlineTablesMissing", false), (report.Reason, report.CoverageChecked));
        Assert.Equal(new[] { "glyf" }, report.MissingOutlineTables);
    }

    [Theory]
    [InlineData("truncated", OfdFontPayloadStatus.Malformed, "MalformedTableDirectory")]
    [InlineData("garbage", OfdFontPayloadStatus.Unsupported, "UnrecognizedFontFormat")]
    [InlineData("empty", OfdFontPayloadStatus.Unsupported, "UnrecognizedFontFormat")]
    [InlineData("woff", OfdFontPayloadStatus.Unsupported, "WoffNotSupported")]
    [InlineData("collection", OfdFontPayloadStatus.Unsupported, "FontCollectionNotSupported")]
    [InlineData("corrupt-glyf", OfdFontPayloadStatus.Parsed, "Parsed")]
    public void UnusableBytesProduceAStatusNotAnException(string variant, OfdFontPayloadStatus status, string reason)
    {
        var bytes = NotoSans();
        bytes = variant switch
        {
            "truncated" => bytes.Take(4096).ToArray(),
            "garbage" => Enumerable.Range(0, 512).Select(x => (byte)(x * 37)).ToArray(),
            "empty" => Array.Empty<byte>(),
            "woff" => new byte[] { 0x77, 0x4F, 0x46, 0x46 }.Concat(new byte[64]).ToArray(),
            "collection" => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "testdata", "style-collection.ttc")),
            _ => CorruptGlyph(bytes, 'A')
        };
        var report = Inspect(bytes, "A");
        Assert.Equal((status, reason), (report.Status, report.Reason));
        if (variant == "corrupt-glyf")
        {
            Assert.NotEmpty(report.UnreadableGlyphIds);
            Assert.False(report.CanRenderRequest);
        }
    }

    [Fact]
    public void BudgetsAndCancellationAreHonoured()
    {
        var bytes = NotoSans();
        var small = new OfdFontPayloadRequest { MaxPayloadBytes = 1024 };
        Assert.Equal(OfdFontPayloadStatus.BudgetExceeded, OfdFontPayloadInspector.Inspect(bytes, small).Status);
        var checks = new OfdFontPayloadRequest { MaxGlyphChecks = 1 };
        checks.UnicodeScalars.Add('A'); checks.UnicodeScalars.Add('B');
        Assert.Equal(OfdFontPayloadStatus.BudgetExceeded, OfdFontPayloadInspector.Inspect(bytes, checks).Status);
        var cmap = new OfdFontPayloadRequest { MaxCmapWork = 4 };
        cmap.UnicodeScalars.Add('A');
        Assert.Equal(("CmapBudgetExceeded", OfdFontPayloadStatus.BudgetExceeded),
            (OfdFontPayloadInspector.Inspect(bytes, cmap).Reason, OfdFontPayloadInspector.Inspect(bytes, cmap).Status));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfdFontPayloadInspector.Inspect(bytes, new OfdFontPayloadRequest(), cancelled.Token));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfdFontPayloadInspector.Inspect(bytes, new OfdFontPayloadRequest { MaxPayloadBytes = 0 }));
    }

    [Fact]
    public void InputBytesAreNotModified()
    {
        var bytes = NotoSans();
        var copy = (byte[])bytes.Clone();
        Inspect(bytes, "Ofdrw");
        Assert.Equal(copy, bytes);
    }

    private static OfdFontPayloadReport Inspect(byte[] bytes, string text, params int[] glyphs)
    {
        var request = new OfdFontPayloadRequest();
        foreach (var scalar in OpenTypeFace.Scalars(text)) request.UnicodeScalars.Add(scalar);
        foreach (var glyph in glyphs) request.GlyphIds.Add(glyph);
        return OfdFontPayloadInspector.Inspect(bytes, request);
    }

    // Overwrite the glyph record of one character with an invalid composite header, keeping the table directory valid.
    private static byte[] CorruptGlyph(byte[] bytes, char character)
    {
        var face = new OpenTypeFace(bytes);
        var gid = new OpenTypeCmap(face).Glyph(character);
        var head = face.Table("head", 54);
        var loca = face.Table("loca");
        var offset = OpenTypeFace.U16(head, 50) == 0 ? OpenTypeFace.U16(loca, gid * 2) * 2 : (int)OpenTypeFace.U32(loca, gid * 4);
        var glyf = face.Table("glyf");
        glyf[offset] = 0x80; glyf[offset + 1] = 0x00; // numberOfContours = -32768
        return face.Build();
    }

    private static byte[] NotoSans() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "testdata", "NotoSans-Regular.ttf"));

    private static string Repository(string relative)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return Path.Combine(root.FullName, relative);
    }
}
