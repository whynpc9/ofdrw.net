using System.IO.Compression;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using static Ofdrw.Net.Core.Fonts.OpenTypeFace;
namespace Ofdrw.Net.Packaging.Tests;

public sealed class FontSubsetTests
{
    private static readonly Lazy<byte[]> Latin = new(() => File.ReadAllBytes(Font("NotoSans-Regular.ttf")));
    private static readonly Lazy<byte[]> Chinese = new(() => File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
    private static string Font(string name)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln"))) directory = directory.Parent;
        return Path.Combine(directory!.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts", name);
    }
    private static OfdDocumentPackage Package(byte[] bytes, string text = "AV office ffi é e\u0301")
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Source", Data = bytes });
        var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 };
        page.Elements.Add(new OfdTextElement { FontResourceId = "10", FontName = "Source", Text = text, WidthMillimeters = 180, HeightMillimeters = 20 });
        package.Pages.Add(page); return package;
    }
    private static async Task<(OfdDocumentPackage Package, OfdPackageWriteResult Report, byte[] Zip)> Save(OfdDocumentPackage package)
    {
        using var output = new MemoryStream(); var report = await new OfdPackageWriter().WriteWithResultAsync(package, output);
        var bytes = output.ToArray(); output.Position = 0;
        return (await new OfdReader().ReadAsync(output), report, bytes);
    }
    [Fact]
    public async Task RealUnicodeSubsetPreservesMappingsWidthsStylesAndSourceBytes()
    {
        var source = Latin.Value.ToArray(); var before = source.ToArray(); var package = Package(source);
        var saved = await Save(package); var result = Assert.Single(saved.Report.FontEmbedding);
        Assert.True(result.IsSubset); Assert.True(result.PayloadBytes < result.SourceBytes / 2);
        Assert.Equal(before, source); Assert.Same(source, package.Fonts[0].Data);
        var original = new OpenTypeFace(source); var subset = new OpenTypeFace(saved.Package.Fonts[0].Data);
        var a = new OpenTypeCmap(original); var b = new OpenTypeCmap(subset);
        foreach (var scalar in Scalars(package.Pages[0].Elements.OfType<OfdTextElement>().Single().Text)) Assert.Equal(a.Glyph(scalar), b.Glyph(scalar));
        Assert.Equal(original.Table("hmtx"), subset.Table("hmtx")); Assert.Equal(original.Table("hhea"), subset.Table("hhea"));
        Assert.Equal(original.Table("maxp"), subset.Table("maxp")); Assert.Equal(original.Style, subset.Style);
        Assert.Equal(0xB1B0AFBAu, Checksum(saved.Package.Fonts[0].Data));
        Assert.Equal(0, b.Glyph('Z'));
        Assert.Equal(original.Table("GSUB"), subset.Table("GSUB")); Assert.Equal(original.Table("GPOS"), subset.Table("GPOS"));
        Assert.Equal(original.Table("name").Length > 0, subset.Table("name").Length > 0);
        Assert.NotEqual(original.Table("name"), subset.Table("name"));
    }
    [Fact]
    public async Task RealChineseSubsetIncludesAllPagesCanonicalAndVariationGlyphClosure()
    {
        var package = Package(Chinese.Value, "中文 AV e\u0301");
        var page = new OfdPage { Index = 1, WidthMillimeters = 210, HeightMillimeters = 297 };
        page.Elements.Add(new OfdTextElement { FontResourceId = "10", Text = "第二页增量字形复用" }); package.Pages.Add(page);
        var saved = await Save(package); var report = Assert.Single(saved.Report.FontEmbedding);
        Assert.True(report.IsSubset); Assert.True(report.PayloadBytes < report.SourceBytes / 10);
        var original = new OpenTypeFace(Chinese.Value); var subset = new OpenTypeFace(saved.Package.Fonts[0].Data);
        var cmap = new OpenTypeCmap(subset); var sourceCmap = new OpenTypeCmap(original);
        foreach (var scalar in Scalars("中文 AV é e\u0301第二页增量字形复用")) Assert.Equal(sourceCmap.Glyph(scalar), cmap.Glyph(scalar));
        Assert.Equal(sourceCmap.VariationSequences, cmap.VariationSequences);
        foreach (var gid in sourceCmap.VariationGlyphs) Assert.True(GlyphLength(subset, gid) > 0);
    }
    [Fact]
    public async Task SharedContentAliasesUseOnePayloadWithUnionOfTheirText()
    {
        var package = Package(Latin.Value, "Alpha");
        package.Fonts.Add(new OfdFontResource { Id = "11", FontName = "OtherName", Bold = true, Data = Latin.Value.ToArray(), FileName = "alias.otf" });
        package.Pages[0].Elements.Add(new OfdTextElement { FontResourceId = "11", Text = "Zebra", Weight = 700 });
        var saved = await Save(package); var result = Assert.Single(saved.Report.FontEmbedding); Assert.Equal(2, result.ResourceCount);
        Assert.Equal(2, saved.Package.Fonts.Count); Assert.True(saved.Package.Fonts.Single(font => font.Bold).Bold);
        Assert.Equal(saved.Package.Fonts[0].FileName, saved.Package.Fonts[1].FileName);
        using var zip = new ZipArchive(new MemoryStream(saved.Zip)); Assert.Single(zip.Entries.Where(entry => entry.FullName.EndsWith(".ttf")));
        var cmap = new OpenTypeCmap(new OpenTypeFace(saved.Package.Fonts[0].Data)); Assert.NotEqual(0, cmap.Glyph('Z'));
    }
    [Fact]
    public async Task SameNamesDifferentContentDoNotMerge()
    {
        var package = Package(Latin.Value, "Alpha"); var second = new OpenTypeFace(Latin.Value);
        var head = second.Table("head"); Put16(head, 44, U16(head, 44) | 1);
        package.Fonts.Add(new OfdFontResource { Id = "11", FontName = "Source", Bold = true, Data = second.Build() });
        package.Pages[0].Elements.Add(new OfdTextElement { FontResourceId = "11", Text = "Beta", Weight = 700 });
        var saved = await Save(package); Assert.Equal(2, saved.Report.FontEmbedding.Count);
        Assert.NotEqual(saved.Package.Fonts[0].FileName, saved.Package.Fonts[1].FileName);
    }
    [Fact]
    public async Task MissingGlyphExplicitlyFailsBeforeWritingAnyBytes()
    {
        var package = Package(Latin.Value, "中文"); using var output = new MemoryStream();
        var exception = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, output));
        Assert.Contains("U+4E2D", exception.Message); Assert.Equal(0, output.Length);
    }
    [Fact]
    public async Task ExplicitFallbackBindingSucceedsAndUnboundResourceDoesNotReplaceItByName()
    {
        var package = Package(Latin.Value, "中文");
        package.Fonts.Add(new OfdFontResource { Id = "11", FontName = "Source", Data = Chinese.Value });
        package.Pages[0].Elements.OfType<OfdTextElement>().Single().FontResourceId = "11";
        var saved = await Save(package); var selected = saved.Package.Fonts.Single(font => font.Id == "11");
        Assert.NotEqual(0, new OpenTypeCmap(new OpenTypeFace(selected.Data)).Glyph('中'));
    }
    [Theory]
    [InlineData("raw")][InlineData("source")][InlineData("clip")][InlineData("preserved")][InlineData("template")]
    public async Task UnmodeledContentPreservesFullFontWithObservableReason(string kind)
    {
        var package = Package(Latin.Value);
        switch (kind)
        {
            case "raw": package.Pages[0].Elements.Add(new OfdRawElement { Xml = "<VendorObject Glyphs='999'/>" }); break;
            case "source": package.Pages[0].Elements.OfType<OfdTextElement>().Single().SourceXml = "<TextObject><CGTransform><Glyphs>999</Glyphs></CGTransform></TextObject>"; break;
            case "clip": package.Pages[0].Elements.OfType<OfdTextElement>().Single().ClippingXml = "<Clips/>"; break;
            case "preserved": package.PreservedEntries["Extensions/private.bin"] = [1, 2, 3]; break;
            case "template": package.Pages[0].PreservedPageElements.Add("<VendorPageData Glyph='999'/>"); break;
        }
        var saved = await Save(package); Assert.Equal(Latin.Value, saved.Package.Fonts[0].Data);
        Assert.False(Assert.Single(saved.Report.FontEmbedding).IsSubset); Assert.Contains(saved.Report.Diagnostics, message => message.Contains("unmodeled"));
    }
    [Fact]
    public async Task UnmodeledContentCannotBypassTypedMissingGlyphFailure()
    {
        var package = Package(Latin.Value, "中文"); package.Pages[0].Elements.Add(new OfdRawElement { Xml = "<VendorObject/>" });
        await Assert.ThrowsAsync<InvalidDataException>(() => Save(package));
    }
    [Fact]
    public async Task SubsetRoundTripKeepsOriginalPayloadAndCannotAddMissingCharacters()
    {
        var first = await Save(Package(Latin.Value, "Alpha")); var second = await Save(first.Package);
        Assert.Equal(first.Package.Fonts[0].Data, second.Package.Fonts[0].Data);
        var edited = first.Package.Pages[0].Elements.OfType<OfdTextElement>().Single();
        edited.Text = "Zebra"; edited.Runs.Clear(); edited.SourceXml = null;
        await Assert.ThrowsAsync<InvalidDataException>(() => Save(first.Package));
    }
    [Theory]
    [InlineData(0x0002)][InlineData(0x0200)]
    public async Task RestrictedEmbeddingFlagsFail(int flags)
    {
        var face = new OpenTypeFace(Latin.Value); Put16(face.Table("OS/2"), 8, flags);
        await Assert.ThrowsAsync<InvalidDataException>(() => Save(Package(face.Build())));
    }
    [Fact]
    public async Task NoSubsettingFlagKeepsWholeFontButStillValidatesCoverage()
    {
        var face = new OpenTypeFace(Latin.Value); Put16(face.Table("OS/2"), 8, 0x0100); var bytes = face.Build();
        var saved = await Save(Package(bytes)); Assert.Equal(bytes, saved.Package.Fonts[0].Data);
        Assert.Contains("prohibits subsetting", Assert.Single(saved.Report.Diagnostics));
        await Assert.ThrowsAsync<InvalidDataException>(() => Save(Package(bytes, "中文")));
    }
    [Fact]
    public async Task UnsupportedGlyphDependentTablePreservesFullWithReason()
    {
        var face = new OpenTypeFace(Latin.Value); face.Tables["JSTF"] = [0, 1, 0, 0]; var bytes = face.Build();
        var saved = await Save(Package(bytes)); Assert.Equal(bytes, saved.Package.Fonts[0].Data);
        Assert.Contains("JSTF", Assert.Single(saved.Report.Diagnostics));
    }
    [Fact]
    public async Task FullPolicyRetainsExplicitLegacyBehavior()
    {
        var package = Package(Latin.Value); package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        var saved = await Save(package); Assert.Equal(Latin.Value, saved.Package.Fonts[0].Data);
        Assert.Contains("Explicit full", Assert.Single(saved.Report.Diagnostics));
    }
    [Theory]
    [InlineData("bytes")][InlineData("scalars")][InlineData("operations")]
    public async Task SubsetBudgetsActuallyStopWork(string kind)
    {
        var package = Package(Latin.Value);
        if (kind == "bytes") package.Options.FontEmbedding.MaximumFontBytes = 10;
        if (kind == "scalars") package.Options.FontEmbedding.MaximumUsedScalars = 2;
        if (kind == "operations") package.Options.FontEmbedding.MaximumGlyphClosureOperations = 1;
        await Assert.ThrowsAsync<InvalidDataException>(() => Save(package));
    }
    [Fact]
    public void OverlappingTablesAreRejectedBeforeCopyingPayloads()
    {
        var bytes = Latin.Value.ToArray(); var firstOffset = U32(bytes, 20); var firstLength = U32(bytes, 24);
        Put32(bytes, 36, firstOffset); Put32(bytes, 40, firstLength);
        Assert.Throws<InvalidDataException>(() => new OpenTypeFace(bytes));
    }
    [Fact]
    public async Task TtcFaceSelectionIsStableAcrossFullSubsetAndOpaqueContent()
    {
        var path = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(Font("NotoSans-Regular.ttf"))!, "../../../Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/style-collection.ttc"));
        var originalFace = OpenTypeCollection.SelectFace(File.ReadAllBytes(path), 0);
        var styled = new OpenTypeFace(originalFace); Put16(styled.Table("head"), 44, 1);
        var source = Collection(originalFace, styled.Build()); var expected = OpenTypeCollection.SelectFace(source, 1);
        foreach (var opaque in new[] { false, true })
        {
            var package = Package(source, "ABC"); package.Fonts[0].CollectionFaceIndex = 1;
            if (opaque) package.Pages[0].Elements.Add(new OfdRawElement { Xml = "<VendorObject/>" });
            var saved = await Save(package); var face = new OpenTypeFace(saved.Package.Fonts[0].Data);
            Assert.Equal(new OpenTypeFace(expected).Style, face.Style); Assert.Equal(new OpenTypeFace(expected).Table("hmtx"), face.Table("hmtx"));
            if (opaque) Assert.Equal(expected, saved.Package.Fonts[0].Data);
        }
        var invalid = Package(source, "ABC"); invalid.Fonts[0].CollectionFaceIndex = 999;
        await Assert.ThrowsAsync<InvalidDataException>(() => Save(invalid));
    }
    [Fact]
    public void LargeCmapKeepsFormat4AndPreservesGids()
    {
        var map = Enumerable.Range(1, 9000).ToDictionary(scalar => scalar, scalar => scalar);
        var face = new OpenTypeFace(Latin.Value); face.Tables["cmap"] = OpenTypeCmap.Build(map);
        var cmap = new OpenTypeCmap(face); Assert.Equal(9000, cmap.Glyph(9000));
        var table = face.Table("cmap"); Assert.Equal(4, U16(table, (int)U32(table, 8)));
        var uncompressible = Enumerable.Range(1, 9000).ToDictionary(scalar => scalar, scalar => scalar * 2);
        Assert.Throws<NotSupportedException>(() => OpenTypeCmap.Build(uncompressible));
    }
    [Fact]
    public async Task InvalidUtf16AndUnknownVariationSequencesFail()
    {
        await Assert.ThrowsAsync<InvalidDataException>(() => Save(Package(Latin.Value, "\uD800")));
        await Assert.ThrowsAsync<InvalidDataException>(() => Save(Package(Latin.Value, "A\uFE0F")));
    }
    [Fact]
    public async Task ExplicitGlyphSourceXmlKeepsFontEvenWithoutUnicodeCmapCoverage()
    {
        var package = Package(Latin.Value, "中文");
        package.Pages[0].Elements.OfType<OfdTextElement>().Single().SourceXml = "<TextObject><CGTransform CodePosition='0' CodeCount='2' GlyphCount='2'><Glyphs>1 2</Glyphs></CGTransform><TextCode>中文</TextCode></TextObject>";
        var saved = await Save(package); Assert.Equal(Latin.Value, saved.Package.Fonts[0].Data);
        Assert.Contains(saved.Report.Diagnostics, message => message.Contains("FONT_COVERAGE_UNVERIFIED"));
        Assert.Contains("CGTransform", saved.Package.Pages[0].Elements.OfType<OfdTextElement>().Single().SourceXml);
    }
    [Fact]
    public async Task RealNonBmpCmapIsPreservedInOfd()
    {
        var original = new OpenTypeCmap(new OpenTypeFace(Chinese.Value)); Assert.NotEqual(0, original.Glyph(0x107A5));
        var saved = await Save(Package(Chinese.Value, "A\U000107A5B"));
        var cmap = new OpenTypeCmap(new OpenTypeFace(saved.Package.Fonts[0].Data)); Assert.Equal(original.Glyph(0x107A5), cmap.Glyph(0x107A5));
        Assert.Equal("A\U000107A5B", saved.Package.Pages[0].Elements.OfType<OfdTextElement>().Single().Text);
    }
    private static byte[] Collection(params byte[][] faces)
    {
        var header = 12 + faces.Length * 4; var result = new byte[header + faces.Sum(face => face.Length)];
        Put32(result, 0, 0x74746366); Put32(result, 4, 0x00010000); Put32(result, 8, (uint)faces.Length);
        var offset = header;
        for (var index = 0; index < faces.Length; index++)
        {
            var data = faces[index].ToArray(); var count = U16(data, 4);
            for (var table = 0; table < count; table++) Put32(data, 20 + table * 16, U32(data, 20 + table * 16) + (uint)offset);
            Put32(result, 12 + index * 4, (uint)offset); data.CopyTo(result, offset); offset += data.Length;
        }
        return result;
    }
    private static int GlyphLength(OpenTypeFace face, int gid)
    {
        var loca = face.Table("loca"); var format = U16(face.Table("head"), 50);
        return format == 1 ? (int)(U32(loca, (gid + 1) * 4) - U32(loca, gid * 4)) : 2 * (U16(loca, (gid + 1) * 2) - U16(loca, gid * 2));
    }
}
