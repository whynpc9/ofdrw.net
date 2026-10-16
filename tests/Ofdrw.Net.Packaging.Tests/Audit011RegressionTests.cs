using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using static Ofdrw.Net.Core.Fonts.OpenTypeFace;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Audit011RegressionTests
{
    private static string Root
    {
        get
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (!File.Exists(Path.Combine(directory!.FullName, "Ofdrw.Net.sln"))) directory = directory.Parent;
            return directory.FullName;
        }
    }
    private static byte[] Font => File.ReadAllBytes(Path.Combine(Root, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));
    private static OpenTypeFace Face(byte[] cmap)
    { var face = new OpenTypeFace(Font); face.Tables["cmap"] = cmap; return face; }

    // One format 12 physical group and optionally a format 14 with two selectors
    // pointing to the same default/nondefault arrays. All data stays tiny.
    private static byte[] Aliases(int unicodeAliases, int variationAliases = 0)
    {
        var records = unicodeAliases + variationAliases;
        var header = 4 + records * 8;
        var bytes = new byte[header + 28 + (variationAliases == 0 ? 0 : 49)];
        Put16(bytes, 2, records);
        for (var i = 0; i < records; i++)
        {
            Put16(bytes, 4 + i * 8, i < unicodeAliases ? 3 : 0);
            Put16(bytes, 6 + i * 8, i < unicodeAliases ? 10 : 5);
            Put32(bytes, 8 + i * 8, (uint)(header + (i < unicodeAliases ? 0 : 28)));
        }
        Put16(bytes, header, 12); Put32(bytes, header + 4, 28); Put32(bytes, header + 12, 1);
        Put32(bytes, header + 16, 65); Put32(bytes, header + 20, 65); Put32(bytes, header + 24, 1);
        if (variationAliases > 0)
        {
            var start = header + 28;
            Put16(bytes, start, 14); Put32(bytes, start + 2, 49); Put32(bytes, start + 6, 2);
            for (var i = 0; i < 2; i++)
            {
                var record = start + 10 + i * 11;
                bytes[record + 1] = 0xFE; bytes[record + 2] = (byte)i;
                Put32(bytes, record + 3, 32); Put32(bytes, record + 7, 40);
            }
            Put32(bytes, start + 32, 1); bytes[start + 38] = 65;
            Put32(bytes, start + 40, 1); bytes[start + 46] = 66; Put16(bytes, start + 47, 2);
        }
        return bytes;
    }

    [Fact]
    public void Audit011_SharedOffset_ChargesPhysicalFormat12Once()
    {
        var budget = new FontParseBudget(49);
        var cmap = new OpenTypeCmap(Face(Aliases(20)), budget);
        Assert.Equal(1, cmap.Glyph('A')); Assert.Equal(0, cmap.Glyph('B'));
        Assert.Equal(49, budget.Consumed);
    }

    [Fact]
    public void Audit011_VariationAliases_ChargeCopyAndSharedArraysOnce()
    {
        var budget = new FontParseBudget(104);
        var cmap = new OpenTypeCmap(Face(Aliases(1, 20)), budget);
        Assert.Equal(103, budget.Consumed); // 21 records + 29 format12 + 49 bytes + 2 selectors + 2 arrays
        Assert.Equal(new[] { 2 }, cmap.VariationGlyphs);

        // Use a separate finite budget for lookup because the exact parse limit
        // intentionally leaves insufficient room for repeated coverage queries.
        var lookup = new OpenTypeCmap(Face(Aliases(1, 20)));
        Assert.True(lookup.SupportsVariation('B', 0xFE01));
        Assert.False(lookup.SupportsVariation('C', 0xFE01));
    }

    [Fact]
    public void Audit011_SmallBudget_RejectsBeforeCopyOrExpansionAndRemainsExhausted()
    {
        var budget = new FontParseBudget(30);
        var failure = Assert.Throws<InvalidDataException>(() => new OpenTypeCmap(Face(Aliases(1, 1)), budget));
        Assert.Contains("MaximumCmapParseWork", failure.Message);
        Assert.Equal(30, budget.Consumed); // second encoding charged; variation copy has not started
        Assert.True(budget.IsExhausted);
        Assert.Throws<InvalidDataException>(() => budget.Spend(0));
    }

    [Fact]
    public async Task Audit011_CancelDuringParse_PropagatesThroughSubsetAndWriter()
    {
        var face = Face(Aliases(1, 1));
        var canceled = new CancellationToken(true);
        Assert.Throws<OperationCanceledException>(() => new OpenTypeCmap(face, new FontParseBudget(cancellationToken: canceled)));
        Assert.Throws<OperationCanceledException>(() => TrueTypeSubset.Create(face, new[] { 65 }, canceled, 100, out _, out _));
        var package = Package(face.Build()); using var output = new MemoryStream();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new OfdPackageWriter().WriteAsync(package, output, canceled));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task Audit011_SmallBudget_EmbeddingFailsBeforeWritingAndSuccessfulSubsetReusesCmap()
    {
        var bytes = Face(Aliases(1)).Build();
        var package = Package(bytes); package.Options.FontEmbedding.MaximumCmapParseWork = 29;
        using var failed = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, failed));
        Assert.Equal(0, failed.Length);
        package.Options.FontEmbedding.MaximumCmapParseWork = 30;
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0; var saved = await new OfdReader().ReadAsync(output);
        Assert.Equal(1, new OpenTypeCmap(new OpenTypeFace(saved.Fonts[0].Data)).Glyph('A'));
    }

    [Fact]
    public async Task Audit011_MergePreservesNondefaultCmapLimitAndWriterEnforcesIt()
    {
        var first = Package(Face(Aliases(1)).Build());
        first.Options.FontEmbedding.MaximumCmapParseWork = 29;
        var second = new OfdDocumentPackage(); second.Pages.Add(new OfdPage());
        second.Options.FontEmbedding.MaximumCmapParseWork = 1000;
        var merged = Ofdrw.Net.Layout.Editing.OfdDocumentMerger.Merge(new[] { first, second });
        Assert.Equal(29, merged.Options.FontEmbedding.MaximumCmapParseWork);
        Assert.NotSame(first.Options.FontEmbedding, merged.Options.FontEmbedding);
        using var output = new MemoryStream();
        var failure = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(merged, output));
        Assert.Contains("MaximumCmapParseWork", failure.Message); Assert.Equal(0, output.Length);
        merged.Options.FontEmbedding.MaximumCmapParseWork = 1000;
        Assert.Equal(29, first.Options.FontEmbedding.MaximumCmapParseWork);
    }

    private static OfdDocumentPackage Package(byte[] bytes)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Audit011", Data = bytes });
        package.Pages.Add(new OfdPage { Elements = { new OfdTextElement { FontResourceId = "10", Text = "A" } } });
        return package;
    }

    [Fact]
    public void Audit011_RealFontControls_RecordCorpusCostsAndFiniteDefaultMargin()
    {
        var costs = new List<string>();
        foreach (var path in Directory.GetFiles(Path.Combine(Root, "e2e"), "*.ttf", SearchOption.AllDirectories)
                     .Where(path => path.Contains("FontSubset.E2E") || path.Contains("Converter.Pdf.E2E")))
        {
            var budget = new FontParseBudget();
            var cmap = new OpenTypeCmap(new OpenTypeFace(File.ReadAllBytes(path)), budget);
            Assert.NotEqual(0, cmap.Glyph('A'));
            Assert.True(budget.Consumed < FontParseBudget.DefaultMaximumWork / 8);
            costs.Add(Path.GetRelativePath(Root, path) + " " + budget.Consumed);
        }
        Assert.NotEmpty(costs);
        Directory.CreateDirectory(Path.Combine(Root, "artifacts/plan-011"));
        File.WriteAllLines(Path.Combine(Root, "artifacts/plan-011/corpus-costs.txt"), costs);
    }

    [Fact]
    public void Audit011_RealFontControls_TtcAndFormat12Authority()
    {
        var collection = File.ReadAllBytes(Path.Combine(Root, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/style-collection.ttc"));
        var count = (int)U32(collection, 8); Assert.True(count > 0);
        for (var i = 0; i < count; i++)
        {
            var face = new OpenTypeFace(OpenTypeCollection.SelectFace(collection, i));
            Assert.NotEqual(0, new OpenTypeCmap(face, new FontParseBudget()).Glyph('A'));
        }
        var table = OpenTypeCmap.Build(new Dictionary<int, int> { [65] = 1 });
        var offset4 = (int)U32(table, 8);
        // An invalid unused BMP segment cannot override authoritative format12.
        Put16(table, offset4 + 16 + U16(table, offset4 + 6), 1000);
        var authoritative = new OpenTypeCmap(Face(table), new FontParseBudget());
        Assert.Equal(1, authoritative.Glyph('A'));
        Assert.Equal(0, authoritative.Glyph('B'));
    }

    [Fact]
    public void Audit011_RealFontControls_Format4AndSymbolRetainAuthority()
    {
        var original = new OpenTypeCmap(new OpenTypeFace(Font));
        var table = OpenTypeCmap.Build(new Dictionary<int, int> { [65] = original.Glyph(65) });
        // Keep only the first (format4) encoding, then exercise symbol fallback.
        Put16(table, 2, 1);
        var bmp = new OpenTypeCmap(Face(table)); Assert.Equal(original.Glyph(65), bmp.Glyph(65));
        Put16(table, 6, 0);
        var symbol = new OpenTypeCmap(Face(table)); Assert.True(symbol.IsSymbol); Assert.Equal(bmp.Glyph(65), symbol.Glyph(65));
    }
}
