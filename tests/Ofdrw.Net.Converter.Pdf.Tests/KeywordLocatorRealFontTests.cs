using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Geometry;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[Collection("Trusted font geometry")]
public sealed class KeywordLocatorRealFontTests(Xunit.Abstractions.ITestOutputHelper output)
{
    [RealTrustedFontsTheory]
    [InlineData("latin", "identity", 17, 28.132, 19.552, 31)]
    [InlineData("latin", "rotation", 27, 25, 29.868, 27.552)]
    [InlineData("latin", "shear", 19.566, 28.132, 23.552, 31)]
    [InlineData("cjk", "identity", 17.016, 28.068, 19.416, 31)]
    [InlineData("cjk", "rotation", 27, 25.016, 29.932, 27.416)]
    [InlineData("cjk", "shear", 19.55, 28.068, 23.416, 31)]
    public async Task ActualArchiveReaderAndQ1ProviderHaveIndependentAffineExpectations(string face, string matrix,
        double left, double top, double right, double bottom)
    {
        var bytes = Font(face); var source = await Read(bytes, face, matrix, 1);
        var before = ((OfdTextElement)source.Pages[0].Elements[0]).SourceXml;
        var provider = new ObservedProvider(bytes); var locator = new OfdKeywordLocator(provider);
        var text = face == "cjk" ? "AiW 中文𠀋" : "AiW ";
        var result = locator.Find(source, text); var hit = Assert.Single(result.Hits);
        Assert.True(result.IsComplete); Assert.True(hit.CanReplace);
        Assert.Equal(face == "cjk" ? 7 : 4, hit.Fragments.Count);
        var a = hit.Fragments[0]; var box = a.InkQuadrilateralMillimeters!.BoundsMillimeters;
        Assert.Equal(left, box.Left, 9); Assert.Equal(top, box.Top, 9); Assert.Equal(right, box.Right, 9); Assert.Equal(bottom, box.Bottom, 9);
        Assert.Equal(OfdGlyphBoundsKind.TransformedFontGlyphBounds, hit.BoundsKind);
        Assert.Null(hit.Fragments[3].InkQuadrilateralMillimeters); Assert.Equal(5, hit.Fragments[3].SourceUtf16Length);
        Assert.Equal(face == "cjk" ? 34u : 36u, a.GlyphId);
        Assert.Equal(face == "cjk" ? 74u : 76u, hit.Fragments[1].GlyphId);
        Assert.Equal(face == "cjk" ? 56u : 58u, hit.Fragments[2].GlyphId);
        if (face == "cjk")
        {
            Assert.Equal(59477u, hit.Fragments[6].GlyphId); Assert.Equal(2, hit.Fragments[6].LogicalUtf16Length);
            Assert.Equal(2, hit.Fragments[6].SourceUtf16Length);
            if (matrix == "identity")
            {
                Assert.Equal(24.044, hit.Fragments[3].BaselineMillimeters.X, 9);
                Assert.Equal(32.94, hit.Fragments[6].BaselineMillimeters.X, 9);
                Assert.Equal(31, hit.Fragments[6].BaselineMillimeters.Y, 9);
            }
        }
        Assert.Equal(before, ((OfdTextElement)source.Pages[0].Elements[0]).SourceXml);
        Assert.Equal(Digest(bytes), Digest(source.Fonts[0].Data));
        Assert.Equal(1, result.Statistics.FontHashOperations); Assert.Equal(1, result.Statistics.Provider.FaceParses);
        Assert.Equal(bytes.Length, result.Statistics.CopiedFontBytes); Assert.Equal(bytes.Length, result.Statistics.ComparedFontBytes);
        Assert.Equal(1, provider.Opens); Assert.Equal(1, provider.Session!.StreamsOpened); Assert.Equal(1, provider.Session.StreamsDisposed);
        output.WriteLine(JsonSerializer.Serialize(new { face, matrix, result.Statistics, provider.Session.BackendAddAttempts,
            provider.Session.BackendGlyphCalls, provider.Session.ReadCalls, provider.Session.ReadBytes, provider.Session.SeekCalls }));
    }

    [RealTrustedFontsTheory]
    [InlineData("latin", 1)]
    [InlineData("latin", 10)]
    [InlineData("latin", 100)]
    [InlineData("cjk", 1)]
    [InlineData("cjk", 10)]
    [InlineData("cjk", 100)]
    public async Task ActualLocatorReusesOneOriginalFontAcrossManyReadObjects(string face, int objects)
    {
        var bytes = Font(face); var source = await Read(bytes, face, "identity", objects);
        var provider = new ObservedProvider(bytes); var result = new OfdKeywordLocator(provider).Find(source, "A");
        Assert.True(result.IsComplete); Assert.Equal(objects, result.Hits.Count);
        Assert.Equal(1, provider.Opens); Assert.Equal(1, result.Statistics.FontHashOperations);
        Assert.Equal(1, result.Statistics.Provider.FaceParses);
        Assert.Equal(face == "cjk" ? 7 * objects : 4 * objects, result.Statistics.Provider.GlyphRequests);
        Assert.Equal(face == "cjk" ? 7 : 4, provider.Session!.BackendGlyphCalls);
        Assert.Equal(1, provider.Session.StreamsDisposed);
        Assert.Equal(objects, result.Statistics.ReturnedFragments);
        output.WriteLine(JsonSerializer.Serialize(new { face, objects, result.Statistics, provider.Session.BackendAddAttempts,
            provider.Session.BackendGlyphCalls, provider.Session.ReadCalls, provider.Session.ReadBytes, provider.Session.SeekCalls,
            note = "Actual public Locator -> archive Reader -> original Q1 provider; no timing/RSS or visual claim." }));
    }

    [RealTrustedFontsTheory]
    [InlineData("missing")]
    [InlineData("composite")]
    public async Task ActualReaderAndQ1FontFailuresDoNotBecomeCompleteNoHit(string kind)
    {
        var bytes = Font("latin"); var source = await Read(bytes, "latin", "identity", 1, kind == "missing" ? "中" : "é");
        var result = new OfdKeywordLocator(new ObservedProvider(bytes)).Find(source, "absent",
            new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete));
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
        Assert.Equal("UnsupportedFontGeometry", Assert.Single(result.Diagnostics).Code);
    }

    [RealTrustedFontsTheory]
    [InlineData("latin", "identity", 1)]
    [InlineData("latin", "rotation", 10)]
    [InlineData("latin", "shear", 10)]
    [InlineData("cjk", "identity", 10)]
    [InlineData("cjk", "rotation", 1)]
    [InlineData("cjk", "shear", 10)]
    public async Task PublicWriterDefaultIndentedOutputReadsAndLocatesWithOriginalOflFont(string face, string matrix, int objects)
    {
        var bytes = Font(face); var source = await Read(bytes, face, matrix, objects);
        var keyword = face == "cjk" ? "AiW 中文𠀋" : "AiW ";
        var before = new OfdKeywordLocator(new ObservedProvider(bytes)).Find(source, keyword);
        using var destination = new MemoryStream();
        var write = await new Ofdrw.Net.Packaging.OfdPackageWriter().WriteWithResultAsync(source, destination);
        destination.Position = 0; var roundtrip = await new OfdReader().ReadAsync(destination);
        var after = new OfdKeywordLocator(new ObservedProvider(bytes)).Find(roundtrip, keyword);
        Assert.True(before.IsComplete); Assert.True(after.IsComplete); Assert.Equal(objects, after.Hits.Count);
        Assert.Equal(before.Hits.SelectMany(h => h.Fragments).Select(f => (f.BaselineMillimeters.X, f.BaselineMillimeters.Y, f.GlyphId)),
            after.Hits.SelectMany(h => h.Fragments).Select(f => (f.BaselineMillimeters.X, f.BaselineMillimeters.Y, f.GlyphId)));
        Assert.Equal(Digest(bytes), Digest(roundtrip.Fonts[0].Data)); Assert.False(write.SignaturesInvalidated);
        Assert.Contains("\n  <ofd:Area>", Encoding.UTF8.GetString(roundtrip.PreservedEntries["Doc_0/Pages/Page_0/Content.xml"]));
        output.WriteLine(JsonSerializer.Serialize(new { face, matrix, objects, stage = "public-default-writer-fresh-reader-public-locator", after.Statistics }));
    }

    [RealTrustedFontsTheory]
    [InlineData("latin")]
    [InlineData("cjk")]
    public async Task CanonicalEmptyKeepsOriginalOflResourceIdentityWithoutFaceOrGlyphCalls(string face)
    {
        var bytes = Font(face); var source = await Read(bytes, face, "identity", 1, overrideText: "");
        var provider = new ObservedProvider(bytes);
        var before = new OfdKeywordLocator(provider).Find(source, "A");
        using var destination = new MemoryStream();
        await new Ofdrw.Net.Packaging.OfdPackageWriter().WriteAsync(source, destination);
        destination.Position = 0; var reread = await new OfdReader().ReadAsync(destination);
        var after = new OfdKeywordLocator(provider).Find(reread, "A");
        Assert.True(before.IsComplete); Assert.True(after.IsComplete); Assert.Empty(after.Hits); Assert.Empty(after.Diagnostics);
        Assert.Equal("100", ((OfdTextElement)reread.Pages[0].Elements[0]).ObjectId);
        Assert.Equal(0, provider.Opens); Assert.Null(provider.Session);
        Assert.Equal(0, after.Statistics.ReturnedFragments); Assert.Equal(0, after.Statistics.ReturnedProviderClusters);
        Assert.Equal(0, after.Statistics.ReservedScalarGlyphs); Assert.Equal(0, after.Statistics.Provider.FaceParses);
        Assert.Equal(0, after.Statistics.Provider.GlyphRequests); Assert.Equal(1, after.Statistics.FontHashOperations);
        Assert.Equal(bytes.Length, after.Statistics.CopiedFontBytes); Assert.Equal(bytes.Length, after.Statistics.ComparedFontBytes);
        Assert.Equal(Digest(bytes), Digest(reread.Fonts[0].Data));
        output.WriteLine(JsonSerializer.Serialize(new { face, stage = "canonical-empty-default-writer-readback-resource-identity-only", after.Statistics }));
    }

    private static byte[] Font(string face)
    {
        var bytes = TrustedFontFixtures.Read(face);
        Assert.Equal(face == "cjk" ? "3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa" :
            "b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5", Digest(bytes));
        return bytes;
    }
    private static string Digest(byte[] value) => Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();
    private sealed class ObservedProvider(byte[] bytes) : IOfdGlyphGeometryProvider
    {
        internal int Opens; internal OfdTrustedEmbeddedFontGeometryProvider.Session? Session;
        public IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits limits, CancellationToken token = default)
        {
            Opens++; var provider = new OfdTrustedEmbeddedFontGeometryProvider([new(Digest(bytes), bytes.Length)]);
            Session = (OfdTrustedEmbeddedFontGeometryProvider.Session)provider.OpenSession(limits, token); return Session;
        }
    }
    private static async Task<OfdDocumentPackage> Read(byte[] font, string face, string matrix, int objects, string? overrideText = null)
    {
        const string ns = "http://www.ofdspec.org/2016";
        var ctm = matrix == "rotation" ? "0 1 -1 0 20 0" : matrix == "shear" ? "1 0 0.5 1 0 0" : "1 0 0 1 0 0";
        // Independent raw hmtx oracle values, not output from the tested provider:
        // CJK A/i/W/space/中/文 advances 2.432/1.1/3.512/.896/4/4 mm;
        // Latin A/i/W advances 2.556/1.032/3.72 mm.
        var deltas = face == "cjk" ? "2.432 1.1 3.512 0.896 4 4" : "2.556 1.032 3.72";
        var text = overrideText ?? (face == "cjk" ? "AiW\\0020中文𠀋" : "AiW\\0020");
        var body = new StringBuilder();
        for (var i = 0; i < objects; i++)
            body.Append($"<o:TextObject ID=\"{100 + i}\" Boundary=\"20 30 100 50\" CTM=\"{ctm}\" Font=\"10\" Size=\"4\"><o:FillColor Value=\"0 80 160\"/><o:TextCode X=\"2\" Y=\"8\"" +
                (overrideText is null ? $" DeltaX=\"{deltas}\"" : "") + $">{text}</o:TextCode></o:TextObject>");
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Encoding.UTF8.GetBytes($"<o:OFD xmlns:o=\"{ns}\" Version=\"1.0\" DocType=\"OFD-H\"><o:DocBody><o:DocInfo/><o:DocRoot>Doc_0/Document.xml</o:DocRoot></o:DocBody></o:OFD>"),
            ["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes($"<o:Document xmlns:o=\"{ns}\"><o:CommonData><o:MaxUnitID>1000</o:MaxUnitID><o:PageArea><o:PhysicalBox>5 7 200 200</o:PhysicalBox></o:PageArea><o:DocumentRes>DocumentRes.xml</o:DocumentRes></o:CommonData><o:Pages><o:Page ID=\"1\" BaseLoc=\"Pages/Page_0/Content.xml\"/></o:Pages></o:Document>"),
            ["Doc_0/DocumentRes.xml"] = Encoding.UTF8.GetBytes($"<o:Res xmlns:o=\"{ns}\"><o:Fonts><o:Font ID=\"10\" FontName=\"OriginalOFL\"><o:FontFile>Res/font.ttf</o:FontFile></o:Font></o:Fonts></o:Res>"),
            ["Doc_0/Pages/Page_0/Content.xml"] = Encoding.UTF8.GetBytes($"<o:Page xmlns:o=\"{ns}\"><o:Area><o:PhysicalBox>5 7 200 200</o:PhysicalBox></o:Area><o:Content><o:Layer ID=\"1\">{body}</o:Layer></o:Content></o:Page>"),
            ["Doc_0/Res/font.ttf"] = font
        };
        using var input = new MemoryStream();
        using (var zip = new ZipArchive(input, ZipArchiveMode.Create, true))
            foreach (var pair in entries) { using var entry = zip.CreateEntry(pair.Key, CompressionLevel.NoCompression).Open(); entry.Write(pair.Value); }
        input.Position = 0; return await new OfdReader().ReadAsync(input);
    }
}
