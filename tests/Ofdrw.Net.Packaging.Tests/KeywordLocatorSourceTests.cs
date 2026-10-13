using System.IO.Compression;
using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class KeywordLocatorSourceTests
{
    [Fact]
    public async Task ActualReaderSourceDecodesEscapesAndRetainsNamespaceAndRawValueRanges()
    {
        var source = await Read(Object("<o:TextCode X=\"2\" Y=\"8\" DeltaX=\"2 1\">A\\0020\\d840\\dc0b</o:TextCode>"));
        var original = Text(source).SourceXml;
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, " 𠀋");
        Assert.True(result.IsComplete); var hit = Assert.Single(result.Hits);
        Assert.Equal(5, hit.Fragments[0].SourceUtf16Length); Assert.Equal(1, hit.Fragments[0].SourceUtf16Start);
        Assert.Equal(10, hit.Fragments[1].SourceUtf16Length); Assert.Equal(2, hit.Fragments[1].LogicalUtf16Length);
        Assert.Contains("{http://www.ofdspec.org/2016}TextCode[0]", hit.Fragments[0].TextCodePath);
        Assert.True(hit.Fragments[0].IsAuthoritativeXml); Assert.Equal("Doc_0/Pages/Page_0/Content.xml", hit.Fragments[0].SourceEntryPath);
        Assert.Equal(original, Text(source).SourceXml); Assert.Equal("A\\0020\\d840\\dc0b", Text(source).Text);
        hit.Source.ValidateUnchanged(source);
    }

    [Fact]
    public async Task Table46InheritsPreviousTextCodeStartAndEachAxisIndependently()
    {
        var codes = "<o:TextCode X=\"2\" Y=\"8\" DeltaX=\"4\">AA</o:TextCode>" +
            "<o:TextCode Y=\"9\">B</o:TextCode><o:TextCode X=\"10\">C</o:TextCode>";
        var source = await Read(Object(codes)); var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider());
        var b = Assert.Single(locator.Find(source, "B").Hits).Fragments[0].BaselineMillimeters;
        var c = Assert.Single(locator.Find(source, "C").Hits).Fragments[0].BaselineMillimeters;
        Assert.Equal(17, b.X); Assert.Equal(32, b.Y); Assert.Equal(25, c.X); Assert.Equal(32, c.Y);
        Assert.Empty(locator.Find(source, "AAB").Hits);
        // Reader's old redundant typed defaults remain unchanged and are not used.
        Assert.Equal(0, Text(source).Runs[1].XMillimeters); Assert.Equal(4, Text(source).Runs[2].YMillimeters);
    }

    [Theory]
    [InlineData("X=\"2\"")]
    [InlineData("Y=\"8\"")]
    [InlineData("")]
    public async Task FirstTextCodeRequiresBothCoordinates(string attributes)
    {
        var source = await Read(Object("<o:TextCode " + attributes + ">A</o:TextCode>"));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "A"));
    }

    [Fact]
    public async Task NoDeltaDoesNotUseNonzeroProviderAdvance()
    {
        var source = await Read(Object("<o:TextCode X=\"2\" Y=\"8\">AA</o:TextCode>"));
        var hit = Assert.Single(new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "AA").Hits);
        Assert.Equal(hit.Fragments[0].BaselineMillimeters.X, hit.Fragments[1].BaselineMillimeters.X);
        Assert.Equal(2, hit.Fragments[0].AdvanceXMillimeters);
    }

    [Theory]
    [InlineData("HScale=\"0.5\"")]
    [InlineData("ReadDirection=\"90\"")]
    [InlineData("CharDirection=\"90\"")]
    [InlineData("DrawParam=\"99\"")]
    [InlineData("Alpha=\"128\"")]
    [InlineData("Stroke=\"true\"")]
    [InlineData("Opacity=\"1\"")]
    [InlineData("x:extension=\"1\" xmlns:x=\"urn:foreign\"")]
    public async Task UnresolvedDrawingStateNeverProducesPurportedExactGeometry(string attributes)
    {
        var source = await Read(Object(Code("A"), attributes));
        var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider());
        Assert.Throws<NotSupportedException>(() => locator.Find(source, "A"));
        var result = locator.Find(source, "A", Incomplete()); Assert.False(result.IsComplete); Assert.Empty(result.Hits);
    }

    [Theory]
    [InlineData("<o:Clips><o:Clip><o:Area><o:Path Boundary=\"0 0 10 10\"><o:AbbreviatedData>M 0 0 L 1 1</o:AbbreviatedData></o:Path></o:Area></o:Clip></o:Clips>")]
    [InlineData("<o:CGTransform CodePosition=\"0\" CodeCount=\"1\" GlyphCount=\"1\"><o:Glyphs>1</o:Glyphs></o:CGTransform>")]
    [InlineData("<o:Vendor><o:TextCode X=\"2\" Y=\"8\">secret</o:TextCode></o:Vendor>")]
    public async Task ClipCgTransformAndUnknownChildrenAreUnsupported(string child)
    {
        var source = await Read(Object(Code("A") + child));
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "A", Incomplete());
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
    }

    [Theory]
    [InlineData("DrawParam=\"99\"")]
    [InlineData("CTM=\"1 0 0 1 2 3\"")]
    [InlineData("Visible=\"false\"")]
    [InlineData("Alpha=\"128\"")]
    public async Task DroppedLayerAttributesAreAdmittedFromOriginalEnvelopeOrExplicitlyRefused(string attributes)
    {
        var source = await Read(Object(Code("A")), layerAttributes: attributes);
        Assert.Equal("A", Text(source).Text); // Reader projected it, but inherited semantics are not certified.
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "A", Incomplete());
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
    }

    [Theory]
    [InlineData("page-extension")]
    [InlineData("content-extension")]
    [InlineData("foreign-content")]
    [InlineData("page-rotation")]
    [InlineData("missing-entry")]
    [InlineData("page-resource")]
    public async Task A0OmittedEnvelopeAndResourceSourcesCannotYieldCompleteNoHit(string kind)
    {
        var source = await Read(Object(Code("A")),
            pageExtra: kind == "page-extension" ? "<o:Vendor><o:TextObject>hidden</o:TextObject></o:Vendor>" :
                kind == "page-resource" ? "<o:PageRes>PageRes.xml</o:PageRes>" : "",
            contentExtra: kind == "content-extension" ? "<o:Vendor><o:TextObject>hidden</o:TextObject></o:Vendor>" : "",
            pageAttributes: kind == "page-rotation" ? "Rotate=\"90\"" : "",
            duplicateFont: false);
        if (kind == "missing-entry") source.PreservedEntries.Remove("Doc_0/Pages/Page_0/Content.xml");
        if (kind == "foreign-content")
        {
            var bytes = source.PreservedEntries["Doc_0/Pages/Page_0/Content.xml"];
            source.PreservedEntries["Doc_0/Pages/Page_0/Content.xml"] = Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace("<o:Content>", "<Content xmlns=\"urn:foreign\">").Replace("</o:Content>", "</Content>"));
        }
        var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider());
        Assert.Throws<NotSupportedException>(() => locator.Find(source, "absent"));
        Assert.False(locator.Find(source, "absent", Incomplete()).IsComplete);
    }

    [Fact]
    public async Task DuplicateKnownFontDefinitionsAreRejectedBeforeKeywordQuery()
    {
        var duplicate = await Assert.ThrowsAsync<InvalidDataException>(() => Read(Object(Code("A")), duplicateFont: true));
        Assert.Contains("Duplicate active document-unit ID", duplicate.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReaderRecognizedFontsHiddenOutsideFontsCannotBypassSourceIdentity(bool foreign)
    {
        var hidden = foreign ? "<x:Font xmlns:x=\"urn:foreign\" ID=\"10\" FontName=\"Fake\"><x:FontFile>Res/font.ttf</x:FontFile></x:Font>" :
            "<o:Font ID=\"10\" FontName=\"Fake\"><o:FontFile>Res/font.ttf</o:FontFile></o:Font>";
        var source = await Read(Object(Code("A")), resourceExtra: "<o:DrawParams>" + hidden + "</o:DrawParams>");
        Assert.Single(source.Fonts); // Reader discarded the duplicate instead of certifying uniqueness.
        var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider());
        Assert.Throws<NotSupportedException>(() => locator.Find(source, "A"));
        var result = locator.Find(source, "A", Incomplete());
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
        Assert.Equal("UnexpectedFontDeclarationLocation", Assert.Single(result.Diagnostics).Code);
    }

    [Fact]
    public async Task UnusedKnownNonfontResourceContainerDoesNotPolluteFontIdentity()
    {
        var source = await Read(Object(Code("A")), resourceExtra: "<o:DrawParams><o:DrawParam ID=\"99\"><o:FillColor Value=\"0 0 0\"/></o:DrawParam></o:DrawParams>");
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "A");
        Assert.True(result.IsComplete); Assert.Single(result.Hits);
    }

    [Fact]
    public async Task RawContentAndDistinctLayerScopesNeverCreateFalseAdjacentMatches()
    {
        var a = Object(Code("A")); var b = Object("<o:TextCode X=\"4\" Y=\"8\">B</o:TextCode>", id: "11");
        var raw = "<o:PageBlock ID=\"9\"><o:TextObject ID=\"32\" Boundary=\"0 0 1 1\" Font=\"10\" Size=\"4\">" + Code("hidden") + "</o:TextObject></o:PageBlock>";
        var source = await Read(a + raw + b);
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "AB", Incomplete());
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
        var separate = await Read(a, contentExtra: "<o:Layer ID=\"21\">" + b + "</o:Layer>");
        Assert.NotEqual(Text(separate).LayerId, ((OfdTextElement)separate.Pages[0].Elements[1]).LayerId);
        var complete = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(separate, "AB");
        Assert.True(complete.IsComplete); Assert.Empty(complete.Hits);
    }

    [Fact]
    public async Task DuplicateLayerAndCrossKindUnitIdsAreRejectedBeforeKeywordQuery()
    {
        var a = Object(Code("A")); var b = Object("<o:TextCode X=\"4\" Y=\"8\">B</o:TextCode>", id: "11");
        var duplicateLayer = await Assert.ThrowsAsync<InvalidDataException>(() => Read(a, contentExtra: "<o:Layer ID=\"20\">" + b + "</o:Layer>"));
        Assert.Contains("Duplicate active document-unit ID", duplicateLayer.Message);
        var fontCollision = await Assert.ThrowsAsync<InvalidDataException>(() => Read(Object(Code("A"), id: "10")));
        Assert.Contains("Duplicate active document-unit ID", fontCollision.Message);
        var pageCollision = await Assert.ThrowsAsync<InvalidDataException>(() => Read(a, contentExtra: "<o:Layer ID=\"1\">" + b + "</o:Layer>"));
        Assert.Contains("Duplicate active document-unit ID", pageCollision.Message);
    }

    [Fact]
    public async Task HiddenTextIsADeclaredVisibilityExclusionAndContinuityBarrier()
    {
        var source = await Read(Object(Code("A")) + Object(Code("hidden"), "Visible=\"false\"", "11") +
            Object("<o:TextCode X=\"4\" Y=\"8\">B</o:TextCode>", id: "12"));
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "AB");
        Assert.True(result.IsComplete); Assert.Empty(result.Hits);
    }

    [Theory]
    [InlineData("xml")]
    [InlineData("entry")]
    [InlineData("path")]
    [InlineData("resource-entry")]
    public async Task ImportedSourceEvidenceCannotBeMutatedByProviderCallbacks(string kind)
    {
        var source = await Read(Object(Code("A")));
        var provider = new KeywordLocatorTests.Provider { Callback = phase =>
        {
            if (phase != "measure") return;
            if (kind == "xml") Text(source).SourceXml += " ";
            if (kind == "entry") source.PreservedEntries["Doc_0/Pages/Page_0/Content.xml"][20] ^= 1;
            if (kind == "path") source.Pages[0].SourceEntryPath = "different.xml";
            if (kind == "resource-entry") source.PreservedEntries["Doc_0/Res/font.ttf"] = [9, 9, 9];
        } };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "A"));
        Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public async Task UnknownWhitespaceEncodingAndTypedOverlayConflictsAreExplicitUnsupported()
    {
        foreach (var text in new[] { "A B", "A\\bad", "A\\002Z", "¤" })
        {
            var source = await Read(Object(Code(text)));
            Assert.False(new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "A", Incomplete()).IsComplete);
        }
        var conflict = await Read(Object(Code("A"))); Text(conflict).Weight = 700;
        Assert.False(new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(conflict, "A", Incomplete()).IsComplete);
    }

    [Fact]
    public async Task AdmissionAndDiagnosticBudgetsRejectBeforeDomOrTruncation()
    {
        var source = await Read(Object(Code("A")));
        var provider = new KeywordLocatorTests.Provider();
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider, new(maxSingleXmlBytes: 10)).Find(source, "A"));
        Assert.Equal(0, provider.Opens);
        var raw = await Read("<o:PageBlock /><o:PageBlock />");
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider, new(maxDiagnostics: 1)).Find(raw, "A", Incomplete()));
        var callback = new KeywordLocatorTests.Provider { Callback = phase => { if (phase == "measure") source.Pages[0].PreservedPageElements.Add("<Vendor />"); } };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(callback).Find(source, "A"));
    }

    private static OfdKeywordSearchOptions Incomplete() => new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete);
    private static OfdTextElement Text(OfdDocumentPackage package) => (OfdTextElement)package.Pages[0].Elements[0];
    private static string Code(string value) => "<o:TextCode X=\"2\" Y=\"8\">" + value + "</o:TextCode>";
    private static string Object(string codes, string attributes = "", string id = "30") =>
        "<o:TextObject ID=\"" + id + "\" Boundary=\"20 30 100 50\" Font=\"10\" Size=\"4\" " + attributes + ">" + codes + "</o:TextObject>";
    private static async Task<OfdDocumentPackage> Read(string objects, string layerAttributes = "", string pageExtra = "",
        string contentExtra = "", string pageAttributes = "", bool duplicateFont = false, string resourceExtra = "")
    {
        const string ns = "http://www.ofdspec.org/2016";
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Encoding.UTF8.GetBytes($"<o:OFD xmlns:o=\"{ns}\" Version=\"1.0\" DocType=\"OFD-H\"><o:DocBody><o:DocInfo/><o:DocRoot>Doc_0/Document.xml</o:DocRoot></o:DocBody></o:OFD>"),
            ["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes($"<o:Document xmlns:o=\"{ns}\"><o:CommonData><o:MaxUnitID>100</o:MaxUnitID><o:PageArea><o:PhysicalBox>5 7 200 200</o:PhysicalBox></o:PageArea><o:DocumentRes>DocumentRes.xml</o:DocumentRes></o:CommonData><o:Pages><o:Page ID=\"1\" BaseLoc=\"Pages/Page_0/Content.xml\"/></o:Pages></o:Document>"),
            ["Doc_0/DocumentRes.xml"] = Encoding.UTF8.GetBytes($"<o:Res xmlns:o=\"{ns}\"><o:Fonts><o:Font ID=\"10\" FontName=\"Fake\"><o:FontFile>Res/font.ttf</o:FontFile></o:Font>" + (duplicateFont ? "<o:Font ID=\"10\" FontName=\"Another\"><o:FontFile>Res/font.ttf</o:FontFile></o:Font>" : "") + "</o:Fonts>" + resourceExtra + "</o:Res>"),
            ["Doc_0/Pages/Page_0/Content.xml"] = Encoding.UTF8.GetBytes($"<o:Page xmlns:o=\"{ns}\" {pageAttributes}><o:Area><o:PhysicalBox>5 7 200 200</o:PhysicalBox></o:Area>{pageExtra}<o:Content><o:Layer ID=\"20\" {layerAttributes}>{objects}</o:Layer>{contentExtra}</o:Content></o:Page>"),
            ["Doc_0/Res/font.ttf"] = [1, 2, 3]
        };
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
            foreach (var pair in entries) { using var output = zip.CreateEntry(pair.Key).Open(); output.Write(pair.Value); }
        stream.Position = 0; return await new OfdReader().ReadAsync(stream);
    }
}
