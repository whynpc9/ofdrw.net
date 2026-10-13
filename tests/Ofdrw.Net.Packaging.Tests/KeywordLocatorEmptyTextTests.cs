using System.IO.Compression;
using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed partial class KeywordLocatorEmptyTextTests
{
    private const string Ns = "http://www.ofdspec.org/2016";
    private const string Empty = "<o:TextCode X=\"2\" Y=\"8\"/>";
    private const string PagePath = "Doc_0/Pages/Page_0/Content.xml";

    [Theory]
    [InlineData(Empty)]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\"></o:TextCode>")]
    public async Task CanonicalEmptySurvivesDefaultWriterAndCompleteNoHitWithoutProvider(string code)
    {
        var source = await Read(Object(code));
        var beforeXml = Text(source).SourceXml;
        var provider = NeverCalledProvider(); AssertEmpty(new OfdKeywordLocator(provider).Find(source, "A"));
        Assert.Equal(beforeXml, Text(source).SourceXml); Assert.Empty(Text(source).Text); Assert.Single(Text(source).Runs);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, output);
        output.Position = 0; var reread = await new OfdReader().ReadAsync(output);
        AssertEmpty(new OfdKeywordLocator(provider).Find(reread, "A"));
        Assert.Equal("30", Text(reread).ObjectId); Assert.Equal("10", Text(reread).FontResourceId);
        Assert.Single(Text(reread).Runs); Assert.Empty(Text(reread).Text);
        Assert.Equal(0, provider.Opens); Assert.Equal(0, provider.Disposals);
        Assert.Contains("\n  <ofd:Area>", Encoding.UTF8.GetString(reread.PreservedEntries[PagePath]));
        // The three synthetic resource bytes were hashed, never parsed as a font.
    }

    [Theory]
    [InlineData("1 0 0 1 0 0")]
    [InlineData("0 1 -1 0 20 0")]
    [InlineData("1 0 0.5 1 0 0")]
    public async Task RetainedEmptyOwnerPreventsMatchingAcrossOtherwiseContinuousObjects(string matrix)
    {
        var a = Object("<o:TextCode X=\"2\" Y=\"8\">A</o:TextCode>", "30", matrix: matrix);
        var empty = Object(Empty, "31", matrix: matrix);
        var b = Object("<o:TextCode X=\"4\" Y=\"8\">B</o:TextCode>", "32", matrix: matrix);
        var joined = await Read(a + b); var source = await Read(a + empty + b);
        var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider());
        Assert.Single(locator.Find(joined, "AB").Hits);
        var result = locator.Find(source, "AB"); Assert.True(result.IsComplete); Assert.Empty(result.Hits);
        Assert.Single(locator.Find(source, "A").Hits); Assert.Single(locator.Find(source, "B").Hits);
        Assert.Equal(2, result.Statistics.ReservedScalarGlyphs); Assert.Equal(2, result.Statistics.ReturnedProviderClusters);
        Assert.Equal(3, source.Pages[0].Elements.Count); Assert.Equal("31", source.Pages[0].Elements[1].ObjectId);
    }

    [Theory]
    [InlineData("0", "0")]
    [InlineData("1000000", "-1000000")]
    public async Task EmptyAnchorMayBeOutsideClipWithoutInventingInk(string x, string y)
    {
        var source = await Read(Object($"<o:TextCode X=\"{x}\" Y=\"{y}\"/>"));
        AssertEmpty(new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\" DeltaX=\"0\"/>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\" DeltaY=\"\"/>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\"><![CDATA[]]></o:TextCode>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\"><!--comment--></o:TextCode>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\"><?vendor x?></o:TextCode>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\"> </o:TextCode>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\">\t\n</o:TextCode>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\"><v:Node xmlns:v=\"urn:v\"/></o:TextCode>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\" xml:space=\"preserve\"/>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:nil=\"true\"/>")]
    public async Task NoncanonicalEmptyFormsRemainIncomplete(string codes)
    {
        var source = await Read(Object(codes));
        var provider = NeverCalledProvider(); var locator = new OfdKeywordLocator(provider);
        Assert.Throws<NotSupportedException>(() => locator.Find(source, "A"));
        var result = locator.Find(source, "A", Incomplete()); Assert.False(result.IsComplete); Assert.Empty(result.Hits);
        Assert.Equal(0, provider.Opens);
    }

    [Theory]
    [InlineData("<o:TextCode X=\"2\"/>")]
    [InlineData("<o:TextCode Y=\"8\"/>")]
    [InlineData("<o:TextCode X=\"NaN\" Y=\"8\"/>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"Infinity\"/>")]
    public async Task EmptyDoesNotBypassRequiredFiniteCoordinates(string code)
    {
        var source = await Read(Object(code));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete()));
    }

    [Fact]
    public async Task FiniteInputsWhoseComposedEmptyAnchorOverflowsRejectBeforeProvider()
    {
        var source = await Read(Object("<o:TextCode X=\"1e308\" Y=\"8\"/>", boundary: "1e308 30 100 50"));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A"));
    }

    [Theory]
    [InlineData("missing-id")]
    [InlineData("missing-font")]
    [InlineData("missing-font-file")]
    [InlineData("root-comment")]
    [InlineData("root-cdata")]
    [InlineData("root-mixed")]
    [InlineData("fill-payload")]
    [InlineData("fill-after-code")]
    [InlineData("alpha")]
    [InlineData("clip")]
    [InlineData("degenerate")]
    [InlineData("foreign-textcode")]
    public async Task EmptyDoesNotBypassSourceOwnerStyleOrResourceAuthority(string kind)
    {
        var code = kind == "foreign-textcode" ? "<v:TextCode xmlns:v=\"urn:foreign\" X=\"2\" Y=\"8\"/>" : Empty;
        var attrs = kind == "alpha" ? "Alpha=\"128\"" : "";
        var extra = kind switch
        {
            "root-comment" => "<!--comment-->", "root-cdata" => "<![CDATA[]]>", "root-mixed" => "mixed",
            "clip" => "<o:Clips/>", _ => ""
        };
        var obj = Object(code + extra, kind == "missing-id" ? null : "30", attrs,
            matrix: kind == "degenerate" ? "0 0 0 0 0 0" : "1 0 0 1 0 0",
            fill: kind == "fill-payload" ? "<o:FillColor Value=\"0 80 160\">bad</o:FillColor>" : null);
        if (kind == "fill-after-code") obj = obj.Replace("<o:FillColor Value=\"0 80 160\"/>", "").Replace("</o:TextObject>", "<o:FillColor Value=\"0 80 160\"/></o:TextObject>");
        if (kind == "missing-font") obj = obj.Replace("Font=\"10\"", "Font=\"11\"");
        var source = await Read(obj, missingFont: kind == "missing-font-file");
        var result = new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete());
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
    }

    [Theory]
    [InlineData("text")]
    [InlineData("run")]
    [InlineData("run-count")]
    [InlineData("anchor")]
    [InlineData("delta")]
    [InlineData("size")]
    public async Task EmptyRawTypedProjectionMustAgreeWithAuthoritativeSource(string kind)
    {
        var source = await Read(Object(Empty)); var text = Text(source);
        if (kind == "text") text.Text = "hidden";
        if (kind == "run") text.Runs[0].Text = "hidden";
        if (kind == "run-count") text.Runs.Add(new());
        if (kind == "anchor") text.Runs[0].XMillimeters = 3;
        if (kind == "delta") text.Runs[0].DeltaX = "";
        if (kind == "size") text.FontSizeMillimeters = double.NaN;
        Assert.False(new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete()).IsComplete);
    }

    [Fact]
    public async Task EmptyStillChecksOriginalFontBytesAndSourceFreeEmptyRemainsUnsupported()
    {
        var source = await Read(Object(Empty)); source.Fonts[0].Data = [9, 9, 9];
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A"));
        var model = KeywordLocatorTests.Model("");
        Assert.False(new OfdKeywordLocator(NeverCalledProvider()).Find(model, "A", Incomplete()).IsComplete);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("measure")]
    [InlineData("statistics")]
    [InlineData("dispose")]
    public async Task NonemptyProviderCallbacksCannotMutateAnEmptySourceOwner(string phase)
    {
        var source = await Read(Object("<o:TextCode X=\"2\" Y=\"8\">A</o:TextCode>", "30") + Object(Empty, "31"));
        var provider = new KeywordLocatorTests.Provider { Callback = step =>
        {
            if (step == phase) ((OfdTextElement)source.Pages[0].Elements[1]).Runs[0].XMillimeters = 3;
        } };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "A"));
        Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public async Task ZeroGlyphsStillConsumeAdmissionAndValidationBudgets()
    {
        var source = await Read(Object(Empty));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(NeverCalledProvider(), new(maxSingleXmlBytes: 10)).Find(source, "A"));
        var result = new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A");
        AssertEmpty(new OfdKeywordLocator(NeverCalledProvider(), new(maxValidationCharacters: result.Statistics.ValidationCharacters)).Find(source, "A"));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(NeverCalledProvider(),
            new(maxValidationCharacters: result.Statistics.ValidationCharacters - 1)).Find(source, "A"));
        var many = await Read(string.Concat(Enumerable.Range(0, 100).Select(i => Object(Empty, (30 + i).ToString()))));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(NeverCalledProvider(), new(maxNodes: 20)).Find(many, "A"));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", cancellationToken: cancellation.Token));
    }

    [Fact]
    public async Task EncodedSpaceRemainsARealGlyphRatherThanCanonicalEmpty()
    {
        var source = await Read(Object("<o:TextCode X=\"2\" Y=\"8\">\\0020</o:TextCode>"));
        var provider = new KeywordLocatorTests.Provider(); var result = new OfdKeywordLocator(provider).Find(source, " ");
        Assert.True(result.IsComplete); var hit = Assert.Single(result.Hits);
        Assert.Single(hit.Fragments); Assert.Null(hit.Fragments[0].InkQuadrilateralMillimeters);
        Assert.Equal(1, result.Statistics.ReservedScalarGlyphs); Assert.Equal(1, result.Statistics.Provider.GlyphRequests);
        Assert.Equal(1, provider.Opens); Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public async Task EmptyStillEnforcesFontByteLimitsBeforeAnyProvider()
    {
        var source = await Read(Object(Empty));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A",
            new(geometryLimits: new(maxFontBytes: 2))));
    }

    [Theory]
    [InlineData("size")]
    [InlineData("weight")]
    [InlineData("boundary")]
    public async Task EmptyStillRejectsMalformedCommonGeometryAndStyle(string kind)
    {
        var obj = Object(Empty);
        if (kind == "size") obj = obj.Replace("Size=\"4\"", "Size=\"0\"");
        if (kind == "weight") obj = obj.Replace("Size=\"4\"", "Size=\"4\" Weight=\"130\"");
        if (kind == "boundary") obj = obj.Replace("20 30 100 50", "20 30 0 50");
        var source = await Read(obj);
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HandAuthoredSourceXmlIsNotAnOriginalPackageEnvelope(bool fakeSourcePath)
    {
        var source = KeywordLocatorTests.Model(""); var text = Text(source); text.ObjectId = "30";
        text.SourceXml = Object(Empty).Replace("<o:TextObject", $"<o:TextObject xmlns:o=\"{Ns}\"");
        if (fakeSourcePath) source.Pages[0].SourceEntryPath = PagePath;
        var result = new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete());
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
    }

    private static void AssertEmpty(OfdKeywordSearchResult result)
    {
        Assert.True(result.IsComplete); Assert.Empty(result.Hits); Assert.Empty(result.Diagnostics);
        Assert.Equal(0, result.Statistics.ReservedScalarGlyphs); Assert.Equal(0, result.Statistics.ReturnedProviderClusters);
        Assert.Equal(0, result.Statistics.ReturnedFragments); Assert.Equal(0, result.Statistics.Provider.FaceParses);
        Assert.Equal(0, result.Statistics.Provider.GlyphRequests); Assert.Equal(1, result.Statistics.FontHashOperations);
        Assert.Equal(3, result.Statistics.CopiedFontBytes); Assert.Equal(3, result.Statistics.ComparedFontBytes);
    }
    private static KeywordLocatorTests.Provider NeverCalledProvider() => new() { Callback = _ => throw new InvalidOperationException("Empty text must not invoke provider callbacks.") };
    private static OfdKeywordSearchOptions Incomplete() => new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete);
    private static OfdTextElement Text(OfdDocumentPackage source) => (OfdTextElement)source.Pages[0].Elements[0];
    private static string Object(string code, string? id = "30", string attrs = "", string matrix = "1 0 0 1 0 0", string boundary = "20 30 100 50", string? fill = null) =>
        $"<o:TextObject{(id is null ? "" : $" ID=\"{id}\"")} Boundary=\"{boundary}\" Font=\"10\" Size=\"4\" CTM=\"{matrix}\" {attrs}>{fill ?? "<o:FillColor Value=\"0 80 160\"/>"}{code}</o:TextObject>";
    private static async Task<OfdDocumentPackage> Read(string objects, bool missingFont = false)
    {
        var entries = new Dictionary<string, byte[]>();
        void Add(string path, string xml) => entries.Add(path, Encoding.UTF8.GetBytes(xml));
        Add("OFD.xml", $"<o:OFD xmlns:o=\"{Ns}\" Version=\"1.0\" DocType=\"OFD-H\"><o:DocBody><o:DocInfo/><o:DocRoot>Doc_0/Document.xml</o:DocRoot></o:DocBody></o:OFD>");
        Add("Doc_0/Document.xml", $"<o:Document xmlns:o=\"{Ns}\"><o:CommonData><o:MaxUnitID>1000</o:MaxUnitID><o:PageArea><o:PhysicalBox>5 7 200 200</o:PhysicalBox></o:PageArea><o:PublicRes>PublicRes.xml</o:PublicRes></o:CommonData><o:Pages><o:Page ID=\"1\" BaseLoc=\"Pages/Page_0/Content.xml\"/></o:Pages></o:Document>");
        Add("Doc_0/PublicRes.xml", $"<o:Res xmlns:o=\"{Ns}\"><o:Fonts><o:Font ID=\"10\" FontName=\"SyntheticIdentityOnly\"><o:FontFile>Res/font.ttf</o:FontFile></o:Font></o:Fonts></o:Res>");
        Add(PagePath, $"<o:Page xmlns:o=\"{Ns}\"><o:Area><o:PhysicalBox>5 7 200 200</o:PhysicalBox></o:Area><o:Content><o:Layer ID=\"20\">{objects}</o:Layer></o:Content></o:Page>");
        if (!missingFont) entries.Add("Doc_0/Res/font.ttf", [1, 2, 3]);
        using var input = new MemoryStream();
        using (var zip = new ZipArchive(input, ZipArchiveMode.Create, true))
            foreach (var pair in entries) { using var stream = zip.CreateEntry(pair.Key).Open(); stream.Write(pair.Value); }
        input.Position = 0; return await new OfdReader().ReadAsync(input);
    }
}
