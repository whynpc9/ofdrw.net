using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class KeywordLocatorRoundtripTests
{
    private const string Ns = "http://www.ofdspec.org/2016";
    private const string PagePath = "Doc_0/Pages/Page_0/Content.xml";

    [Theory]
    [InlineData("1 0 0 1 0 0", false)]
    [InlineData("1 0 0 1 0 0", true)]
    [InlineData("0 1 -1 0 40 0", true)]
    [InlineData("1 0 0.5 1 0 0", true)]
    public async Task DefaultPublicWriterIndentationRoundtripsSelectedObjectsAndTextCodes(string matrix, bool multiple)
    {
        var source = await Read(matrix, multiple);
        var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider());
        var before = locator.Find(source, multiple ? new[] { 1 } : new[] { 0 }, "AB");
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(source, output); // Unmodified model, default public Writer.
        output.Position = 0; var afterSource = await new OfdReader().ReadAsync(output);
        var after = locator.Find(afterSource, multiple ? new[] { 1 } : new[] { 0 }, "AB");
        Assert.True(before.IsComplete); Assert.True(after.IsComplete);
        Assert.Equal(multiple ? 3 : 1, after.Hits.Count);
        Assert.Equal(before.Hits.SelectMany(h => h.Fragments).Select(f => (f.BaselineMillimeters.X, f.BaselineMillimeters.Y)),
            after.Hits.SelectMany(h => h.Fragments).Select(f => (f.BaselineMillimeters.X, f.BaselineMillimeters.Y)));
        Assert.All(after.Hits, h => { Assert.True(h.CanReplace); h.Source.ValidateUnchanged(afterSource); });
        Assert.Contains("\n  <ofd:Area>", Encoding.UTF8.GetString(afterSource.PreservedEntries[PagePath]));
    }

    [Theory]
    [InlineData("text")]
    [InlineData("attribute")]
    [InlineData("namespace")]
    [InlineData("code-order")]
    [InlineData("object-order")]
    [InlineData("text-whitespace")]
    [InlineData("cdata-kind")]
    [InlineData("container-cdata")]
    [InlineData("container-nbsp")]
    [InlineData("comment")]
    [InlineData("escape-lexeme")]
    [InlineData("foreign-attribute")]
    [InlineData("xml-space")]
    [InlineData("unknown-child")]
    [InlineData("fill-whitespace")]
    [InlineData("fill-cdata")]
    public async Task CanonicalComparisonDoesNotHideSemanticSourceChanges(string change)
    {
        var source = await Rewrite(await Read("1 0 0 1 0 0", true));
        var root = XElement.Parse(Encoding.UTF8.GetString(source.PreservedEntries[PagePath]), LoadOptions.PreserveWhitespace);
        XNamespace ns = Ns; var objects = root.Descendants(ns + "TextObject").ToArray();
        var text = objects[0]; var code = text.Element(ns + "TextCode")!;
        switch (change)
        {
            case "text": code.Value = "Z"; break;
            case "attribute": code.SetAttributeValue("X", "3"); break;
            case "namespace": code.Name = XNamespace.Get("urn:changed") + "TextCode"; break;
            case "code-order": var last = text.Elements(ns + "TextCode").Last(); last.Remove(); code.AddBeforeSelf(last); break;
            case "object-order": objects[1].Remove(); objects[0].AddBeforeSelf(objects[1]); break;
            case "text-whitespace": code.Value = " A "; break;
            case "cdata-kind": code.ReplaceNodes(new XCData("A")); break;
            case "container-cdata": text.AddFirst(new XCData(" \n")); break;
            case "container-nbsp": text.AddFirst(new XText("\u00a0")); break;
            case "comment": text.AddFirst(new XComment("different source")); break;
            case "escape-lexeme": code.Value = "A\\0020"; break;
            case "foreign-attribute": code.SetAttributeValue(XNamespace.Get("urn:foreign") + "X", "2"); break;
            case "xml-space": text.SetAttributeValue(XNamespace.Xml + "space", "preserve"); break;
            case "unknown-child": text.Add(new XElement(XNamespace.Get("urn:unknown") + "Extension", " ")); break;
            case "fill-whitespace": text.Element(ns + "FillColor")!.Value = " "; break;
            case "fill-cdata": text.Element(ns + "FillColor")!.Add(new XCData(" ")); break;
        }
        source.PreservedEntries[PagePath] = Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting));
        var provider = new KeywordLocatorTests.Provider(); var locator = new OfdKeywordLocator(provider);
        Assert.Throws<NotSupportedException>(() => locator.Find(source, new[] { 0 }, "A"));
        var incomplete = locator.Find(source, new[] { 0 }, "A", new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete));
        Assert.False(incomplete.IsComplete); Assert.Empty(incomplete.Hits); Assert.Equal(0, provider.Opens);
    }

    [Theory]
    [InlineData(" ")]
    [InlineData("\t\n")]
    [InlineData("A B")]
    [InlineData("<![CDATA[ ]]>")]
    [InlineData("<![CDATA[A B]]>")]
    [InlineData("<v:Text xmlns:v=\"urn:unknown\"> </v:Text>")]
    public async Task TextCodeSemanticWhitespaceAndUnknownMixedContentRemainUnsupported(string content)
    {
        var source = await Read("1 0 0 1 0 0", false, content);
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "A",
            new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete));
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("measure")]
    [InlineData("statistics")]
    [InlineData("dispose")]
    public async Task IndentationEquivalenceDoesNotRelaxFinalCallbackSnapshot(string phase)
    {
        var source = await Rewrite(await Read("1 0 0 1 0 0", false));
        var provider = new KeywordLocatorTests.Provider { Callback = step =>
        {
            if (step != phase) return;
            // A semantically harmless indentation edit is still a changed source during an operation.
            var original = Encoding.UTF8.GetString(source.PreservedEntries[PagePath]);
            source.PreservedEntries[PagePath] = Encoding.UTF8.GetBytes(original.Replace("\n  <ofd:Area>", "\n    <ofd:Area>"));
        } };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "AB"));
        Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public async Task ExpandedNamesPermitAnEquivalentPrefixWithoutChangingTextSpans()
    {
        var source = await Read("1 0 0 1 0 0", false);
        var text = (OfdTextElement)source.Pages[0].Elements[0];
        text.SourceXml = text.SourceXml!.Replace("o:", "alias:").Replace("xmlns:o", "xmlns:alias");
        var hit = Assert.Single(new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "AB").Hits);
        Assert.Equal(0, hit.Fragments[0].SourceUtf16Start); Assert.Equal(1, hit.Fragments[0].SourceUtf16Length);
        hit.Source.ValidateUnchanged(source);
    }

    [Fact]
    public async Task ComparisonWorkIsCumulativeAndLargeIndentRejectsBeforeProvider()
    {
        var source = await Rewrite(await Read("1 0 0 1 0 0", true));
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "AB");
        var exact = result.Statistics.ValidationCharacters;
        Assert.True(new OfdKeywordLocator(new KeywordLocatorTests.Provider(), new(maxValidationCharacters: exact)).Find(source, "AB").IsComplete);
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new KeywordLocatorTests.Provider(),
            new(maxValidationCharacters: exact - 1)).Find(source, "AB"));
        var root = XElement.Parse(Encoding.UTF8.GetString(source.PreservedEntries[PagePath]), LoadOptions.PreserveWhitespace);
        root.Descendants(XNamespace.Get(Ns) + "TextObject").First().AddFirst(new XText(new string(' ', 65_536)));
        source.PreservedEntries[PagePath] = Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting));
        var provider = new KeywordLocatorTests.Provider();
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider, new(maxValidationCharacters: exact)).Find(source, "AB"));
        Assert.Equal(0, provider.Opens);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => new OfdKeywordLocator(provider).Find(source, "AB", cancellationToken: cancellation.Token));
    }

    private static async Task<OfdDocumentPackage> Rewrite(OfdDocumentPackage source)
    {
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, output);
        output.Position = 0; return await new OfdReader().ReadAsync(output);
    }

    private static async Task<OfdDocumentPackage> Read(string matrix, bool multiple, string? codeOverride = null)
    {
        var entries = new Dictionary<string, byte[]>();
        void Add(string path, string xml) => entries.Add(path, Encoding.UTF8.GetBytes(xml));
        Add("OFD.xml", $"<o:OFD xmlns:o=\"{Ns}\" Version=\"1.0\" DocType=\"OFD-H\"><o:DocBody><o:DocInfo/><o:DocRoot>Doc_0/Document.xml</o:DocRoot></o:DocBody></o:OFD>");
        Add("Doc_0/Document.xml", $"<o:Document xmlns:o=\"{Ns}\"><o:CommonData><o:MaxUnitID>1000</o:MaxUnitID><o:PageArea><o:PhysicalBox>5 7 200 200</o:PhysicalBox></o:PageArea><o:PublicRes>PublicRes.xml</o:PublicRes></o:CommonData><o:Pages><o:Page ID=\"1\" BaseLoc=\"Pages/Page_0/Content.xml\"/>{(multiple ? "<o:Page ID=\"2\" BaseLoc=\"Pages/Page_1/Content.xml\"/>" : "")}</o:Pages></o:Document>");
        Add("Doc_0/PublicRes.xml", $"<o:Res xmlns:o=\"{Ns}\"><o:Fonts><o:Font ID=\"10\" FontName=\"Fake\"><o:FontFile>Res/font.ttf</o:FontFile></o:Font></o:Fonts></o:Res>");
        for (var page = 0; page < (multiple ? 2 : 1); page++)
        {
            var body = new StringBuilder();
            for (var i = 0; i < (multiple ? 3 : 1); i++)
                body.Append($"<o:TextObject ID=\"{30 + page * 10 + i}\" Boundary=\"20 30 100 50\" CTM=\"{matrix}\" Font=\"10\" Size=\"4\"><o:FillColor Value=\"0 80 160\"/><o:TextCode X=\"2\" Y=\"{8 + i * 10}\">{codeOverride ?? "A"}</o:TextCode><o:TextCode X=\"4\" Y=\"{8 + i * 10}\">B</o:TextCode></o:TextObject>");
            Add($"Doc_0/Pages/Page_{page}/Content.xml", $"<o:Page xmlns:o=\"{Ns}\"><o:Area><o:PhysicalBox>5 7 200 200</o:PhysicalBox></o:Area><o:Content><o:Layer ID=\"{20 + page}\">{body}</o:Layer></o:Content></o:Page>");
        }
        entries.Add("Doc_0/Res/font.ttf", [1, 2, 3]);
        using var input = new MemoryStream();
        using (var zip = new ZipArchive(input, ZipArchiveMode.Create, true))
            foreach (var pair in entries) { using var stream = zip.CreateEntry(pair.Key).Open(); stream.Write(pair.Value); }
        input.Position = 0; return await new OfdReader().ReadAsync(input);
    }
}
