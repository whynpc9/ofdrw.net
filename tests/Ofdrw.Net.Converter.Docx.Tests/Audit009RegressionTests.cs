using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class Audit009RegressionTests
{
    public static IEnumerable<object[]> LossCases()
    {
        foreach (var native in new[] { false, true })
        foreach (var footer in new[] { false, true })
        foreach (var kind in new[] { "default", "first", "even" })
        foreach (var policy in new[] { UnsupportedDocxFeatureBehavior.BestEffort, UnsupportedDocxFeatureBehavior.Placeholder })
            yield return [native, footer, kind, policy];
    }

    [Theory]
    [MemberData(nameof(LossCases))]
    public async Task Audit009_DecorationLoss(bool native, bool footer, string kind, UnsupportedDocxFeatureBehavior policy)
    {
        using var source = CreateFixture(footer, kind, unsupported: true, inherited: true);
        using var output = new MemoryStream();
        var options = Options(native, policy);
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(source, output);
        Assert.False(result.OriginalTextPreserved);
        var diagnostic = Assert.Single(result.Diagnostics, d => d.Code == (footer ? "DOCX_FOOTER_CONTENT_UNSUPPORTED" : "DOCX_HEADER_CONTENT_UNSUPPORTED"));
        Assert.Contains(footer ? "/word/footer" : "/word/header", diagnostic.Message);
        Assert.Contains(kind, diagnostic.Message);
        Assert.Contains(footer ? "sdt" : "tbl", diagnostic.Message);
        Assert.Contains("child[2]", diagnostic.Message);
        output.Position = 0;
        var text = new OfdTextExtractor().Extract(await new OfdReader().ReadAsync(output));
        Assert.DoesNotContain("遗漏中文", text);
        Assert.Contains("Body one", text);
        Assert.Contains("Body inherited", text);
        Assert.Equal(policy == UnsupportedDocxFeatureBehavior.Placeholder, text.Contains("Unsupported DOCX feature:"));
        Save($"loss-{(native ? "native" : "default")}-{(footer ? "footer" : "header")}-{kind}-{policy}", source, output, result);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit009_SupportedDecoration(bool native)
    {
        using var source = CreateFixture(false, "default", unsupported: false, inherited: true);
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter(Options(native)).ConvertWithResultAsync(source, output);
        Assert.True(result.OriginalTextPreserved);
        Assert.DoesNotContain(result.Diagnostics, d => d.Code.Contains("UNSUPPORTED"));
        output.Position = 0;
        var pages = new OfdTextExtractor().ExtractPages(await new OfdReader().ReadAsync(output));
        Assert.Equal(6, pages.Count);
        Assert.Contains("First supported", pages[0]);
        Assert.Contains("Even supported", pages[1]);
        Assert.Contains("Default supported", pages[2]);
        Assert.Contains("First supported", pages[3]);
        for (var index = 0; index < pages.Count; index++)
            Assert.Contains($"Page{index + 1}/6", pages[index].Replace(" ", "").Replace("\n", ""));
        Save(native ? "supported-native" : "supported-default", source, output, result);
        source.Position = 0;
        using var selected = new MemoryStream();
        var selection = await new DocxToOfdConverter(Options(native)).ConvertWithResultAsync(source, selected, [1, 0, 1]);
        Assert.Equal(new[] { 1, 0, 1 }, selection.SourcePages);
        selected.Position = 0;
        var selectedPages = new OfdTextExtractor().ExtractPages(await new OfdReader().ReadAsync(selected));
        Assert.Equal(pages[1], selectedPages[0]);
        Assert.Equal(pages[0], selectedPages[1]);
        Assert.Equal(selectedPages[0], selectedPages[2]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit009_ThrowBeforeWrite(bool footer)
    {
        using var source = CreateFixture(footer, "default", unsupported: true, inherited: false);
        using var output = new MemoryStream();
        output.Write([1, 2, 3, 4, 5]);
        output.Position = 2;
        await Assert.ThrowsAsync<NotSupportedException>(() => new DocxToOfdConverter(Options(true, UnsupportedDocxFeatureBehavior.Throw)).ConvertWithResultAsync(source, output));
        Assert.Equal(new byte[] { 1, 2, 3, 4, 5 }, output.ToArray());
        Assert.Equal(2, output.Position);
    }

    [Fact]
    public async Task Audit009_SharedPartIsAuditedOnceAcrossVariantsAndSections()
    {
        using var source = CreateFixture(false, "default", unsupported: true, inherited: true);
        using (var document = WordprocessingDocument.Open(source, true))
        {
            var references = document.MainDocumentPart!.Document.Descendants<W.HeaderReference>().ToArray();
            foreach (var reference in references) reference.Id = references[0].Id!.Value;
        }
        source.Position = 0;
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter().ConvertWithResultAsync(source, output);
        Assert.False(result.OriginalTextPreserved);
        Assert.Single(result.Diagnostics, d => d.Code == "DOCX_HEADER_CONTENT_UNSUPPORTED");
        output.Position = 0;
        var pages = new OfdTextExtractor().ExtractPages(await new OfdReader().ReadAsync(output));
        Assert.Equal(6, pages.Count);
        Assert.All(pages, page => Assert.Equal(1, page.Split("Default supported").Length - 1));
    }

    [Fact]
    public async Task Audit009_EmptyUnsupportedDecorationDoesNotClaimTextLoss()
    {
        using var source = CreateFixture(false, "default", unsupported: true, inherited: false, empty: true);
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter().ConvertWithResultAsync(source, output);
        Assert.True(result.OriginalTextPreserved);
        Assert.Contains(result.Diagnostics, d => d.Code == "DOCX_HEADER_CONTENT_UNSUPPORTED");
    }

    private static DocxConversionOptions Options(bool native, UnsupportedDocxFeatureBehavior policy = UnsupportedDocxFeatureBehavior.BestEffort)
    {
        var options = new DocxConversionOptions { UnsupportedFeatureBehavior = policy };
        if (native) options.OfdMode = DocxToOfdMode.Native;
        return options;
    }

    private static W.Paragraph Paragraph(string text) => new(new W.Run(new W.Text(text)));

    private static MemoryStream CreateFixture(bool footer, string kind, bool unsupported, bool inherited, bool empty = false)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body());
            var settings = main.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new W.Settings(new W.EvenAndOddHeaders());
            var properties = new W.SectionProperties(new W.TitlePage());
            foreach (var type in new[] { W.HeaderFooterValues.Default, W.HeaderFooterValues.First, W.HeaderFooterValues.Even })
            {
                var label = type == W.HeaderFooterValues.First ? "First" : type == W.HeaderFooterValues.Even ? "Even" : "Default";
                var header = main.AddNewPart<HeaderPart>();
                header.Header = new W.Header(Paragraph(label + " supported"));
                var foot = main.AddNewPart<FooterPart>();
                foot.Footer = new W.Footer(new W.Paragraph(new W.Run(new W.Text("Page ")), new W.SimpleField(new W.Run(new W.Text("1"))) { Instruction = "PAGE" }, new W.Run(new W.Text("/")), new W.SimpleField(new W.Run(new W.Text("6"))) { Instruction = "NUMPAGES" }));
                if (unsupported && label.ToLowerInvariant() == kind)
                {
                    OpenXmlElement node = footer
                        ? new W.SdtBlock(new W.SdtContentBlock(Paragraph(empty ? "" : "遗漏中文 footer SDT")))
                        : new W.Table(new W.TableRow(new W.TableCell(Paragraph(empty ? "" : "遗漏中文 header table"))));
                    if (footer) foot.Footer.Append(node); else header.Header.Append(node);
                }
                properties.Append(new W.HeaderReference { Type = type, Id = main.GetIdOfPart(header) });
                properties.Append(new W.FooterReference { Type = type, Id = main.GetIdOfPart(foot) });
            }
            var body = main.Document.Body!;
            body.Append(Paragraph("Body one 中文比例 Alpha"), new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })), Paragraph("Body two"), new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })), Paragraph("Body three"));
            if (inherited)
            {
                body.Append(new W.Paragraph(new W.ParagraphProperties((W.SectionProperties)properties.CloneNode(true))));
                body.Append(Paragraph("Body inherited"), new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })), Paragraph("Body five"), new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })), Paragraph("Body six"));
                body.Append(new W.SectionProperties(new W.TitlePage()));
            }
            else body.Append(properties);
        }
        stream.Position = 0;
        return stream;
    }

    private static void Save(string name, MemoryStream source, MemoryStream output, DocxToOfdConversionResult result)
    {
        var directory = Environment.GetEnvironmentVariable("OFDRW_PLAN009_ARTIFACTS");
        if (string.IsNullOrEmpty(directory)) return;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, name + ".docx"), source.ToArray());
        File.WriteAllBytes(Path.Combine(directory, name + ".ofd"), output.ToArray());
        File.WriteAllText(Path.Combine(directory, name + ".txt"), $"Mode={result.Mode}; OriginalTextPreserved={result.OriginalTextPreserved}; Pages={result.SourcePageCount}\n" + string.Join("\n", result.Diagnostics.Select(d => d.Code + ": " + d.Message)));
    }
}
