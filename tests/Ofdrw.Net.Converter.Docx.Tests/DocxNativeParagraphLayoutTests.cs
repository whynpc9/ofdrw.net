using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;
using Ofdrw.Net.Text.HarfBuzz;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Tests;

// Issue #109: DOCX Native OFD opts into the final paragraph layout engine.
public sealed class DocxNativeParagraphLayoutTests
{
    private static DocxConversionOptions Options(PinnedCjkFontDirectory fonts, bool engine, bool explicitNative = true)
    {
        var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false };
        if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        options.FontDirectories.Add(fonts.Path);
        if (engine) options.NativeParagraphLayoutEngine = HarfBuzzTextPrelayout.CreateParagraphLayout();
        return options;
    }

    private static async Task<(OfdDocumentPackage Package, DocxToOfdConversionResult Result)> Convert(
        Stream source, DocxConversionOptions options, string? artifact = null)
    {
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(source, output);
        if (artifact is not null && Environment.GetEnvironmentVariable("OFDRW_PARAGRAPH_ARTIFACTS") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            source.Position = 0;
            using var inputCopy = File.Create(Path.Combine(directory, artifact + ".docx"));
            await source.CopyToAsync(inputCopy);
            await File.WriteAllBytesAsync(Path.Combine(directory, artifact + ".ofd"), output.ToArray());
        }
        output.Position = 0;
        return (await new OfdReader().ReadAsync(output), result);
    }

    private static string RawText(OfdDocumentPackage package) =>
        string.Concat(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(text => text.Text));

    private static string Text(OfdDocumentPackage package) => RawText(package).Replace(" ", "");

    [Fact]
    public async Task GeneratedLayoutSampleKeepsItsTextAndUsesClusterMappedFinalLayout()
    {
        using var fonts = new PinnedCjkFontDirectory();
        using var legacySource = File.OpenRead(Sample());
        var (legacy, defaultResult) = await Convert(legacySource, Options(fonts, engine: false, explicitNative: false), "generated-layout-default");
        using var finalSource = File.OpenRead(Sample());
        var (final, result) = await Convert(finalSource, Options(fonts, engine: true), "generated-layout-native");
        Assert.Equal(DocxToOfdMode.Native, defaultResult.Mode);
        Assert.True(defaultResult.OriginalTextPreserved);
        Assert.True(result.OriginalTextPreserved);
        Assert.Equal(Enumerable.Range(0, defaultResult.SourcePageCount), defaultResult.SourcePages);
        Assert.Equal(Enumerable.Range(0, result.SourcePageCount), result.SourcePages);
        Assert.Equal("ParagraphLayout", final.CustomTags["docx-native-text-engine"]);
        Assert.False(legacy.CustomTags.ContainsKey("docx-native-text-engine"));
        // The same logical text, in the same order; only geometry and font binding change.
        Assert.Equal(RawText(legacy), RawText(final));
        var texts = final.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().ToArray();
        Assert.All(texts, text => Assert.Equal("true", XElement.Parse(text.SourceXml!).Attributes()
            .Single(attribute => attribute.Name.LocalName == "ClusterMapV1").Value));
        Assert.All(final.Fonts, font => Assert.NotEmpty(font.Data));
        // Bold runs without a bold face are synthesized with a fill-colored stroke, not Weight.
        var bold = texts.Where(text => XElement.Parse(text.SourceXml!).Attribute("Stroke")?.Value == "true").ToArray();
        Assert.NotEmpty(bold);
        Assert.All(bold, text => Assert.Equal(OfdTextElement.DefaultWeight, text.Weight));
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_TEXT_ENGINE_FALLBACK");
        Assert.InRange(final.Pages.Count, legacy.Pages.Count - 1, legacy.Pages.Count + 1);
    }

    [Fact]
    public async Task TabsStrikeHighlightAndHangingIndentReachTheEngine()
    {
        const string body = """
            <w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2880"/><w:tab w:val="right" w:pos="8640"/></w:tabs></w:pPr>
              <w:r><w:t>Name</w:t></w:r><w:r><w:tab/><w:t>张三</w:t></w:r><w:r><w:tab/><w:t>12.50</w:t></w:r></w:p>
            <w:p><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
              <w:r><w:t>1.</w:t></w:r><w:r><w:tab/><w:t>Hanging list item 列表项</w:t></w:r></w:p>
            <w:p><w:r><w:rPr><w:strike/></w:rPr><w:t>struck 删除</w:t></w:r><w:r><w:t> and </w:t></w:r>
              <w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>marked 高亮</w:t></w:r></w:p>
            """;
        using var fonts = new PinnedCjkFontDirectory();
        using var source = DocxReleaseFixtures.CreateMinimalDocx(body);
        var (package, result) = await Convert(source, Options(fonts, engine: true));
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_INDENT_LIMITED");
        var page = package.Pages.Single();
        var texts = page.Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(3, texts.Count(text => text.Text == "\t"));
        // Word measures stops from the text column: 2880 twips = 50.8 mm, right stop 8640 twips = 152.4 mm.
        var left = texts.Single(text => text.Text == "张三");
        var margin = texts.Single(text => text.Text == "Name").XMillimeters;
        Assert.InRange(left.XMillimeters - margin, 50.7, 51.0);
        var right = texts.Single(text => text.Text == "12.50");
        Assert.InRange(right.XMillimeters + right.WidthMillimeters - margin, 152.0, 152.6);
        // The hanging first line starts 360 twips left of the 720-twip indent; its tab reaches the indent.
        var item = texts.Single(text => text.Text.StartsWith("Hanging", StringComparison.Ordinal));
        Assert.InRange(item.XMillimeters - margin, 12.6, 12.9);
        // Strike and highlight are filled rectangles; the highlight precedes the text it marks.
        var paths = page.Elements.OfType<OfdPathElement>().ToArray();
        Assert.Contains(paths, path => path.Fill && path.FillColor is { Red: 255, Green: 255, Blue: 0 });
        var highlightIndex = page.Elements.ToList().FindIndex(element => element is OfdPathElement { Fill: true } path && path.FillColor is { Red: 255, Blue: 0 });
        var markedIndex = page.Elements.ToList().FindIndex(element => element is OfdTextElement text && text.Text.StartsWith("marked", StringComparison.Ordinal));
        Assert.True(highlightIndex >= 0 && highlightIndex < markedIndex);
        // "marked 高亮" spans a Latin and a Han face but is one highlight path (Issue #115), and the
        // strikethrough over "struck 删除" is one more.
        Assert.Single(paths, path => path.Fill && path.FillColor is { Red: 255, Green: 255, Blue: 0 });
        Assert.True(paths.Count(path => path.Fill) >= 2);
    }

    [Theory]
    [InlineData("<w:ind w:left=\"10000\" w:hanging=\"720\"/>")]
    [InlineData("<w:ind w:left=\"4000\" w:right=\"4000\" w:firstLine=\"2000\"/>")]
    public async Task IndentThatDoesNotFitTheColumnIsLimitedNotFatal(string indent)
    {
        var body = "<w:p><w:pPr>" + indent + "</w:pPr><w:r><w:t>Indented 缩进 text</w:t></w:r></w:p>" +
                   "<w:p><w:r><w:t>Next 下一段</w:t></w:r></w:p>";
        using var fonts = new PinnedCjkFontDirectory();
        using var source = DocxReleaseFixtures.CreateMinimalDocx(body);
        var artifact = "indent-limited-" + (indent.Contains("hanging", StringComparison.Ordinal) ? "hanging" : "first-line");
        var (package, result) = await Convert(source, Options(fonts, engine: true), artifact);
        source.Position = 0;
        var (stock, stockResult) = await Convert(source, Options(fonts, engine: false, explicitNative: false), artifact + "-default");
        Assert.Equal("Indented 缩进 textNext 下一段", RawText(package));
        Assert.Equal(RawText(package), RawText(stock));
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_INDENT_LIMITED");
        Assert.DoesNotContain(stockResult.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_INDENT_LIMITED");
        Assert.Equal(DocxToOfdMode.Native, stockResult.Mode);
        Assert.True(result.OriginalTextPreserved);
        Assert.Equal(0, Assert.Single(result.SourcePages));
        Assert.Equal(0, Assert.Single(stockResult.SourcePages));
        var page = Assert.Single(package.Pages);
        Assert.All(page.Elements.OfType<OfdTextElement>(), text =>
        {
            Assert.InRange(text.XMillimeters, 0, page.WidthMillimeters);
            Assert.InRange(text.XMillimeters + text.WidthMillimeters, 0, page.WidthMillimeters);
        });
        Assert.Equal(Assert.Single(Assert.Single(stock.Pages).Elements.OfType<OfdTextElement>(),
                text => text.Text.StartsWith("Next", StringComparison.Ordinal)).XMillimeters,
            Assert.Single(page.Elements.OfType<OfdTextElement>(),
                text => text.Text.StartsWith("Next", StringComparison.Ordinal)).XMillimeters, precision: 3);
    }

    [Fact]
    public async Task LeadingBreakDoesNotReportAnUnusedIndent()
    {
        const string body = """
            <w:p><w:pPr><w:ind w:left="4000" w:right="4000" w:firstLine="2000"/></w:pPr>
              <w:r><w:br/><w:t>After break</w:t></w:r></w:p>
            """;
        using var fonts = new PinnedCjkFontDirectory();
        using var source = DocxReleaseFixtures.CreateMinimalDocx(body);
        var (package, result) = await Convert(source, Options(fonts, engine: true));
        Assert.Equal("Afterbreak", Text(package));
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_INDENT_LIMITED");
    }

    [Fact]
    public async Task IndentLimitIsReportedOnceAcrossParagraphsAndPaginationPasses()
    {
        const string body = """
            <w:p><w:pPr><w:ind w:left="10000" w:hanging="720"/></w:pPr><w:r><w:t>Hanging</w:t></w:r></w:p>
            <w:p><w:pPr><w:pageBreakBefore/><w:ind w:left="4000" w:right="4000" w:firstLine="2000"/></w:pPr>
              <w:r><w:t>First line</w:t></w:r></w:p>
            """;
        using var fonts = new PinnedCjkFontDirectory();
        using var source = DocxReleaseFixtures.CreateMinimalDocx(body);
        var (package, result) = await Convert(source, Options(fonts, engine: true));
        Assert.Equal(2, result.SourcePageCount);
        Assert.Equal(new[] { 0, 1 }, result.SourcePages);
        Assert.Equal("HangingFirstline", Text(package));
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_INDENT_LIMITED");
    }

    [Fact]
    public async Task InlineImageFallbackDoesNotReportAnUnusedIndent()
    {
        const string label = "Native image";
        var (docx, _) = DocxReleaseFixtures.CreateInlineJpegDocx(label);
        using var source = docx;
        using (var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(source, true))
        {
            var body = Assert.IsType<W.Body>(document.MainDocumentPart?.Document?.Body);
            var paragraph = Assert.Single(body.Elements<W.Paragraph>(), value => value.Descendants<W.Drawing>().Any());
            paragraph.ParagraphProperties = new(new W.Indentation
            {
                Left = "4000", Right = "4000", FirstLine = "2000"
            });
        }
        source.Position = 0;
        using var fonts = new PinnedCjkFontDirectory();
        var (package, result) = await Convert(source, Options(fonts, engine: true));
        Assert.Equal(label.Replace(" ", ""), Text(package));
        Assert.Single(Assert.Single(package.Pages).Elements.OfType<OfdImageElement>());
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_TEXT_ENGINE_FALLBACK");
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_INDENT_LIMITED");
    }

    [Fact]
    public async Task InvalidTabOriginUsesNativeFailureContract()
    {
        const string body = """
            <w:p><w:pPr><w:ind w:left="60000000"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>
            """;
        using var fonts = new PinnedCjkFontDirectory();
        using var source = DocxReleaseFixtures.CreateMinimalDocx(body);
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new DocxToOfdConverter(Options(fonts, engine: true)).ConvertAsync(source, output));
        Assert.Equal("DOCX native paragraph layout failed: Invalid default tab origin.", error.Message);
        Assert.IsType<Ofdrw.Net.Layout.Text.TextLayoutException>(
            Assert.IsType<InvalidOperationException>(error.InnerException).InnerException);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task BodyPageFieldParagraphsUseFinalLayout()
    {
        const string body = """
            <w:p><w:r><w:t>Before</w:t></w:r></w:p>
            <w:p><w:r><w:t xml:space="preserve">Body page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple>
              <w:r><w:t xml:space="preserve"> 页</w:t></w:r></w:p>
            <w:p><w:r><w:br/><w:t xml:space="preserve">After a leading break, page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>
            """;
        using var fonts = new PinnedCjkFontDirectory();
        using var source = DocxReleaseFixtures.CreateMinimalDocx(body);
        var (package, _) = await Convert(source, Options(fonts, engine: true));
        var text = Text(package);
        Assert.Contains("Bodypage1页", text);
        Assert.Contains("Afteraleadingbreak,page1", text);
        Assert.All(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>(),
            element => Assert.Contains("ClusterMapV1", element.SourceXml!));
    }

    [Fact]
    public async Task StrictBoldPolicyStillRefusesSynthesisWithTheEngine()
    {
        using var fonts = new PinnedCjkFontDirectory();
        var options = Options(fonts, engine: true);
        options.NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.Throw;
        using var source = File.OpenRead(Sample());
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new DocxToOfdConverter(options).ConvertAsync(source, output));
        Assert.Contains("bold", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, output.Length);
    }

    private static string Sample()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return Path.Combine(root.FullName, "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx");
    }
}
