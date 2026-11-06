using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Converter.Docx.Tests;

// Issue #109: DOCX Native OFD opts into the final paragraph layout engine.
public sealed class DocxNativeParagraphLayoutTests
{
    private static DocxConversionOptions Options(PinnedCjkFontDirectory fonts, bool engine)
    {
        var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false, OfdMode = DocxToOfdMode.Native };
        options.FontDirectories.Add(fonts.Path);
        if (engine) options.NativeParagraphLayoutEngine = HarfBuzzTextPrelayout.CreateParagraphLayout();
        return options;
    }

    private static async Task<(OfdDocumentPackage Package, DocxToOfdConversionResult Result)> Convert(Stream source, DocxConversionOptions options)
    {
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(source, output);
        output.Position = 0;
        return (await new OfdReader().ReadAsync(output), result);
    }

    private static string Text(OfdDocumentPackage package) =>
        string.Concat(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(text => text.Text)).Replace(" ", "");

    [Fact]
    public async Task GeneratedLayoutSampleKeepsItsTextAndUsesClusterMappedFinalLayout()
    {
        using var fonts = new PinnedCjkFontDirectory();
        using var legacySource = File.OpenRead(Sample());
        var (legacy, _) = await Convert(legacySource, Options(fonts, engine: false));
        using var finalSource = File.OpenRead(Sample());
        var (final, result) = await Convert(finalSource, Options(fonts, engine: true));
        Assert.Equal("ParagraphLayout", final.CustomTags["docx-native-text-engine"]);
        Assert.False(legacy.CustomTags.ContainsKey("docx-native-text-engine"));
        // The same logical text, in the same order; only geometry and font binding change.
        Assert.Equal(Text(legacy), Text(final));
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
        var (package, _) = await Convert(source, Options(fonts, engine: true));
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
        Assert.True(paths.Count(path => path.Fill) >= 3);
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
