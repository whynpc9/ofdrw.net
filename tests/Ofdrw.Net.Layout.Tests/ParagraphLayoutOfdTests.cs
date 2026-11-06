using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class ParagraphLayoutOfdTests
{
    // Shared deterministic sample: Chinese, English, Hebrew, Arabic and Tamil, mixed sizes and colors,
    // alignment, justification, explicit breaks and pagination.
    internal const string SampleText =
        "Paragraph layout 段落排版：中文与 English 混排，Unicode 17.0.0 双向算法、断行与字素簇。" +
        "The quick brown fox jumps over the lazy dog; 敏捷的棕色狐狸跳过了懒狗。\n" +
        "Hebrew in an LTR line: שלום עולם (ver. 2.0) and back to Latin.\n" +
        "مرحبا بالعالم، هذا نص عربي يلتف عبر عدة أسطر مع أرقام 123 وكلمة English داخل السطر.\n" +
        "தமிழ் எழுத்துகள் கூட்டெழுத்துகளுடன் சரியாக வடிவமைக்கப்படுகின்றன க்ஷ ஸ்ரீ.\n" +
        "Line one\u2028line two after a line separator.";

    internal static ParagraphLayoutResult SampleLayout(TextAlignment alignment = TextAlignment.Justify, double? pageHeight = 60)
    {
        var body = ParagraphLayoutTests.Style(4.2);
        var accent = ParagraphLayoutTests.Style(5, new TextPaint(new TextColor(180, 30, 30)));
        var blue = ParagraphLayoutTests.Style(4.2, new TextPaint(new TextColor(20, 60, 200, 200)));
        var accentStart = SampleText.IndexOf("English", StringComparison.Ordinal);
        var blueStart = SampleText.IndexOf("שלום", StringComparison.Ordinal);
        var spans = new[]
        {
            new TextStyleSpan(0, accentStart, body),
            new TextStyleSpan(accentStart, 7, accent),
            new TextStyleSpan(accentStart + 7, blueStart - accentStart - 7, body),
            new TextStyleSpan(blueStart, 9, blue),
            new TextStyleSpan(blueStart + 9, SampleText.Length - blueStart - 9, body),
        };
        return HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(SampleText, spans,
            new ParagraphLayoutOptions(90, pageHeight, alignment: alignment, lineHeight: TextLineHeight.Multiple(1.2),
                paragraphSpacingMillimeters: 2, firstLineIndentMillimeters: 6)));
    }

    [Fact]
    public async Task WriterRoundTripsLogicalTextOriginalGidsAndGeometry()
    {
        var layout = SampleLayout();
        ParagraphLayoutTests.AssertCoverage(layout);
        Assert.True(layout.Pages.Count >= 2);
        var package = ParagraphLayoutOfdWriter.CreateDocument(layout, 110, 80, 10, 10);
        Assert.Equal(layout.Pages.Count, package.Pages.Count);
        var bytes = await Write(package);

        // Serialized form: one CGTransform per cluster, complete glyph-slot deltas.
        using (var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read))
        {
            var objects = zip.Entries.Where(entry => entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal) &&
                    entry.FullName.Contains("Page", StringComparison.Ordinal))
                .SelectMany(entry => { using var stream = entry.Open(); return XDocument.Load(stream).Descendants()
                    .Where(node => node.Name.LocalName == "TextObject").ToArray(); }).ToArray();
            var runs = layout.Lines.SelectMany(line => line.Runs).ToArray();
            Assert.Equal(runs.Length, objects.Length);
            for (var index = 0; index < runs.Length; index++)
            {
                var run = runs[index];
                var transforms = objects[index].Elements().Where(node => node.Name.LocalName == "CGTransform").ToArray();
                var code = objects[index].Elements().Single(node => node.Name.LocalName == "TextCode");
                Assert.Equal(run.Text, code.Value);
                Assert.Equal(run.Glyphs.Select(glyph => glyph.ClusterStart).Distinct().Count(), transforms.Length);
                Assert.Equal(run.Glyphs.Count, transforms.Sum(node => int.Parse(node.Attribute("GlyphCount")!.Value)));
                Assert.Equal(run.Text.Length, transforms.Sum(node => int.Parse(node.Attribute("CodeCount")!.Value)));
                var expectedIds = run.Glyphs.GroupBy(glyph => glyph.ClusterStart).OrderBy(group => group.Key)
                    .SelectMany(group => group).Select(glyph => glyph.GlyphId.ToString());
                Assert.Equal(expectedIds, transforms.SelectMany(node => node.Value.Split(' ')));
                if (run.Glyphs.Count > 1)
                    Assert.Equal(run.Glyphs.Count - 1, code.Attribute("DeltaX")!.Value.Split(' ').Length);
            }
        }

        // Read back: logical text per line is intact and each object projects to its exact GIDs.
        var document = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        var texts = document.Pages.SelectMany(page => page.Elements.OfType<OfdTextElement>()).ToArray();
        Assert.Equal(layout.Lines.SelectMany(line => line.Runs).Select(run => run.Text), texts.Select(text => text.Text));
        var restored = string.Concat(layout.Lines.Select(line =>
            string.Concat(line.Runs.OrderBy(run => run.SourceStart).Select(run => run.Text))));
        Assert.Equal(string.Concat(layout.Lines.Select(line =>
            layout.Text.Substring(line.SourceStart, line.SourceLength - line.TerminatorLength))), restored);
        var fonts = document.Fonts;
        foreach (var (text, run) in texts.Zip(layout.Lines.SelectMany(line => line.Runs)))
        {
            var projected = OfdExplicitGlyphRunProbe.GlyphIds(text, fonts.Single(font => font.Id == text.FontResourceId));
            var expected = run.Glyphs.GroupBy(glyph => glyph.ClusterStart).OrderBy(group => group.Key)
                .SelectMany(group => group).Select(glyph => (int)glyph.GlyphId);
            Assert.Equal(expected, projected);
        }
        // Negative control: the same complete map without the writer-owned marker is a foreign
        // CGTransform and stays on the historical fail-closed profiles.
        var arabic = texts.First(text => text.Text.Contains(' ') && text.Text.Any(value => value is >= '\u0600' and <= '\u06ff'));
        var foreign = new OfdTextElement { Text = arabic.Text, FontResourceId = arabic.FontResourceId,
            SourceXml = arabic.SourceXml!.Replace("ClusterMapV1=\"true\"", "ClusterMapV1=\"false\"") };
        foreach (var run in arabic.Runs) foreign.Runs.Add(run);
        Assert.NotEqual(arabic.SourceXml, foreign.SourceXml);
        Assert.Throws<NotSupportedException>(() => OfdExplicitGlyphRunProbe.GlyphIds(foreign,
            fonts.Single(font => font.Id == arabic.FontResourceId)));
        // Subset embedding keeps the original GIDs.
        Assert.All(fonts, font => Assert.True(font.Data!.Length > 0));
    }

    [Theory]
    [InlineData("start-single-page", TextAlignment.Start, TextParagraphDirection.Auto, null)]
    [InlineData("justify-paginated", TextAlignment.Justify, TextParagraphDirection.Auto, 60.0)]
    [InlineData("rtl-center", TextAlignment.Center, TextParagraphDirection.RightToLeft, null)]
    public async Task PdfAndSvgExportDrawEveryGlyphFromOriginalOutlines(string name, TextAlignment alignment,
        TextParagraphDirection direction, double? pageHeight)
    {
        var layout = direction == TextParagraphDirection.Auto ? SampleLayout(alignment, pageHeight) :
            HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(SampleText,
                ParagraphLayoutTests.Style(4.2), new ParagraphLayoutOptions(90, pageHeight, direction, alignment)));
        var bytes = await Write(ParagraphLayoutOfdWriter.CreateDocument(layout, 110, pageHeight is null ? 160 : 80, 10, 10));
        using var pdf = new MemoryStream();
        await new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        Assert.True(pdf.Length > 1000);
        using var svg = new MemoryStream();
        await new Ofdrw.Net.Converter.Svg.Converters.OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        var svgText = System.Text.Encoding.UTF8.GetString(svg.ToArray());
        // The SVG converter emits one page (the first) per call.
        var firstPage = layout.Pages[0].Lines.SelectMany(line => line.Runs).ToArray();
        var visible = firstPage.SelectMany(run => run.Glyphs).Count(glyph => glyph.InkBoundsMillimeters is not null);
        Assert.True(System.Text.RegularExpressions.Regex.Matches(svgText, "data-original-gid=").Count >= visible);
        Assert.All(firstPage.Where(run => run.Text.Trim().Length > 0), run => Assert.Contains(
            System.Security.SecurityElement.Escape(run.Text)!, svgText));
        // Set OFDRW_PARAGRAPH_ARTIFACTS to a directory to keep the outputs for visual review.
        if (Environment.GetEnvironmentVariable("OFDRW_PARAGRAPH_ARTIFACTS") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            await File.WriteAllBytesAsync(Path.Combine(directory, name + ".ofd"), bytes);
            await File.WriteAllBytesAsync(Path.Combine(directory, name + ".pdf"), pdf.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(directory, name + ".svg"), svg.ToArray());
        }
    }

    [Fact]
    public void WriterRejectsFramesLeavingThePage()
    {
        var tooWide = SampleLayout();
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            ParagraphLayoutOfdWriter.CreateDocument(tooWide, 95, 80, 10, 10)).Error);
    }

    internal static async Task<byte[]> Write(OfdDocumentPackage package)
    {
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        return stream.ToArray();
    }
}

internal static class OfdExplicitGlyphRunProbe
{
    internal static IReadOnlyList<int> GlyphIds(OfdTextElement text, OfdFontResource font) =>
        (OfdExplicitGlyphRun.Read(text, font) ?? throw new InvalidOperationException("No explicit glyph run.")).GlyphIds;
}
