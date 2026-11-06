using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

// Issue #109 engine switch: FlowDocument and PlainTextOfdConverter opt into final paragraph layout.
public sealed class FlowTextEngineTests
{
    private static readonly Lazy<FlowTextEngine> Engine = new(() => new FlowTextEngine(HarfBuzzTextPrelayout.CreateParagraphLayout(),
        new[]
        {
            new TextFontFamily("Noto Sans", ParagraphLayoutTests.Face("Noto Sans", "NotoSans-Regular.ttf")),
            new TextFontFamily("LXGW WenKai", ParagraphLayoutTests.Face("LXGW WenKai", "LXGWWenKai-Regular.ttf")),
            new TextFontFamily("DejaVu Sans", ParagraphLayoutTests.Face("Ofdrw Test DejaVu Sans", "DejaVuSans-TextFixture.ttf")),
        }));

    private static FlowDocument Document(bool engine)
    {
        var document = new FlowDocument();
        document.Options.PageWidthMillimeters = 120; document.Options.PageHeightMillimeters = 80;
        document.Options.MarginLeftMillimeters = document.Options.MarginRightMillimeters = 10;
        document.Options.MarginTopMillimeters = document.Options.MarginBottomMillimeters = 10;
        document.Options.DefaultFontFamily = "Noto Sans";
        if (engine) document.Options.TextEngine = Engine.Value;
        var paragraph = new Paragraph { FirstLineIndentMillimeters = 6, Alignment = ParagraphAlignment.Left };
        paragraph.Spans.Add(new Span("Flow text 中文段落 with "));
        paragraph.Spans.Add(new Span("bold") { Bold = true });
        paragraph.Spans.Add(new Span(", "));
        paragraph.Spans.Add(new Span("italic") { Italic = true, Color = new OfdColor(180, 20, 20) });
        paragraph.Spans.Add(new Span(" and ") );
        paragraph.Spans.Add(new Span("underline") { Underline = true });
        paragraph.Spans.Add(new Span(" spans, a\ttab and an Ext-B 张\U00020087 that wrap across several lines of the frame."));
        document.Blocks.Add(paragraph);
        var table = new Table { ColumnWidthsMillimeters = { 40d, 60d } };
        table.Rows.Add(new Row { Cells = { new Cell("单元格 A"), new Cell("Cell B wraps inside its column width.") } });
        document.Blocks.Add(table);
        return document;
    }

    [Fact]
    public async Task DefaultFlowOutputIsUnchangedAndEngineOutputIsClusterMapped()
    {
        var legacy = Document(engine: false).Render();
        Assert.All(legacy.Fonts, font => Assert.Empty(font.Data));
        Assert.All(legacy.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>(),
            text => Assert.All(text.Runs, run => Assert.Null(run.PositionedGlyphs)));

        var final = Document(engine: true).Render();
        var texts = final.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().ToArray();
        Assert.All(texts, text => Assert.Equal(OfdPositionedGlyphMapping.ClusterMap, Assert.Single(text.Runs).PositionedGlyphs!.Mapping));
        Assert.All(final.Fonts, font => Assert.NotEmpty(font.Data));
        Assert.Contains(texts, text => text.Text == "\t");
        Assert.Contains(final.Pages.SelectMany(page => page.Elements).OfType<OfdPathElement>(), path => path.Fill);
        // Logical text survives, in reading order, and the package writes and reads back.
        var expected = "Flow text 中文段落 with bold, italic and underline spans, a\ttab and an Ext-B 张\U00020087 that wrap across several lines of the frame.";
        var bytes = await ParagraphLayoutOfdTests.Write(final);
        var read = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        var readText = read.Pages.SelectMany(page => page.Elements.OfType<OfdTextElement>()).Select(text => text.Text).ToArray();
        Assert.StartsWith(expected.Replace(" ", ""), string.Concat(readText).Replace(" ", ""));
        Assert.Contains("单元格", string.Concat(readText));
        using var pdf = new MemoryStream();
        await new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        Assert.True(pdf.Length > 1000);
        if (Environment.GetEnvironmentVariable("OFDRW_PARAGRAPH_ARTIFACTS") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            await File.WriteAllBytesAsync(Path.Combine(directory, "flow-engine.ofd"), bytes);
            await File.WriteAllBytesAsync(Path.Combine(directory, "flow-engine.pdf"), pdf.ToArray());
            using var legacyStream = new MemoryStream();
            await new Ofdrw.Net.Packaging.OfdPackageWriter().WriteAsync(Document(engine: false).Render(), legacyStream);
            await File.WriteAllBytesAsync(Path.Combine(directory, "flow-legacy.ofd"), legacyStream.ToArray());
        }
    }

    [Fact]
    public void PlainTextKeepsLineAndPageControlSemanticsWithTheEngine()
    {
        string Layout(string text, bool engine, out int pages)
        {
            var options = new PlainTextOfdOptions();
            options.Layout.DefaultFontFamily = "Noto Sans";
            if (engine) options.Layout.TextEngine = Engine.Value;
            var package = new PlainTextOfdConverter(options).Render(text);
            pages = package.Pages.Count;
            return string.Join("|", package.Pages.Select(page => string.Concat(page.Elements.OfType<OfdTextElement>().Select(element => element.Text))));
        }
        foreach (var text in new[] { "", "one\r\ntwo\rthree\nfour", "a\fb\f", "trailing\n\n", "\n", "tab\there\n中文 \U00020087" })
        {
            var legacy = Layout(text, engine: false, out var legacyPages);
            var final = Layout(text, engine: true, out var finalPages);
            Assert.Equal(legacyPages, finalPages);
            Assert.Equal(legacy, final);
        }
        // Line heights follow real font metrics; a long text paginates without losing characters.
        var builder = new StringBuilder();
        for (var line = 0; line < 120; line++) builder.Append("Line ").Append(line).Append(" 中文 mixed text.\n");
        var paged = Layout(builder.ToString(), engine: true, out var count);
        Assert.True(count > 1);
        Assert.Equal(builder.ToString().Replace("\n", ""), paged.Replace("|", ""));
    }

    [Fact]
    public void PlainTextTabsLeaveAtLeastHalfAnEmBeforeTheNextColumn()
    {
        // Issue #115: columns such as "数量 12\t说明" ended just before a default stop and touched.
        var builder = new StringBuilder();
        for (var row = 1; row <= 60; row++)
            builder.Append(row.ToString("00", System.Globalization.CultureInfo.InvariantCulture)).Append("\t项目 Item ").Append(row)
                .Append("\t数量 ").Append(row * 3).Append("\t说明：中英文混排的纯文本行。\n");
        var options = new PlainTextOfdOptions();
        options.Layout.DefaultFontFamily = "Noto Sans";
        options.Layout.TextEngine = Engine.Value;
        var package = new PlainTextOfdConverter(options).Render(builder.ToString());
        var size = options.Layout.DefaultFontSizeMillimeters;
        var tabs = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Where(text => text.Text == "\t").ToArray();
        Assert.Equal(180, tabs.Length);
        Assert.All(tabs, tab => Assert.True(tab.WidthMillimeters >= size / 2 - 1e-6, $"tab advance {tab.WidthMillimeters}"));
    }

    [Fact]
    public void AStyleChangeInsideAGraphemeClusterOnlyRestylesThatCluster()
    {
        var document = new FlowDocument();
        document.Options.TextEngine = Engine.Value;
        var paragraph = new Paragraph();
        paragraph.Spans.Add(new Span("a") { Color = new OfdColor(200, 0, 0) });
        paragraph.Spans.Add(new Span("\u0301XYZ") { Color = new OfdColor(0, 0, 200) });
        document.Blocks.Add(paragraph);
        var texts = document.Render().Pages.Single().Elements.OfType<OfdTextElement>().ToArray();
        var accented = texts.Single(text => text.Text.Contains("a\u0301", StringComparison.Ordinal));
        Assert.Equal(200, accented.FillColor.Red);
        Assert.DoesNotContain("XYZ", accented.Text, StringComparison.Ordinal);
        var rest = texts.Single(text => text.Text.Contains("XYZ", StringComparison.Ordinal));
        Assert.Equal(200, rest.FillColor.Blue);
        Assert.Equal(0, rest.FillColor.Red);
    }

    [Fact]
    public void EngineFailuresSurfaceAsFlowErrorsAndMissingFamiliesFallBack()
    {
        var document = new FlowDocument();
        document.Options.TextEngine = Engine.Value;
        document.Options.DefaultFontFamily = "Unknown Family";
        document.Blocks.Add(new Paragraph("Fallback 中文"));
        var package = document.Render();
        Assert.Contains(package.Fonts, font => font.FontName == "Noto Sans");
        Assert.Contains(package.Fonts, font => font.FontName == "LXGW WenKai");
        var invalid = new FlowDocument();
        invalid.Options.TextEngine = Engine.Value;
        invalid.Blocks.Add(new Paragraph("bidi ‮ override"));
        var error = Assert.Throws<InvalidOperationException>(() => invalid.Render());
        Assert.IsType<TextLayoutException>(error.InnerException);
    }
}
