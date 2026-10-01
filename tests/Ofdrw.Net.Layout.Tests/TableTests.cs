using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TableTests
{
    [Fact]
    public void EmptyThinRow_CannotLetBorderStrokeEscapeBounds()
    {
        var document = SmallDocument(); var table = new Table(); var row = new Row { MinimumHeightMillimeters = 0.1 };
        row.Cells.Add(new Cell { PaddingMillimeters = 0 }); table.Rows.Add(row); document.Blocks.Add(table);
        Assert.Throws<ArgumentException>(() => document.Render());
        row.MinimumHeightMillimeters = 0.2;
        Assert.All(document.Render().Pages.Single().Elements, e => Assert.Equal(0.2, e.HeightMillimeters));
    }

    [Fact]
    public void MergedGridFloatingSums_StillInsetLastBorder()
    {
        var document = SmallDocument(); var table = new Table();
        foreach (var width in new[] { 1.1, 29.5, 22.8, 3.7 }) table.ColumnWidthsMillimeters.Add(width);
        var row = new Row(); row.Cells.Add(new Cell("A") { ColumnSpan = 2 }); row.Cells.Add(new Cell("B") { ColumnSpan = 2 });
        table.Rows.Add(row); document.Blocks.Add(table);
        var edge = document.Render().Pages.Single().Elements.OfType<OfdPathElement>().Last();
        var coords = edge.AbbreviatedData.Split(' '); var right = double.Parse(coords[1], System.Globalization.CultureInfo.InvariantCulture);
        Assert.Equal(edge.WidthMillimeters - table.BorderWidthMillimeters / 2, right);
        Assert.Equal("M", coords[0]); Assert.Equal("L", coords[3]); Assert.Equal(coords[1], coords[4]);
    }

    [Fact]
    public async Task MergedCells_KeepFixedWidthsStylesAndRawTextThroughPackageRoundtrip()
    {
        var document = SmallDocument();
        var table = new Table { Alignment = ParagraphAlignment.Center };
        table.ColumnWidthsMillimeters.Add(20); table.ColumnWidthsMillimeters.Add(30); table.ColumnWidthsMillimeters.Add(10);
        var merged = new Cell { ColumnSpan = 2, BackgroundColor = new OfdColor(220, 230, 250) };
        var paragraph = new Paragraph();
        paragraph.Spans.Add(new Span("中文 ")); paragraph.Spans.Add(new Span("Bold") { Bold = true, Color = new OfdColor(200, 0, 0) });
        paragraph.Spans.Add(new Span("Italic") { Italic = true }); merged.Paragraphs.Add(paragraph);
        var row = new Row(); row.Cells.Add(merged); row.Cells.Add(new Cell("Z")); table.Rows.Add(row);
        document.Blocks.Add(table);
        var package = document.Render();
        var elements = package.Pages.Single().Elements;
        var fill = elements.OfType<OfdPathElement>().Single(path => path.Fill);
        Assert.Equal(20, fill.XMillimeters); Assert.Equal(50, fill.WidthMillimeters);
        Assert.Equal("中文 BoldItalicZ", Text(package));
        var bold = elements.OfType<OfdTextElement>().Single(element => element.Text == "Bold");
        Assert.Equal(OfdTextElement.BoldWeight, bold.Weight); Assert.Equal(200, bold.FillColor!.Red);
        Assert.True(elements.OfType<OfdTextElement>().Single(element => element.Text == "Italic").Italic);
        Assert.Equal(OfdTextElement.DefaultWeight, elements.OfType<OfdTextElement>().First().Weight);
        var grid = elements.OfType<OfdPathElement>().Where(path => path.Stroke).ToArray();
        Assert.Equal(5, grid.Length);
        Assert.DoesNotContain(grid, path => path.AbbreviatedData.StartsWith("M 20 "));
        Assert.Contains(grid, path => path.AbbreviatedData.StartsWith("M 50 0 L 50 "));
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0; Assert.Equal(Text(package), Text(await new OfdReader().ReadAsync(output)));
    }

    [Theory]
    [InlineData(ParagraphAlignment.Left, CellVerticalAlignment.Top)]
    [InlineData(ParagraphAlignment.Center, CellVerticalAlignment.Center)]
    [InlineData(ParagraphAlignment.Right, CellVerticalAlignment.Bottom)]
    public void Alignment_UsesPaddedMeasuredContent(ParagraphAlignment horizontal, CellVerticalAlignment vertical)
    {
        var document = SmallDocument(); var table = new Table(); table.ColumnWidthsMillimeters.Add(60);
        var cell = new Cell("Wi") { VerticalAlignment = vertical, PaddingMillimeters = 2 };
        cell.Paragraphs[0].Alignment = horizontal;
        var row = new Row { MinimumHeightMillimeters = 25 }; row.Cells.Add(cell); table.Rows.Add(row); document.Blocks.Add(table);
        var text = document.Render().Pages.Single().Elements.OfType<OfdTextElement>().Single();
        var expectedX = horizontal == ParagraphAlignment.Left ? 12 : horizontal == ParagraphAlignment.Center ?
            10 + (60 - text.WidthMillimeters) / 2 : 68 - text.WidthMillimeters;
        var expectedY = vertical == CellVerticalAlignment.Top ? 12 : vertical == CellVerticalAlignment.Center ?
            10 + (25 - text.HeightMillimeters) / 2 : 33 - text.HeightMillimeters;
        Assert.Equal(expectedX, text.XMillimeters, 5); Assert.Equal(expectedY, text.YMillimeters, 5);
    }

    [Fact]
    public void WholeRows_MoveToNextPage_AndFollowingParagraphContinuesAfterTable()
    {
        var document = SmallDocument(); document.Options.PageHeightMillimeters = 60;
        var table = new Table();
        for (var i = 0; i < 3; i++)
        {
            var row = new Row { MinimumHeightMillimeters = 18 };
            row.Cells.Add(new Cell("A" + i)); row.Cells.Add(new Cell("B" + i)); table.Rows.Add(row);
        }
        document.Blocks.Add(table); document.Blocks.Add(new Paragraph("after"));
        var package = document.Render(); Assert.Equal(2, package.Pages.Count);
        Assert.Equal("A0B0A1B1", string.Concat(package.Pages[0].Elements.OfType<OfdTextElement>().Select(e => e.Text)));
        Assert.Equal("A2B2after", string.Concat(package.Pages[1].Elements.OfType<OfdTextElement>().Select(e => e.Text)));
        foreach (var page in package.Pages)
        {
            Assert.All(page.Elements, element => Assert.InRange(element.YMillimeters + element.HeightMillimeters, 10, 50.000001));
            var text = page.Elements.OfType<OfdTextElement>().Where(e => e.Text != "after").ToArray();
            for (var i = 0; i < text.Length; i += 2) Assert.Equal(text[i].YMillimeters, text[i + 1].YMillimeters);
        }
        var after = package.Pages[1].Elements.OfType<OfdTextElement>().Last(); Assert.Equal(28, after.YMillimeters, 5);
        var grids = package.Pages[0].Elements.OfType<OfdPathElement>().ToArray();
        Assert.Equal(9, grids.Length);
        Assert.DoesNotContain(grids.Where(path => path.YMillimeters == 28), path => path.AbbreviatedData.StartsWith("M 0.1 0.1 L"));
        Assert.Contains("M 0.1 0.1 L", package.Pages[1].Elements.OfType<OfdPathElement>().First().AbbreviatedData);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OversizeRow_FailsRatherThanClippingOrLooping(bool laterRow)
    {
        var document = SmallDocument(); document.Options.PageHeightMillimeters = 40;
        var table = new Table();
        if (laterRow) { var first = new Row(); first.Cells.Add(new Cell("ok")); table.Rows.Add(first); }
        var row = new Row { MinimumHeightMillimeters = 21 }; row.Cells.Add(new Cell("too tall")); table.Rows.Add(row); document.Blocks.Add(table);
        var error = Assert.Throws<InvalidOperationException>(() => document.Render());
        Assert.Contains(laterRow ? "row 2" : "row 1", error.Message); Assert.Contains("21 mm", error.Message);
    }

    [Fact]
    public void CellParagraphs_PreserveIndentsGapsAndNewlineHeights()
    {
        var document = SmallDocument(); var table = new Table(); var row = new Row(); var cell = new Cell();
        var paragraph = new Paragraph("A\nB") { LeftIndentMillimeters = 3, FirstLineIndentMillimeters = 2,
            SpaceBeforeMillimeters = 4, SpaceAfterMillimeters = 5 };
        cell.Paragraphs.Add(paragraph); cell.Paragraphs.Add(new Paragraph("C")); row.Cells.Add(cell); table.Rows.Add(row); document.Blocks.Add(table);
        var text = document.Render().Pages.Single().Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal("ABC", string.Concat(text.Select(e => e.Text)));
        Assert.Equal(15.8, text[0].XMillimeters, 5); Assert.Equal(13.8, text[1].XMillimeters, 5);
        Assert.Equal(14.8, text[0].YMillimeters, 5);
        Assert.Equal(text[0].YMillimeters + text[0].HeightMillimeters, text[1].YMillimeters, 5);
        Assert.Equal(text[1].YMillimeters + text[1].HeightMillimeters + 5, text[2].YMillimeters, 5);
    }

    [Theory]
    [InlineData("short")]
    [InlineData("long")]
    [InlineData("span0")]
    [InlineData("rowspan")]
    [InlineData("padding")]
    [InlineData("width")]
    [InlineData("nan")]
    [InlineData("pagebreak")]
    public void InvalidGeometryAndUnsupportedMerging_AreRejected(string kind)
    {
        var document = SmallDocument(); var table = new Table(); table.ColumnWidthsMillimeters.Add(30); table.ColumnWidthsMillimeters.Add(30);
        var row = new Row(); var cell = new Cell("A"); row.Cells.Add(cell); row.Cells.Add(new Cell("B")); table.Rows.Add(row); document.Blocks.Add(table);
        switch (kind)
        {
            case "short": row.Cells.RemoveAt(1); break;
            case "long": cell.ColumnSpan = 2; break;
            case "span0": cell.ColumnSpan = 0; break;
            case "rowspan": cell.RowSpan = 2; break;
            case "padding": cell.PaddingMillimeters = 15; break;
            case "width": table.ColumnWidthsMillimeters[0] = 90; break;
            case "nan": row.MinimumHeightMillimeters = double.NaN; break;
            case "pagebreak": cell.Paragraphs[0].Spans[0].Text = "A\fB"; break;
        }
        var error = Record.Exception(() => document.Render()); Assert.NotNull(error);
        if (kind is "rowspan" or "pagebreak") Assert.IsType<NotSupportedException>(error);
        else Assert.IsAssignableFrom<ArgumentException>(error);
    }

    [Fact]
    public void PendingNewlinesAndPageBreaks_AreConsumedBeforeTable()
    {
        var document = SmallDocument(); document.Options.PageHeightMillimeters = 28;
        document.Blocks.Add(new Paragraph("A\n\n"));
        var table = new Table { PageBreakBefore = true, SpaceAfterMillimeters = 2 }; var row = new Row(); row.Cells.Add(new Cell("B")); table.Rows.Add(row);
        document.Blocks.Add(table); document.Blocks.Add(new Paragraph("C") { PageBreakBefore = true });
        var package = document.Render(); Assert.Equal(3, package.Pages.Count); Assert.Equal("ABC", Text(package));
        Assert.Equal(10.8, package.Pages[1].Elements.OfType<OfdTextElement>().Single().YMillimeters, 5);
        Assert.Equal(10, package.Pages[2].Elements.OfType<OfdTextElement>().Single().YMillimeters, 5);
    }

    [Fact]
    public void CellAndTextBudgetsCancellationAndRepeatRender_AreEnforced()
    {
        var document = SmallDocument(); var table = new Table(); var row = new Row(); row.Cells.Add(new Cell("one")); row.Cells.Add(new Cell("two"));
        table.Rows.Add(row); document.Blocks.Add(table);
        Assert.Equal(Text(document.Render()), Text(document.Render()));
        document.Options.MaxTableCells = 1; Assert.Throws<InvalidOperationException>(() => document.Render());
        document.Options.MaxTableCells = 10; document.Options.MaxCharacters = 5; Assert.Throws<InvalidOperationException>(() => document.Render());
        document.Options.MaxCharacters = 100; document.Options.MaxTextElements = 1; Assert.Throws<InvalidOperationException>(() => document.Render());
        Assert.Throws<OperationCanceledException>(() => document.Render(new CancellationToken(true)));
        document.Options.MaxTextElements = 100; document.Options.MaxPageCount = 1;
        table.Rows.Add(row); row.MinimumHeightMillimeters = 50;
        Assert.Throws<InvalidOperationException>(() => document.Render());
    }

    private static FlowDocument SmallDocument()
    {
        var document = new FlowDocument(); document.Options.PageWidthMillimeters = 100; document.Options.PageHeightMillimeters = 100;
        document.Options.MarginLeftMillimeters = document.Options.MarginRightMillimeters = 10;
        document.Options.MarginTopMillimeters = document.Options.MarginBottomMillimeters = 10;
        return document;
    }
    private static string Text(OfdDocumentPackage package) => string.Concat(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(e => e.Text));
}
