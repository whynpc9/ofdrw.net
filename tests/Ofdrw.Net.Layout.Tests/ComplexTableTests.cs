using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class ComplexTableTests
{
    private static readonly byte[] Pixel = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");

    [Fact]
    public async Task RowSpanAndColumnSpan_MoveAsOneGroupAndPreserveMergedGrid()
    {
        var flow = SmallPage(60);
        var table = new Table();
        foreach (var width in new[] { 20d, 20d, 20d }) table.ColumnWidthsMillimeters.Add(width);
        var before = new Row { MinimumHeightMillimeters = 18 };
        before.Cells.Add(new Cell("before") { ColumnSpan = 3 }); table.Rows.Add(before);
        var first = new Row { MinimumHeightMillimeters = 14 };
        first.Cells.Add(new Cell("A") { RowSpan = 2, BackgroundColor = new OfdColor(220, 230, 250) });
        first.Cells.Add(new Cell("B") { ColumnSpan = 2 }); table.Rows.Add(first);
        var second = new Row { MinimumHeightMillimeters = 14 };
        second.Cells.Add(new Cell("C") { ColumnSpan = 2 }); table.Rows.Add(second);
        flow.Blocks.Add(table);
        var placements = new List<FlowElementPlacement>();
        flow.OnElement = (_, item) => placements.Add(item);
        var package = flow.Render();
        Assert.Equal(2, package.Pages.Count);
        Assert.Equal("before", PageText(package, 0));
        Assert.Equal("ABC", PageText(package, 1));
        Assert.Equal(new[] { 0, 1, 1 }, placements.Select(item => item.Boundary.PageIndex));
        var fill = Assert.Single(package.Pages[1].Elements.OfType<OfdPathElement>(), path => path.Fill);
        Assert.Equal(10, fill.XMillimeters, 5);
        Assert.Equal(20, fill.WidthMillimeters, 5);
        Assert.Equal(28, fill.HeightMillimeters, 5);
        // The row seam is absent through the vertically merged first column.
        Assert.DoesNotContain(package.Pages[1].Elements.OfType<OfdPathElement>(), path =>
            path.Stroke && path.YMillimeters == 10 && path.AbbreviatedData.StartsWith("M 0.1 13.9 L"));
        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        Assert.Equal("beforeABC", string.Concat((await new OfdReader().ReadAsync(stream)).Pages
            .SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(element => element.Text)));
    }

    [Fact]
    public void LongRow_SplitsAtMeasuredLines_AndKeepsEachSourceTokenOnce()
    {
        var flow = SmallPage(40);
        var table = new Table { SplitRowsAcrossPages = true };
        var row = new Row();
        var cell = new Cell(string.Join("\n", Enumerable.Range(1, 9).Select(index => $"行{index}")))
        { BackgroundColor = new OfdColor(244, 248, 252) };
        cell.Paragraphs[0].Spans[0].FontSizeMillimeters = 4;
        row.Cells.Add(cell); table.Rows.Add(row);
        flow.Blocks.Add(table);
        flow.Blocks.Add(new Paragraph("after"));
        var fragments = new List<FlowBoundary>();
        flow.OnElement = (_, placement) => { if (ReferenceEquals(placement.Block, table)) fragments.Add(placement.Boundary); };
        var package = flow.Render();
        Assert.True(package.Pages.Count >= 2);
        Assert.True(fragments.Count >= 2);
        var text = string.Concat(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(e => e.Text));
        Assert.Equal(string.Concat(Enumerable.Range(1, 9).Select(index => $"行{index}")) + "after", text);
        Assert.All(package.Pages.SelectMany(page => page.Elements), element =>
            Assert.True(element.YMillimeters + element.HeightMillimeters <= 30.000001));
        foreach (var pageIndex in fragments.Select(fragment => fragment.PageIndex).Distinct())
            Assert.Contains(package.Pages[pageIndex].Elements.OfType<OfdPathElement>(), path => path.Fill);
    }

    [Fact]
    public void HeaderRepeatsOnlyOnContinuationPages_AndCallbacksMatchPhysicalRows()
    {
        var flow = SmallPage(50);
        var table = new Table { RepeatHeaderRowCount = 1 };
        var header = new Row { MinimumHeightMillimeters = 7 };
        header.Cells.Add(new Cell("H") { BackgroundColor = new OfdColor(220, 230, 250) }); table.Rows.Add(header);
        for (var i = 1; i <= 6; i++)
        {
            var body = new Row { MinimumHeightMillimeters = 10 };
            body.Cells.Add(new Cell("B" + i)); table.Rows.Add(body);
        }
        flow.Blocks.Add(table);
        var fragments = new List<FlowBoundary>();
        flow.OnElement = (_, placement) => fragments.Add(placement.Boundary);
        var package = flow.Render();
        Assert.Equal(3, package.Pages.Count);
        Assert.Equal(9, fragments.Count);
        for (var page = 0; page < package.Pages.Count; page++)
        {
            Assert.StartsWith("H", PageText(package, page));
            Assert.True(PageText(package, page).Length > 1);
        }
        Assert.Equal(3, package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Count(e => e.Text == "H"));
        Assert.Equal("B1B2B3B4B5B6", string.Concat(package.Pages.SelectMany(page => page.Elements)
            .OfType<OfdTextElement>().Where(e => e.Text != "H").Select(e => e.Text)));
    }

    [Fact]
    public async Task ImageCell_UsesR04AdmissionAndRoundTripsWithTextNeighbor()
    {
        var flow = SmallPage(70);
        var table = new Table(); table.ColumnWidthsMillimeters.Add(25); table.ColumnWidthsMillimeters.Add(35);
        var row = new Row();
        row.Cells.Add(new Cell { Image = new ImageBlock(Pixel, 12, 10) { Alignment = ParagraphAlignment.Center },
            BackgroundColor = new OfdColor(245, 245, 245) });
        row.Cells.Add(new Cell("图片 image")); table.Rows.Add(row); flow.Blocks.Add(table);
        var package = flow.Render();
        var image = Assert.Single(package.Pages.Single().Elements.OfType<OfdImageElement>());
        Assert.Equal(16.5, image.XMillimeters, 5);
        Assert.Equal(10.8, image.YMillimeters, 5);
        Assert.Equal("图片 image", PageText(package, 0));
        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        Assert.Single((await new OfdReader().ReadAsync(stream)).Pages.Single().Elements.OfType<OfdImageElement>());
    }

    [Fact]
    public void InvalidComplexGeometryAndBudgets_FailExplicitly()
    {
        var flow = SmallPage(50);
        var table = new Table { RepeatHeaderRowCount = 1 };
        var header = new Row(); header.Cells.Add(new Cell("H") { RowSpan = 2 }); table.Rows.Add(header);
        var body = new Row(); table.Rows.Add(body); flow.Blocks.Add(table);
        Assert.Throws<ArgumentException>(() => flow.Render());
        header.Cells[0].RowSpan = 1;
        body.Cells.Add(new Cell("B"));
        flow.Options.MaxTableGridSlots = 1;
        Assert.Throws<InvalidOperationException>(() => flow.Render());
        flow.Options.MaxTableGridSlots = 10;
        flow.Options.MaxTableFragments = 1;
        Assert.Throws<InvalidOperationException>(() => flow.Render());
        flow.Options.MaxTableFragments = 10;
        body.Cells[0].Image = new ImageBlock(Pixel, 8, 8);
        Assert.Throws<NotSupportedException>(() => flow.Render()); // An image cell cannot also contain paragraphs.
        body.Cells[0].Paragraphs.Clear();
        flow.Options.MaxTableImageBytes = 1;
        Assert.Throws<ArgumentException>(() => flow.Render());
    }

    [Fact]
    public void UnsafeSplitAndOrphanHeader_AreRejectedBeforeReturningAPackage()
    {
        var flow = SmallPage(35); // 15 mm usable height.
        var table = new Table { SplitRowsAcrossPages = true };
        var row = new Row();
        var cell = new Cell(string.Join("\n", Enumerable.Range(1, 8).Select(i => "L" + i)))
        { VerticalAlignment = CellVerticalAlignment.Center };
        cell.Paragraphs[0].Spans[0].FontSizeMillimeters = 4;
        row.Cells.Add(cell); table.Rows.Add(row); flow.Blocks.Add(table);
        Assert.Throws<NotSupportedException>(() => flow.Render());

        row.Cells[0].VerticalAlignment = CellVerticalAlignment.Top;
        row.Cells[0].Paragraphs.Clear();
        row.Cells[0].Image = new ImageBlock(Pixel, 10, 20);
        Assert.Throws<InvalidOperationException>(() => flow.Render());

        row.Cells[0].Image = null;
        table.SplitRowsAcrossPages = false;
        table.RepeatHeaderRowCount = 1;
        row.MinimumHeightMillimeters = 12;
        var body = new Row { MinimumHeightMillimeters = 10 };
        body.Cells.Add(new Cell("body")); table.Rows.Add(body);
        Assert.Throws<InvalidOperationException>(() => flow.Render());
    }

    private static FlowDocument SmallPage(double pageHeight)
    {
        var flow = new FlowDocument();
        flow.Options.PageWidthMillimeters = 100;
        flow.Options.PageHeightMillimeters = pageHeight;
        flow.Options.MarginLeftMillimeters = flow.Options.MarginRightMillimeters = 10;
        flow.Options.MarginTopMillimeters = flow.Options.MarginBottomMillimeters = 10;
        flow.Options.DefaultFontSizeMillimeters = 4;
        return flow;
    }

    private static string PageText(OfdDocumentPackage package, int page) =>
        string.Concat(package.Pages[page].Elements.OfType<OfdTextElement>().Select(element => element.Text));
}
