using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Layout.Internal.Flow;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class FlowFixedBlockTests
{
    private static readonly byte[] Pixel = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");

    private static FlowDocument SmallPage()
    {
        var flow = new FlowDocument();
        flow.Options.PageWidthMillimeters = 50;
        flow.Options.PageHeightMillimeters = 50;
        flow.Options.MarginLeftMillimeters = flow.Options.MarginRightMillimeters = 5;
        flow.Options.MarginTopMillimeters = flow.Options.MarginBottomMillimeters = 5;
        flow.Options.DefaultFontSizeMillimeters = 4;
        return flow;
    }

    [Fact]
    public async Task ImageAndCanvas_ReserveFinalBoxesAndKeepPaintOrderAcrossPages()
    {
        var flow = SmallPage();
        var pages = new List<FlowBoundary>();
        var fragments = new List<FlowElementPlacement>();
        flow.OnPage = (_, page) => pages.Add(page);
        flow.OnElement = (_, fragment) => fragments.Add(fragment);
        var image = new ImageBlock(Pixel, 20, 30) { Alignment = ParagraphAlignment.Center };
        var canvasCalls = 0;
        var canvas = new CanvasBlock(15, 15, (graphics, boundary) =>
        {
            canvasCalls++;
            Assert.Equal(1, boundary.PageIndex);
            graphics.FillRectangle(new OfdBrush(new OfdColor(30, 90, 180)), 0, 0, 15, 15);
        });
        flow.Blocks.Add(image);
        flow.Blocks.Add(canvas);
        flow.Blocks.Add(new Paragraph("After"));
        var result = flow.Render();
        Assert.Equal(2, result.Pages.Count);
        Assert.Equal(new[] { 0, 1 }, pages.Select(page => page.PageIndex));
        Assert.Equal(3, fragments.Count);
        Assert.Equal(1, canvasCalls);
        Assert.Same(image, fragments[0].Block);
        Assert.Equal(15, fragments[0].Boundary.XMillimeters, 5);
        Assert.Equal(5, fragments[0].Boundary.YMillimeters, 5);
        Assert.Same(canvas, fragments[1].Block);
        Assert.Equal(5, fragments[1].Boundary.YMillimeters, 5);
        Assert.Collection(result.Pages[1].Elements,
            element => Assert.IsType<OfdPathElement>(element),
            element => Assert.IsType<OfdTextElement>(element));
        Assert.IsType<OfdImageElement>(Assert.Single(result.Pages[0].Elements));

        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(result, stream);
        stream.Position = 0;
        var reopened = await new OfdReader().ReadAsync(stream);
        Assert.Equal(2, reopened.Pages.Count);
        Assert.IsType<OfdImageElement>(Assert.Single(reopened.Pages[0].Elements));
        Assert.Contains(reopened.Pages[1].Elements, element => element is OfdPathElement);
    }

    [Fact]
    public void BlankCanvasAndFailures_HaveDeterministicCallbacks()
    {
        var flow = SmallPage();
        var pages = 0;
        var fragments = 0;
        flow.OnPage = (_, _) => pages++;
        flow.OnElement = (_, _) => fragments++;
        flow.Blocks.Add(new CanvasBlock(10, 10, (_, _) => { }));
        var rendered = flow.Render();
        Assert.Single(rendered.Pages);
        Assert.Empty(rendered.Pages[0].Elements);
        Assert.Equal(1, pages);
        Assert.Equal(1, fragments);

        pages = fragments = 0;
        flow.Blocks.Add(new ImageBlock(Pixel, 10, 41));
        Assert.Throws<InvalidOperationException>(() => flow.Render());
        Assert.Equal(0, pages);
        Assert.Equal(0, fragments);

        flow.Blocks.RemoveAt(1);
        flow.Blocks.Add(new ImageBlock(Array.Empty<byte>(), 10, 10));
        Assert.Throws<ArgumentException>(() => flow.Render());
        Assert.Equal(0, pages);
        Assert.Equal(0, fragments);

        var failing = SmallPage();
        var drawCalls = 0;
        failing.Blocks.Add(new CanvasBlock(10, 10, (_, _) =>
        { drawCalls++; throw new InvalidOperationException("draw failed"); }));
        Assert.Throws<InvalidOperationException>(() => failing.Render());
        Assert.Equal(1, drawCalls);
        Assert.Throws<InvalidOperationException>(() => failing.Render());
        Assert.Equal(2, drawCalls);
    }

    [Fact]
    public void FixedBlocks_UseExactPageBottomAndRejectOversizeGeometry()
    {
        var flow = SmallPage();
        var atBottom = new CanvasBlock(40, 40, (_, _) => { });
        flow.Blocks.Add(atBottom);
        var placements = new List<FlowElementPlacement>();
        flow.OnElement = (_, placement) => placements.Add(placement);
        Assert.Single(flow.Render().Pages);
        Assert.Equal(5, Assert.Single(placements).Boundary.YMillimeters, 5);
        flow.Blocks.Add(new ImageBlock(Pixel, 40, 1));
        var pages = flow.Render().Pages;
        Assert.Equal(2, pages.Count);
        Assert.Equal(1, placements[^1].Boundary.PageIndex);
        Assert.Equal(5, placements[^1].Boundary.YMillimeters, 5);

        foreach (var (width, height) in new[]
        {
            (41d, 1d), (1d, 41d), (0d, 1d), (double.NaN, 1d),
            (1d, double.PositiveInfinity), (0.0000001d, 1d)
        })
        {
            var invalid = SmallPage();
            invalid.Blocks.Add(new CanvasBlock(width, height, (_, _) => { }));
            Assert.ThrowsAny<Exception>(() => invalid.Render());
        }
    }

    [Fact]
    public void ImageBlock_SnapshotsInputAndRejectsInvalidAlpha()
    {
        var bytes = (byte[])Pixel.Clone();
        var image = new ImageBlock(bytes, 10, 10);
        Array.Fill(bytes, (byte)0);
        var flow = SmallPage();
        flow.Blocks.Add(image);
        Assert.IsType<OfdImageElement>(Assert.Single(Assert.Single(flow.Render().Pages).Elements));
        image.Alpha = 256;
        Assert.Throws<ArgumentOutOfRangeException>(() => flow.Render());
    }

    [Fact]
    public void PageAndElementDrawingCallbacks_UseFinalBoundaryAndPaintOrder()
    {
        var flow = SmallPage();
        flow.Blocks.Add(new Paragraph("A"));
        var calls = new List<string>();
        flow.OnPage = (graphics, page) =>
        {
            calls.Add("page:" + page.PageIndex);
            graphics.FillRectangle(new OfdBrush(new OfdColor(240, 240, 240)), 0, 0,
                page.WidthMillimeters, page.HeightMillimeters);
        };
        flow.OnElement = (graphics, placement) =>
        {
            calls.Add("element:" + placement.Boundary.PageIndex);
            var box = placement.Boundary;
            graphics.DrawLine(new OfdPen(OfdColor.Black), box.XMillimeters,
                box.YMillimeters + box.HeightMillimeters, box.XMillimeters + box.WidthMillimeters,
                box.YMillimeters + box.HeightMillimeters);
        };
        var page = Assert.Single(flow.Render().Pages);
        Assert.Equal(new[] { "page:0", "element:0" }, calls);
        Assert.Collection(page.Elements,
            element => Assert.IsType<OfdPathElement>(element),
            element => Assert.IsType<OfdTextElement>(element),
            element => Assert.IsType<OfdPathElement>(element));
    }

    [Fact]
    public void PageCallbackFailure_StopsLaterCallbacks()
    {
        var flow = SmallPage();
        flow.Blocks.Add(new Paragraph("A"));
        var elementCalls = 0;
        flow.OnPage = (_, _) => throw new InvalidOperationException("page failed");
        flow.OnElement = (_, _) => elementCalls++;
        Assert.Throws<InvalidOperationException>(() => flow.Render());
        Assert.Equal(0, elementCalls);
    }

    [Fact]
    public void SpanSpacingWeightAndUnderline_ShareFinalLineGeometry()
    {
        var flow = SmallPage();
        var paragraph = new Paragraph();
        paragraph.Spans.Add(new Span("AV")
        {
            FontSizeMillimeters = 4, FontWeight = 500, LetterSpacingMillimeters = 1,
            Underline = true, Color = new OfdColor(180, 20, 20)
        });
        paragraph.Spans.Add(new Span(" X"));
        flow.Blocks.Add(paragraph);
        var placed = new List<FlowElementPlacement>();
        flow.OnElement = (_, fragment) => placed.Add(fragment);
        var page = Assert.Single(flow.Render().Pages);
        var text = Assert.Single(page.Elements.OfType<OfdTextElement>(), value => value.Text == "AV");
        var underline = Assert.Single(page.Elements.OfType<OfdPathElement>());
        Assert.Equal(500, text.Weight);
        Assert.Single(placed);
        var baseWidth = FlowLatinMetrics.AdvanceMillimeters("A", new FlowTextStyle { FontSizeMillimeters = 4 }) +
            FlowLatinMetrics.AdvanceMillimeters("V", new FlowTextStyle { FontSizeMillimeters = 4 });
        Assert.Equal(baseWidth + 1, text.WidthMillimeters, 5);
        Assert.Equal(baseWidth + 1, underline.WidthMillimeters, 5);
        Assert.Equal(text.XMillimeters, underline.XMillimeters, 5);
        Assert.Equal(FlowLatinMetrics.AdvanceMillimeters("A", new FlowTextStyle { FontSizeMillimeters = 4 }) + 1,
            double.Parse(text.Runs[0].DeltaX!, System.Globalization.CultureInfo.InvariantCulture), 5);

        var tableFlow = SmallPage();
        var table = new Table();
        var row = new Row();
        var cell = new Cell();
        cell.Paragraphs.Add(new Paragraph("AV"));
        cell.Paragraphs[0].Spans[0].FontWeight = 500;
        cell.Paragraphs[0].Spans[0].LetterSpacingMillimeters = 1;
        cell.Paragraphs[0].Spans[0].Underline = true;
        row.Cells.Add(cell); table.Rows.Add(row); tableFlow.Blocks.Add(table);
        var tablePage = Assert.Single(tableFlow.Render().Pages);
        Assert.Contains(tablePage.Elements.OfType<OfdTextElement>(), value => value.Weight == 500);
        Assert.Contains(tablePage.Elements.OfType<OfdPathElement>(), value => value.Fill == false && value.Stroke);

        paragraph.Spans[0].FontWeight = 950;
        Assert.Throws<ArgumentOutOfRangeException>(() => flow.Render());
    }

    [Fact]
    public void LetterSpacing_ChangesWrappingWithoutTrailingSpacing()
    {
        var style = new FlowTextStyle { FontSizeMillimeters = 4 };
        var baseWidth = FlowLatinMetrics.AdvanceMillimeters("A", style) +
            FlowLatinMetrics.AdvanceMillimeters("V", style);
        var flow = SmallPage();
        flow.Options.PageWidthMillimeters = 10 + baseWidth + 0.5;
        var paragraph = new Paragraph();
        paragraph.Spans.Add(new Span("AV")
        { FontSizeMillimeters = 4, LetterSpacingMillimeters = 1, Underline = true });
        flow.Blocks.Add(paragraph);
        var fragments = new List<FlowElementPlacement>();
        flow.OnElement = (_, fragment) => fragments.Add(fragment);
        var page = Assert.Single(flow.Render().Pages);
        var text = page.Elements.OfType<OfdTextElement>().ToArray();
        var underlines = page.Elements.OfType<OfdPathElement>().ToArray();
        Assert.Equal(new[] { "A", "V" }, text.Select(element => element.Text));
        Assert.Equal(2, underlines.Length);
        Assert.Equal(text.Select(element => element.WidthMillimeters),
            underlines.Select(element => element.WidthMillimeters));
        Assert.Equal(2, fragments.Count);
        Assert.All(text, element => Assert.Null(element.Runs[0].DeltaX));
        Assert.Equal(FlowLatinMetrics.AdvanceMillimeters("A", style), text[0].WidthMillimeters, 5);
    }

    [Fact]
    public void LetterSpacedNonbreakingUnit_FitsAfterWrappingWithoutPreviousLineGap()
    {
        var flow = SmallPage();
        flow.Options.PageWidthMillimeters = 19; // 9 mm usable width.
        var paragraph = new Paragraph();
        paragraph.Spans.Add(new Span("X A\u00A0B")
        { FontSizeMillimeters = 4, LetterSpacingMillimeters = 1 });
        flow.Blocks.Add(paragraph);

        var text = Assert.Single(flow.Render().Pages).Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "X ", "A\u00A0B" }, text.Select(element => element.Text));
        Assert.True(text[1].YMillimeters > text[0].YMillimeters);
        Assert.Equal(5, text[1].XMillimeters, 5);
        Assert.InRange(text[1].WidthMillimeters, 8, 9);
        Assert.Equal(2, text[1].Runs[0].DeltaX!.Split(' ').Length);

        var tableFlow = SmallPage();
        tableFlow.Options.PageWidthMillimeters = 21;
        var table = new Table();
        table.ColumnWidthsMillimeters.Add(11);
        var row = new Row();
        var cell = new Cell { PaddingMillimeters = 1 };
        var cellParagraph = new Paragraph();
        cellParagraph.Spans.Add(new Span("X A\u00A0B")
        { FontSizeMillimeters = 4, LetterSpacingMillimeters = 1 });
        cell.Paragraphs.Add(cellParagraph);
        row.Cells.Add(cell);
        table.Rows.Add(row);
        tableFlow.Blocks.Add(table);
        var tableText = Assert.Single(tableFlow.Render().Pages).Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "X ", "A\u00A0B" }, tableText.Select(element => element.Text));
        Assert.Equal(text[1].WidthMillimeters, tableText[1].WidthMillimeters, 5);
        Assert.Equal(text[1].Runs[0].DeltaX, tableText[1].Runs[0].DeltaX);

        flow.Options.PageWidthMillimeters = 18; // An 8 mm usable width cannot fit the unit itself.
        Assert.Throws<InvalidOperationException>(() => flow.Render());
    }
}
