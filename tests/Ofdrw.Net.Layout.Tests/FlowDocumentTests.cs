using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Internal.Flow;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class FlowDocumentTests
{
    [Fact]
    public async Task PublicFlow_MultipageStyledText_RoundTripsWithoutCoordinates()
    {
        var document = new FlowDocument();
        document.Options.PageHeightMillimeters = 70;
        document.Options.MarginTopMillimeters = 8;
        document.Options.MarginBottomMillimeters = 8;
        document.Options.MarginLeftMillimeters = 10;
        document.Options.MarginRightMillimeters = 10;
        document.Options.DefaultFontFamily = "SimSun";
        var paragraph = new Paragraph();
        paragraph.Spans.Add(new Span("中英混排 Alpha "));
        paragraph.Spans.Add(new Span("粗体") { Bold = true });
        paragraph.Spans.Add(new Span(" italic") { Italic = true });
        paragraph.Spans.Add(new Span(" 红色") { Color = new OfdColor(192, 0, 0) });
        document.Blocks.Add(paragraph);
        for (var i = 0; i < 30; i++)
            document.Blocks.Add(new Paragraph($"第{i + 1}段 The public flow layout wraps proportional words and preserves source text."));

        var package = document.Render();
        Assert.True(package.Pages.Count >= 2);
        var first = package.Pages[0].Elements.OfType<OfdTextElement>().ToList();
        var normal = Assert.Single(first, value => value.Text == "中英混排 Alpha ");
        var bold = Assert.Single(first, value => value.Text == "粗体");
        var italic = Assert.Single(first, value => value.Text == " italic");
        var red = Assert.Single(first, value => value.Text == " 红色");
        Assert.Equal(OfdTextElement.DefaultWeight, normal.Weight);
        Assert.Equal(OfdTextElement.BoldWeight, bold.Weight);
        Assert.False(normal.Italic);
        Assert.True(italic.Italic);
        Assert.Equal((192, 0, 0), (red.FillColor.Red, red.FillColor.Green, red.FillColor.Blue));
        Assert.Equal(normal.YMillimeters + normal.Runs[0].YMillimeters,
            red.YMillimeters + red.Runs[0].YMillimeters, 5);
        Assert.All(package.Pages, page => Assert.All(page.Elements.OfType<OfdTextElement>(), value =>
        {
            Assert.True(value.YMillimeters >= 7.99);
            Assert.True(value.YMillimeters + value.HeightMillimeters <= page.HeightMillimeters - 7.99);
        }));

        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var read = await new OfdReader().ReadAsync(stream);
        Assert.Equal(package.Pages.Count, read.Pages.Count);
        Assert.Equal(Compact(package), Compact(read));
        Assert.Contains("中英混排Alpha粗体italic红色", Compact(read));
    }

    [Fact]
    public void LatinWordAcrossSpans_StaysTogetherAndUsesProportionalAdvances()
    {
        var document = new FlowDocument();
        document.Options.PageWidthMillimeters = 53;
        document.Options.MarginLeftMillimeters = 5;
        document.Options.MarginRightMillimeters = 5;
        var paragraph = new Paragraph();
        paragraph.Spans.Add(new Span("prefix "));
        paragraph.Spans.Add(new Span("infor"));
        paragraph.Spans.Add(new Span("mation") { Bold = true });
        paragraph.Spans.Add(new Span(" Alpha"));
        document.Blocks.Add(paragraph);
        var package = document.Render();
        var text = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().ToArray();
        var info = Assert.Single(text, value => value.Text == "infor");
        var suffix = Assert.Single(text, value => value.Text == "mation");
        Assert.Equal(info.YMillimeters, suffix.YMillimeters);
        var alpha = Assert.Single(text, value => value.Text.Contains("Alpha"));
        var glyphs = alpha.Text.ToCharArray();
        var advances = alpha.Runs[0].DeltaX!.Split(' ').Select(double.Parse).ToArray();
        Assert.True(advances[Array.IndexOf(glyphs, 'A')] > advances[Array.IndexOf(glyphs, 'l')]);
    }

    [Fact]
    public void TailSpacing_AndRepeatedRender_DoNotCreateBlankPageOrMutateEarlierResult()
    {
        var document = new FlowDocument();
        document.Options.PageHeightMillimeters = 30;
        document.Options.MarginTopMillimeters = 5;
        document.Options.MarginBottomMillimeters = 5;
        var paragraph = new Paragraph("First") { SpaceAfterMillimeters = 20 };
        document.Blocks.Add(paragraph);
        var first = document.Render();
        Assert.Single(first.Pages);
        paragraph.Spans[0].Text = "Second";
        var second = document.Render();
        Assert.Single(second.Pages);
        Assert.Equal("First", Assert.Single(first.Pages[0].Elements.OfType<OfdTextElement>()).Text);
        Assert.Equal("Second", Assert.Single(second.Pages[0].Elements.OfType<OfdTextElement>()).Text);
    }

    [Fact]
    public void TrailingExplicitPageBreak_IsDeferredUntilThereIsFollowingContent()
    {
        var document = new FlowDocument();
        document.Blocks.Add(new Paragraph("First\f"));
        Assert.Single(document.Render().Pages);
        document.Blocks.Add(new Paragraph("Second"));
        var pages = document.Render().Pages;
        Assert.Equal(2, pages.Count);
        Assert.Equal("First", Assert.Single(pages[0].Elements.OfType<OfdTextElement>()).Text);
        Assert.Equal("Second", Assert.Single(pages[1].Elements.OfType<OfdTextElement>()).Text);
    }

    [Fact]
    public void ConsecutiveExplicitBreaks_AndPageBreakBefore_KeepTheirPageCount()
    {
        var consecutive = new FlowDocument();
        consecutive.Blocks.Add(new Paragraph("A\f\fB"));
        var pages = consecutive.Render().Pages;
        Assert.Equal(3, pages.Count);
        Assert.Equal("A", Assert.Single(pages[0].Elements.OfType<OfdTextElement>()).Text);
        Assert.Empty(pages[1].Elements);
        Assert.Equal("B", Assert.Single(pages[2].Elements.OfType<OfdTextElement>()).Text);

        var before = new FlowDocument();
        before.Blocks.Add(new Paragraph("A\f"));
        before.Blocks.Add(new Paragraph("B") { PageBreakBefore = true });
        Assert.Equal(3, before.Render().Pages.Count);
    }

    [Fact]
    public void HangulAndSupplementaryHan_UseOneEmAdvances()
    {
        Assert.True(FlowTextMetrics.IsCjkTypographicUnit("한"));
        Assert.True(FlowTextMetrics.IsCjkTypographicUnit("한"));
        Assert.True(FlowTextMetrics.IsCjkTypographicUnit("𠀀"));
        Assert.True(FlowTextMetrics.IsCjkTypographicUnit("𰀀"));
        var document = new FlowDocument();
        document.Blocks.Add(new Paragraph("한𠀀𰀀A"));
        var text = Assert.Single(Assert.Single(document.Render().Pages).Elements.OfType<OfdTextElement>());
        var advances = text.Runs[0].DeltaX!.Split(' ')
            .Select(value => double.Parse(value, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(3, advances.Length);
        Assert.Equal(document.Options.DefaultFontSizeMillimeters, advances[0], 5);
        Assert.Equal(document.Options.DefaultFontSizeMillimeters, advances[1], 5);
        Assert.Equal(document.Options.DefaultFontSizeMillimeters, advances[2], 5);

        var decomposed = new FlowDocument();
        decomposed.Blocks.Add(new Paragraph("한A"));
        var run = Assert.Single(Assert.Single(decomposed.Render().Pages)
            .Elements.OfType<OfdTextElement>()).Runs[0];
        var decomposedAdvances = run.DeltaX!.Split(' ');
        Assert.Single(decomposedAdvances);
        Assert.Equal(decomposed.Options.DefaultFontSizeMillimeters,
            double.Parse(decomposedAdvances[0], System.Globalization.CultureInfo.InvariantCulture), 5);
    }

    [Fact]
    public void ParagraphSpacing_TravelsWithNextLineAcrossPageBoundary()
    {
        var document = new FlowDocument();
        document.Options.PageHeightMillimeters = 30;
        document.Options.MarginTopMillimeters = 5;
        document.Options.MarginBottomMillimeters = 5;
        document.Blocks.Add(new Paragraph("A") { SpaceAfterMillimeters = 12 });
        document.Blocks.Add(new Paragraph("B"));
        var pages = document.Render().Pages;
        Assert.Equal(2, pages.Count);
        Assert.Equal("A", Assert.Single(pages[0].Elements.OfType<OfdTextElement>()).Text);
        var second = Assert.Single(pages[1].Elements.OfType<OfdTextElement>());
        Assert.Equal("B", second.Text);
        Assert.Equal(17, second.YMillimeters, 4);

        var before = new FlowDocument();
        before.Options.PageHeightMillimeters = 30;
        before.Options.MarginTopMillimeters = 5;
        before.Options.MarginBottomMillimeters = 5;
        before.Blocks.Add(new Paragraph("A"));
        before.Blocks.Add(new Paragraph("B") { SpaceBeforeMillimeters = 18 });
        var beforePages = before.Render().Pages;
        Assert.Equal(2, beforePages.Count);
        Assert.Single(beforePages[0].Elements.OfType<OfdTextElement>());
        var moved = Assert.Single(beforePages[1].Elements.OfType<OfdTextElement>());
        Assert.Equal("B", moved.Text);
        Assert.True(moved.YMillimeters + moved.HeightMillimeters <= 25.000001);
    }

    [Fact]
    public void ExplicitPageBreak_DiscardsPriorParagraphTailSpacing()
    {
        var before = new FlowDocument();
        before.Options.PageHeightMillimeters = 30;
        before.Options.MarginTopMillimeters = 5;
        before.Options.MarginBottomMillimeters = 5;
        before.Blocks.Add(new Paragraph("A") { SpaceAfterMillimeters = 12 });
        before.Blocks.Add(new Paragraph("B") { PageBreakBefore = true });
        var pages = before.Render().Pages;
        Assert.Equal(2, pages.Count);
        Assert.Equal(5, Assert.Single(pages[1].Elements.OfType<OfdTextElement>()).YMillimeters, 4);

        var inline = new FlowDocument();
        inline.Options.PageHeightMillimeters = 30;
        inline.Options.MarginTopMillimeters = 5;
        inline.Options.MarginBottomMillimeters = 5;
        inline.Blocks.Add(new Paragraph("A") { SpaceAfterMillimeters = 12 });
        inline.Blocks.Add(new Paragraph("\fB"));
        var inlinePages = inline.Render().Pages;
        Assert.Equal(2, inlinePages.Count);
        Assert.Equal(5, Assert.Single(inlinePages[1].Elements.OfType<OfdTextElement>()).YMillimeters, 4);
    }

    [Fact]
    public void FullWidthImage_WithFirstLineIndent_IsScaledIntoAvailableLine()
    {
        var marker = new object();
        var style = new FlowTextStyle();
        var lines = FlowParagraphLayout.Layout(
            new[] { new FlowInline { Style = style, Image = marker,
                ImageWidthMillimeters = 20, ImageHeightMillimeters = 10 } },
            new FlowParagraphFormat { FirstLineIndentMillimeters = 5 },
            20, new FixedMetrics(), style, default);
        var glyph = Assert.Single(Assert.Single(lines).Glyphs);
        Assert.Same(marker, glyph.Image);
        Assert.Equal(15, glyph.Width, 5);
        Assert.Equal(7.5, glyph.ImageHeight, 5);
        Assert.Throws<InvalidDataException>(() => FlowParagraphLayout.Layout(
            new[] { new FlowInline { Style = style, Image = marker,
                ImageWidthMillimeters = 0, ImageHeightMillimeters = 10 } },
            new FlowParagraphFormat(), 20, new FixedMetrics(), style, default));
    }

    [Fact]
    public void InvalidGeometryOversizeAndBudgets_FailExplicitly()
    {
        var document = new FlowDocument();
        document.Options.MarginLeftMillimeters = 200;
        Assert.Throws<ArgumentException>(() => document.Render());
        document.Options.MarginLeftMillimeters = 10;
        document.Options.PageHeightMillimeters = 14;
        document.Options.MarginTopMillimeters = 5;
        document.Options.MarginBottomMillimeters = 5;
        document.Blocks.Add(new Paragraph("A"));
        Assert.Throws<InvalidOperationException>(() => document.Render());
        document.Options.PageHeightMillimeters = 297;
        ((Paragraph)document.Blocks[0]).Spans[0].Text = "AB";
        document.Options.MaxCharacters = 1;
        Assert.Throws<InvalidOperationException>(() => document.Render());
        document.Options.MaxCharacters = 10;
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.Render(cancelled.Token));
    }

    private static string Compact(OfdDocumentPackage package) =>
        new string(string.Concat(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
            .Select(value => value.Text)).Where(value => !char.IsWhiteSpace(value)).ToArray());

    private sealed class FixedMetrics : IFlowFontMetrics
    {
        public double AdvanceMillimeters(string grapheme, FlowTextStyle style) => 1;
    }
}
