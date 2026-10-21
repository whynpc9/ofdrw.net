using System.IO;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class NamedAreaTests
{
    [Fact]
    public void HolderParticipatesInPaginationAndRejectsDuplicateName()
    {
        var flow = new FlowDocument();
        flow.Options.PageHeightMillimeters = 35;
        flow.Options.MarginTopMillimeters = flow.Options.MarginBottomMillimeters = 5;
        flow.Blocks.Add(new Paragraph("BEFORE"));
        flow.Blocks.Add(new AreaHolderBlock("customer", 22));
        flow.Blocks.Add(new Paragraph("AFTER"));
        var package = flow.Render();
        var area = Assert.Single(new OfdNamedAreaEditor().GetAreas(package));
        Assert.Equal(1, area.PageIndex);
        Assert.Equal(5, area.YMillimeters, 5);
        Assert.Equal(22, area.HeightMillimeters);
        Assert.Equal("BEFORE", Assert.Single(package.Pages[0].Elements.OfType<OfdTextElement>()).Text);
        Assert.Equal("AFTER", Assert.Single(package.Pages[^1].Elements.OfType<OfdTextElement>()).Text);
        flow.Blocks.Add(new AreaHolderBlock("customer", 2));
        Assert.Throws<ArgumentException>(() => flow.Render());
    }

    [Fact]
    public async Task TextFillSurvivesTwoRoundTripsWithoutMovingOutsideContent()
    {
        var flow = new FlowDocument();
        flow.Blocks.Add(new Paragraph("HEADER"));
        flow.Blocks.Add(new AreaHolderBlock("name", 18));
        flow.Blocks.Add(new Paragraph("FOOTER"));
        var reopened = await RoundTrip(flow.Render());
        var before = reopened.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>()
            .ToDictionary(e => e.Text, e => (e.XMillimeters, e.YMillimeters));
        var editor = new OfdNamedAreaEditor();
        var area = Assert.Single(editor.GetAreas(reopened));
        Assert.False(area.IsFilled);
        editor.FillText(reopened, "name", "Alice 中文");
        Assert.Throws<InvalidOperationException>(() => editor.FillText(reopened, "name", "again"));
        var final = await RoundTrip(reopened);
        Assert.True(Assert.Single(editor.GetAreas(final)).IsFilled);
        var text = final.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().ToArray();
        Assert.Contains(text, e => e.Text == "Alice 中文");
        foreach (var (value, position) in before)
        {
            var current = Assert.Single(text, e => e.Text == value);
            Assert.Equal(position.XMillimeters, current.XMillimeters, 5);
            Assert.Equal(position.YMillimeters, current.YMillimeters, 5);
        }
    }

    [Fact]
    public async Task ImageFillFitsAndPersists()
    {
        var flow = new FlowDocument();
        flow.Blocks.Add(new AreaHolderBlock("logo", 15));
        var package = await RoundTrip(flow.Render());
        var bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l3sAAAAASUVORK5CYII=");
        var editor = new OfdNamedAreaEditor();
        editor.FillImage(package, "logo", bytes, "image/png", 1, 1);
        var final = await RoundTrip(package);
        var area = Assert.Single(editor.GetAreas(final));
        var image = Assert.Single(final.Pages[area.PageIndex].Elements.OfType<OfdImageElement>());
        Assert.True(area.IsFilled);
        Assert.Equal(bytes, image.Data);
        Assert.True(image.XMillimeters >= area.XMillimeters && image.YMillimeters >= area.YMillimeters);
        Assert.True(image.XMillimeters + image.WidthMillimeters <= area.XMillimeters + area.WidthMillimeters + 0.000001);
        Assert.True(image.YMillimeters + image.HeightMillimeters <= area.YMillimeters + area.HeightMillimeters + 0.000001);
    }

    [Fact]
    public void BadReferencesAndOverflowRejectWithoutMutation()
    {
        var flow = new FlowDocument();
        flow.Blocks.Add(new AreaHolderBlock("tiny", 4));
        var package = flow.Render();
        var editor = new OfdNamedAreaEditor();
        var original = package.CustomTags[OfdNamedAreaIndex.Key];
        Assert.Throws<KeyNotFoundException>(() => editor.FillText(package, "missing", "x"));
        Assert.Throws<InvalidOperationException>(() => editor.FillText(package, "tiny", "A very long string that cannot fit in four millimeters."));
        Assert.Equal(original, package.CustomTags[OfdNamedAreaIndex.Key]);
        Assert.Empty(package.Pages[0].Elements);
        package.CustomTags[OfdNamedAreaIndex.Key] = original.Replace("PageID=\"1\"", "PageID=\"999\"");
        Assert.Throws<InvalidDataException>(() => editor.GetAreas(package));
    }

    [Fact]
    public void InvalidImageAndSignedSourceRejectBeforeFill()
    {
        var flow = new FlowDocument();
        flow.Blocks.Add(new AreaHolderBlock("image", 12));
        var package = flow.Render();
        var editor = new OfdNamedAreaEditor();
        var index = package.CustomTags[OfdNamedAreaIndex.Key];
        Assert.Throws<ArgumentException>(() => editor.FillText(package, "image", "\n"));
        Assert.Throws<InvalidDataException>(() => editor.FillImage(package, "image", new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, "image/png", 1, 1));
        Assert.Equal(index, package.CustomTags[OfdNamedAreaIndex.Key]);
        Assert.Empty(package.Pages[0].Elements);
        package.PreservedDocBodyElements.Add("<Signatures/>");
        Assert.Throws<NotSupportedException>(() => editor.FillText(package, "image", "x"));
        Assert.Equal(index, package.CustomTags[OfdNamedAreaIndex.Key]);
    }

    private static async Task<OfdDocumentPackage> RoundTrip(OfdDocumentPackage package)
    {
        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        return await new OfdReader().ReadAsync(stream);
    }
}
