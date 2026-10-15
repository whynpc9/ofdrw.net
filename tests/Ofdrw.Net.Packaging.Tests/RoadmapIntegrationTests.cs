using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class RoadmapIntegrationTests
{
    [Fact]
    public async Task SubsettingIncludesTextUsedOnlyByTypedAnnotationAppearance()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln"))) directory = directory.Parent;
        var bytes = await File.ReadAllBytesAsync(Path.Combine(directory!.FullName,
            "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new() { Id = "10", FontName = "Noto Sans", Data = bytes });
        var page = new OfdPage(); package.Pages.Add(page);
        page.Elements.Add(new OfdTextElement { FontResourceId = "10", Text = "A", WidthMillimeters = 30, HeightMillimeters = 10 });
        page.Annotations = new();
        var annotation = new OfdAnnotation { Type = OfdAnnotationType.Watermark, Creator = "integration",
            LastModifiedDate = new DateTime(2026, 10, 15),
            Appearance = new() { WidthMillimeters = 30, HeightMillimeters = 10 } };
        annotation.Appearance.Elements.Add(new OfdTextElement { FontResourceId = "10", Text = "Z",
            WidthMillimeters = 30, HeightMillimeters = 10 });
        page.Annotations.Items.Add(annotation);
        using var output = new MemoryStream();
        var report = await new OfdPackageWriter().WriteWithResultAsync(package, output);
        Assert.True(Assert.Single(report.FontEmbedding).IsSubset);
        output.Position = 0; var read = await new OfdReader().ReadAsync(output);
        var cmap = new OpenTypeCmap(new OpenTypeFace(Assert.Single(read.Fonts).Data));
        Assert.NotEqual(0, cmap.Glyph('A')); Assert.NotEqual(0, cmap.Glyph('Z'));
        Assert.Equal("Z", Assert.IsType<OfdTextElement>(read.Pages[0].Annotations!.Items[0].Appearance.Elements[0]).Text);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void ControlledGuardIncludesNewFontOptionsAndCollectionFace(bool changeFace)
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new());
        package.Fonts.Add(new() { Id = "10", FontName = "marker", Data = [1, 2, 3] });
        using var context = new OfdControlledEditContext(new());
        var guard = context.CaptureSource(package);
        if (changeFace) package.Fonts[0].CollectionFaceIndex++;
        else package.Options.FontEmbedding.MaximumUsedScalars--;
        Assert.Throws<InvalidDataException>(() => guard.AssertUnchanged(package, context));
    }
}
