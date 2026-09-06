using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using PdfSharpCore.Fonts;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[CollectionDefinition("Font registry", DisableParallelization = true)]
public sealed class FontRegistryCollection;

[Collection("Font registry")]
public sealed class FontIsolationTests
{
    public FontIsolationTests() => PdfFontRegistry.EnsureInstalled();

    [Fact]
    public void ComposedResolver_ShouldPreserveHostRequestsAndResolveEmbeddedFaces()
    {
        var hostData = File.ReadAllBytes(FontPath("narrow"));
        var host = new HostResolver(hostData);
        var combined = PdfFontRegistry.CreateResolver(host);
        Assert.Equal("host-face", combined.ResolveTypeface("host", false, false).FaceName);
        Assert.Equal(hostData, combined.GetFont("host-face"));
        var family = PdfFontRegistry.RegisterFontFace(File.ReadAllBytes(FontPath("wide")));
        var face = combined.ResolveTypeface(family, false, false);
        Assert.StartsWith("ofd:", face.FaceName);
        using var data = new MemoryStream(combined.GetFont(face.FaceName));
        Assert.StartsWith("Ofdrw-", SixLabors.Fonts.FontDescription.LoadDescription(data).FontFamilyInvariantCulture);
        Assert.Equal(1, host.Requests);
    }

    [Fact]
    public async Task SameNamedFonts_ShouldRemainDistinctAcrossSequentialAndConcurrentDocuments()
    {
        var narrow = await ConvertAsync("narrow");
        var wide = await ConvertAsync("wide");
        Assert.True(wide > narrow * 1.5, $"Font widths were conflated: narrow={narrow}, wide={wide}");
        var widths = await Task.WhenAll(Enumerable.Range(0, 8).Select(index => ConvertAsync(index % 2 == 0 ? "narrow" : "wide")));
        for (var index = 0; index < widths.Length; index++) Assert.Equal(index % 2 == 0 ? narrow : wide, widths[index], precision: 3);
    }

    [Fact]
    public void Registration_ShouldSnapshotBytesAndRejectNameRebinding()
    {
        var original = File.ReadAllBytes(FontPath("narrow"));
        var mutable = (byte[])original.Clone();
        var family = PdfFontRegistry.RegisterFontFace(mutable);
        var retainedBytes = PdfFontRegistry.RegisteredFontBytes;
        Array.Clear(mutable);
        var face = GlobalFontSettings.FontResolver.ResolveTypeface(family, false, false);
        Assert.Equal(original, PdfFontRegistry.GetOriginalFont(face.FaceName));
        Assert.Equal(family, PdfFontRegistry.RegisterFontFace(original));
        Assert.Equal(retainedBytes, PdfFontRegistry.RegisteredFontBytes);
        var name = "legacy-test-" + Guid.NewGuid();
        PdfFontRegistry.RegisterFont(name, original);
        Assert.Throws<InvalidOperationException>(() => PdfFontRegistry.RegisterFont(name, File.ReadAllBytes(FontPath("wide"))));
        Assert.Equal(face.FaceName, GlobalFontSettings.FontResolver.ResolveTypeface(name, false, false).FaceName);
    }

    [Fact]
    public void Registration_ShouldBoundRetainedBytesWithoutEvictingActiveFaces()
    {
        PdfFontRegistry.RegisterFontFace(File.ReadAllBytes(FontPath("narrow")));
        var originalBudget = PdfFontRegistry.MaximumRegisteredFontBytes;
        var retained = PdfFontRegistry.RegisteredFontBytes;
        try
        {
            PdfFontRegistry.MaximumRegisteredFontBytes = retained;
            Assert.Throws<InvalidOperationException>(() => PdfFontRegistry.RegisterFontFace(File.ReadAllBytes(FontPath("budget"))));
            Assert.Equal(retained, PdfFontRegistry.RegisteredFontBytes);
        }
        finally { PdfFontRegistry.MaximumRegisteredFontBytes = originalBudget; }
    }

    private static async Task<double> ConvertAsync(string variant)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "12", FontName = "Ofdrw Test Face", Data = File.ReadAllBytes(FontPath(variant)) });
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 60 };
        page.Elements.Add(new OfdTextElement
        {
            FontName = "Ofdrw Test Face", FontResourceId = "12", FontSizeMillimeters = 5,
            XMillimeters = 10, YMillimeters = 10, Text = "ABCD"
        });
        package.Pages.Add(page);
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        using var output = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, output);
        output.Position = 0;
        using var pdf = PdfPigDocument.Open(output);
        Assert.Equal("ABCD", pdf.GetPage(1).Text);
        var letters = pdf.GetPage(1).Letters;
        var artifactDirectory = Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS");
        if (!string.IsNullOrEmpty(artifactDirectory))
        {
            Directory.CreateDirectory(artifactDirectory);
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, $"font-{variant}.pdf"), output.ToArray());
            await File.WriteAllLinesAsync(Path.Combine(artifactDirectory, $"font-{variant}-letters.txt"), letters.Select(letter =>
                $"{letter.Value}: {letter.Location}; bounds={letter.BoundingBox}; width={letter.Width}"));
        }
        return letters.Max(letter => letter.BoundingBox.Right) - letters.Min(letter => letter.BoundingBox.Left);
    }

    private static string FontPath(string variant) => Path.Combine(AppContext.BaseDirectory, "fonts", variant + ".ttf");

    private sealed class HostResolver(byte[] data) : IFontResolver
    {
        internal int Requests { get; private set; }
        public string DefaultFontName => "host";
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            Requests++;
            return new FontResolverInfo("host-face");
        }
        public byte[] GetFont(string faceName) => data;
    }
}
