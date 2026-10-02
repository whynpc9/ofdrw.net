using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class ImageDocumentToolsIntegrationTests
{
    private static async Task<byte[]> Write(OfdDocumentPackage package)
    {
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); return output.ToArray();
    }

    private static async Task<(byte[] Bytes, int XmlBytes)> Annotated(string artwork, bool repeated = false, string? matrix = null)
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80, Elements =
        { new OfdImageElement { Data = ImageIoTests.ImageBytes(), FileName = "quad.png", XMillimeters = 90, YMillimeters = 70, WidthMillimeters = 1, HeightMillimeters = 1 } } });
        if (repeated) package.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 100, HeightMillimeters = 80 });
        using var input = new MemoryStream(await Write(package)); var read = await new OfdReader().ReadAsync(input);
        var ns = XNamespace.Get(read.Options.Namespace);
        var document = XDocument.Parse(Encoding.UTF8.GetString(read.PreservedEntries["Doc_0/Document.xml"]).TrimStart('\uFEFF'));
        document.Root!.Add(new XElement(ns + "Annotations", "Annots/Annotations.xml"));
        read.PreservedEntries["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString());
        read.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes(new XElement(ns + "Annotations",
            read.Pages.Select(page => new XElement(ns + "Page", new XAttribute("PageID", page.Id!), new XElement(ns + "FileLoc", "Page.xml")))).ToString());
        var media = read.Pages[0].Elements.OfType<OfdImageElement>().Single().ResourceId!;
        var xml = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot ID='900'><Appearance Boundary='10 10 20 10' {(matrix is null ? "" : $"CTM='{matrix}'")}>{artwork.Replace("MEDIA", media)}</Appearance></Annot></PageAnnot>");
        read.PreservedEntries["Doc_0/Annots/Page.xml"] = xml;
        using var raw = new MemoryStream();
        using (var zip = new ZipArchive(raw, ZipArchiveMode.Create, true)) foreach (var entry in read.PreservedEntries)
        { using var output = zip.CreateEntry(entry.Key).Open(); output.Write(entry.Value); }
        return (raw.ToArray(), xml.Length);
    }

    private const string TwoPaths = "<PathObject Boundary='0 0 5 5' Fill='true' Stroke='false'><FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 5 0 L 5 5 L 0 5 C</AbbreviatedData></PathObject>" +
        "<PathObject Visible='false' Boundary='5 0 5 5' Fill='true'><AbbreviatedData>M 0 0 L 5 0 L 5 5 C</AbbreviatedData></PathObject>";

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task ImageEntryPreservesCumulativeAnnotationBudgetAndSnapshotsNewFields(bool xmlBytes)
    {
        var fixture = await Annotated(TwoPaths, repeated: true);
        var options = new OfdToImageOptions { PixelsPerMillimeter = 2 };
        if (xmlBytes) options.PackageLoadOptions.MaxAnnotationXmlBytes = 2L * fixture.XmlBytes - 1;
        else options.PackageLoadOptions.MaxAnnotationObjectCount = 3;
        var exporter = new OfdToImageConverter(options);
        options.PackageLoadOptions.MaxAnnotationXmlBytes = long.MaxValue; options.PackageLoadOptions.MaxAnnotationObjectCount = int.MaxValue;
        using var input = new MemoryStream(fixture.Bytes); using var target = new MemoryStream(Encoding.ASCII.GetBytes("sentinel"));
        await Assert.ThrowsAsync<InvalidDataException>(() => exporter.ConvertAsync(input, target));
        Assert.Equal("sentinel", Encoding.ASCII.GetString(target.ToArray()));
        input.Position = 0; using var accepted = new MemoryStream();
        await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2, PackageLoadOptions = new()
        { MaxAnnotationObjectCount = 4, MaxAnnotationXmlBytes = 2L * fixture.XmlBytes } }).ConvertAsync(input, accepted);
        using var image = Image.Load<Rgb24>(accepted.ToArray()); Assert.True(image[24, 24].R > 220); Assert.True(image[24, 24].G < 30);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task NestedSignatureRetainsAnnotationBudgetAndRepeatedPayloadCache(bool xmlBytes)
    {
        var nested = await Annotated(TwoPaths);
        var source = await ImageIoTests.WithSeals([nested.Bytes], stampsPerRecord: 2);
        var options = new OfdToImageOptions { PixelsPerMillimeter = 8 };
        if (xmlBytes) options.PackageLoadOptions.MaxAnnotationXmlBytes = nested.XmlBytes - 1;
        else options.PackageLoadOptions.MaxAnnotationObjectCount = 1;
        using var input = new MemoryStream(source); using var target = new MemoryStream(Encoding.ASCII.GetBytes("sentinel"));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(options).ConvertAsync(input, target));
        Assert.Equal("sentinel", Encoding.ASCII.GetString(target.ToArray()));
        input.Position = 0; using var accepted = new MemoryStream();
        options.PackageLoadOptions.MaxAnnotationObjectCount = 2; options.PackageLoadOptions.MaxAnnotationXmlBytes = nested.XmlBytes;
        await new OfdToImageConverter(options).ConvertAsync(input, accepted);
        using var image = Image.Load<Rgb24>(accepted.ToArray());
        // Same nested annotation paints in both reused forms; hidden primitive still charges the reader budget.
        Assert.True(image[50, 92].R > 220 && image[50, 92].G < 30);
        Assert.True(image[194, 92].R > 220 && image[194, 92].G < 30);
    }

    [Fact]
    public async Task ImageExportKeepsExactShearedAnnotationClip()
    {
        var fixture = await Annotated("<PageBlock ID='901'><ImageObject ID='902' ResourceID='MEDIA' Boundary='0 0 200 10'/></PageBlock>", matrix: "1 0 0.5 1 0 0");
        using var input = new MemoryStream(fixture.Bytes); using var output = new MemoryStream();
        await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 4 }).ConvertAsync(input, output);
        using var image = Image.Load<Rgb24>(output.ToArray());
        Assert.True(image[80, 52].R > 220); Assert.True(image[80, 52].G < 30);
        Assert.Equal(Color.White.ToPixel<Rgb24>(), image[48, 72]); // AABB would incorrectly paint here.
        Assert.Equal(Color.White.ToPixel<Rgb24>(), image[160, 60]);
    }

    [Fact]
    public async Task UnsupportedAnnotationFailsWholeImageBeforePublication()
    {
        var fixture = await Annotated(TwoPaths + "<UnsupportedDrawing/>");
        using var input = new MemoryStream(fixture.Bytes); using var target = new MemoryStream(Encoding.ASCII.GetBytes("sentinel"));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToImageConverter().ConvertAsync(input, target));
        Assert.Equal("sentinel", Encoding.ASCII.GetString(target.ToArray()));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task FullGlyphCtmChangesSlopeThroughImageEntryWithFixedBaseline(bool italic)
    {
        async Task<Image<Rgb24>> Render(bool shear)
        {
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Ofdrw Test Face", Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts/narrow.ttf")) });
            var page = new OfdPage { WidthMillimeters = 70, HeightMillimeters = 40 };
            var text = new OfdTextElement { Text = "I", FontName = "Ofdrw Test Face", FontResourceId = "10", Italic = italic,
                FontSizeMillimeters = 8, XMillimeters = 15, YMillimeters = 15, FillColor = new OfdColor(128, 0, 128),
                Transform = shear ? [1, 0, -0.2, 1, 1.6, 0] : [1, 0, 0, 1, 0, 0] };
            text.Runs.Add(new OfdTextRun { Text = "I", YMillimeters = 8 }); page.Elements.Add(text);
            var control = new OfdTextElement { Text = "I", FontName = "Ofdrw Test Face", FontResourceId = "10", FontSizeMillimeters = 8, XMillimeters = 45, YMillimeters = 15 };
            control.Runs.Add(new OfdTextRun { Text = "I", YMillimeters = 8 }); page.Elements.Add(control); package.Pages.Add(page);
            using var input = new MemoryStream(await Write(package)); using var output = new MemoryStream();
            await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 8 }).ConvertAsync(input, output);
            return Image.Load<Rgb24>(output.ToArray());
        }
        double Slope(Image<Rgb24> image)
        {
            var rows = new Dictionary<int, int>();
            for (var y = 0; y < image.Height; y++) for (var x = 0; x < image.Width / 2; x++)
                if (image[x, y].G < 50 && image[x, y].R > 70 && image[x, y].B > 70) { rows.TryAdd(y, x); break; }
            Assert.NotEmpty(rows); var top = rows.Keys.Min() + 5; var bottom = rows.Keys.Max() - 5;
            return (rows[top] - rows[bottom]) / (double)(bottom - top);
        }
        using var before = await Render(false); using var after = await Render(true);
        Assert.InRange(Slope(after) - Slope(before), 0.15, 0.25);
        for (var y = 0; y < before.Height; y++) for (var x = before.Width * 4 / 7; x < before.Width; x++) Assert.Equal(before[x, y], after[x, y]);
    }

    [Fact]
    public async Task ImportedCenteredImagesSurviveWatermarkAndReorderedSplit()
    {
        using var first = new MemoryStream(ImageIoTests.ImageBytes()); using var second = new MemoryStream(ImageIoTests.ImageBytes(true)); using var ofd = new MemoryStream();
        await new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = 2, PageSize = new() { WidthMillimeters = 60, HeightMillimeters = 50 } }).ConvertAsync([first, second], ofd);
        ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
        OfdWatermark.AddImage(package, [0], ImageIoTests.ImageBytes(), "image/png", new OfdWatermarkOptions { XMillimeters = 2, YMillimeters = 2, WidthMillimeters = 8, HeightMillimeters = 4 }, alpha: 255);
        var split = OfdDocumentSplitter.Split(package, [1, 0]); var saved = await Write(split);
        using var input = new MemoryStream(saved); var roundtrip = await new OfdReader().ReadAsync(input);
        Assert.Equal(2, roundtrip.Pages.Count); Assert.Single(roundtrip.Pages[0].Elements.OfType<OfdImageElement>()); Assert.Equal(2, roundtrip.Pages[1].Elements.OfType<OfdImageElement>().Count());
        using var output = new MemoryStream(); input.Position = 0; await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 4 }).ConvertAsync(input, output, 1);
        using var raster = Image.Load<Rgb24>(output.ToArray()); Assert.True(raster[12, 12].R > 220); Assert.True(raster[60, 80].R > 220);
        Assert.Equal(Color.White.ToPixel<Rgb24>(), raster[120, 20]);
    }
}
