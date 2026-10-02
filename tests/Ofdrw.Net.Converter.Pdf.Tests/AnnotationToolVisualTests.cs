using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class AnnotationToolVisualTests
{
    public AnnotationToolVisualTests() => PdfFontRegistry.EnsureInstalled();

    [Fact]
    public async Task ShearedAppearance_IsClippedToExactPolygonBeforeAndAfterMix()
    {
        var source = await Annotated($"<PageBlock ID='901'><ImageObject ID='902' ResourceID='MEDIA' Boundary='0 0 200 10'/></PageBlock>", "10 10 20 10", "1 0 0.5 1 0 0");
        var image = Assert.IsType<OfdImageElement>(Assert.Single(source.Pages[0].AnnotationAppearances));
        Assert.Contains("M 0 0 L 20 0 L 20 10 L 0 10 C", image.ClippingXml);
        foreach (var package in new[] { source, OfdDocumentMixer.Mix([new(source, 0)]) })
        {
            using var ofd = await Write(package); using var pdf = new MemoryStream();
            await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using var document = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d));
            using var page = document.GetPageReader(0); var pixels = page.GetImage();
            int RedAt(double x, double y) => Pixel(pixels, page.GetPageWidth(), x, y, 2);
            int GreenAt(double x, double y) => Pixel(pixels, page.GetPageWidth(), x, y, 1);
            Assert.InRange(RedAt(20, 15), 245, 255); Assert.InRange(GreenAt(20, 15), 0, 10);
            Assert.InRange(GreenAt(12, 18), 245, 255); // outside parallelogram, inside its AABB
            Assert.InRange(GreenAt(40, 15), 245, 255); // oversized child must be clipped
            ofd.Position = 0; using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(ofd, svg);
            Assert.Contains("clip-path", Encoding.UTF8.GetString(svg.ToArray()));
        }
    }

    [Fact]
    public async Task OuterShear_PreservesImageOwnBoundaryClipWhenItsCtmExceedsThatBoundary()
    {
        var source = await Annotated("<ImageObject ID='902' ResourceID='MEDIA' Boundary='0 0 10 10' CTM='20 0 0 10 0 0'/>", "10 10 60 20", "1 0 0.5 1 0 0");
        foreach (var package in new[] { source, OfdDocumentMixer.Mix([new(source, 0)]) })
        {
            using var ofd = await Write(package); using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using var document = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d)); using var page = document.GetPageReader(0); var pixels = page.GetImage();
            Assert.InRange(Pixel(pixels, page.GetPageWidth(), 19, 15, 1), 0, 10);
            Assert.InRange(Pixel(pixels, page.GetPageWidth(), 22, 11, 1), 245, 255);
        }
    }

    [Fact]
    public async Task ExtraAppearanceMetadata_PreservesDrawableArtworkButStillBlocksMix()
    {
        var source = await Annotated("<PathObject ID='901' Boundary='0 0 10 10' Fill='true' Stroke='false'><FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 10 0 L 10 10 L 0 10 C</AbbreviatedData></PathObject>", "10 10 20 20", extraAttributes: "xmlns:v='urn:vendor' v:Style='keep'");
        Assert.Contains(source.Pages[0].AnnotationAppearances, element => element is OfdRawElement);
        Assert.Single(source.Pages[0].AnnotationAppearances.OfType<OfdPathElement>());
        using var ofd = await Write(source); using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var doc = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d)); using var page = doc.GetPageReader(0); var pixels = page.GetImage();
        Assert.InRange(Pixel(pixels, page.GetPageWidth(), 15, 15, 2), 245, 255); Assert.InRange(Pixel(pixels, page.GetPageWidth(), 15, 15, 1), 0, 10);
        ofd.Position = 0; using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(ofd, svg); Assert.Contains("255,0,0", Encoding.UTF8.GetString(svg.ToArray()).Replace(" ", ""));
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(source, 0)]));
    }

    [Fact]
    public async Task NestedUnknownDrawing_CannotExportAPartialSuccessfulAppearance()
    {
        var source = await Annotated("<PathObject ID='901' Boundary='0 0 5 5' Fill='true'><AbbreviatedData>M 0 0 L 5 0 L 5 5 C</AbbreviatedData></PathObject><CompositeObject ID='902' ResourceID='999' Boundary='0 0 10 10'/>", "10 10 20 20");
        Assert.IsType<OfdRawElement>(Assert.Single(source.Pages[0].AnnotationAppearances));
        using var ofd = await Write(source); using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf));
        ofd.Position = 0; using var svg = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(ofd, svg));
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(source, 0)]));
    }

    [Fact]
    public async Task DuplicateAnnotationAppearancesFailExportWithoutDroppingKnownArtwork()
    {
        var source = await Annotated("<PathObject Boundary='0 0 5 5' Fill='true'><AbbreviatedData>M 0 0 L 5 0 L 5 5 C</AbbreviatedData></PathObject>", "10 10 20 20");
        var xml = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries["Doc_0/Annots/Page.xml"])); var ns = xml.Root!.Name.Namespace;
        var annotation = xml.Root.Element(ns + "Annot")!; annotation.Add(new XElement(annotation.Element(ns + "Appearance")!));
        var bytes = Encoding.UTF8.GetBytes(xml.ToString()); source.PreservedEntries["Doc_0/Annots/Page.xml"] = bytes;
        using var ofd = await Write(source); source = await new OfdReader().ReadAsync(ofd);
        Assert.Contains(source.Pages[0].AnnotationAppearances, element => element is OfdRawElement { LocalName: "UnsupportedAnnotationAppearance" });
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(source, 0)]));
        using var saved = await Write(source); var reread = await new OfdReader().ReadAsync(saved); Assert.Equal(bytes, reread.PreservedEntries["Doc_0/Annots/Page.xml"]);
        saved.Position = 0; using var pdf = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(saved, pdf)); Assert.Equal(0, pdf.Length);
        saved.Position = 0; using var svg = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(saved, svg)); Assert.Equal(0, svg.Length);
    }

    [Fact]
    public async Task TextMatrix_ScalesGlyphsAndRotatesThemInsteadOfOnlyMovingAnchors()
    {
        async Task<(int Width, int Height)> Bounds(double[] matrix)
        {
            var source = new OfdDocumentPackage();
            source.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Test", Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts/narrow.ttf")) });
            var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
            var text = new OfdTextElement { Text = "ABCD", FontName = "Test", FontResourceId = "10", XMillimeters = 40, YMillimeters = 40, FontSizeMillimeters = 6, Transform = matrix };
            text.Runs.Add(new OfdTextRun { Text = "ABCD", YMillimeters = 6 }); page.Elements.Add(text); source.Pages.Add(page);
            using var ofd = await Write(source); using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using var doc = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d)); using var rasterPage = doc.GetPageReader(0);
            var pixels = rasterPage.GetImage(); var width = rasterPage.GetPageWidth(); var height = rasterPage.GetPageHeight();
            var points = new List<(int X, int Y)>();
            for (var y = 0; y < height; y++) for (var x = 0; x < width; x++)
            {
                var offset = (y * width + x) * 4;
                if (pixels[offset + 3] > 100 && pixels[offset] < 100 && pixels[offset + 1] < 100 && pixels[offset + 2] < 100) points.Add((x, y));
            }
            Assert.NotEmpty(points); return (points.Max(p => p.X) - points.Min(p => p.X) + 1, points.Max(p => p.Y) - points.Min(p => p.Y) + 1);
        }
        var plain = await Bounds([1, 0, 0, 1, 0, 0]); var scaled = await Bounds([2, 0, 0, 2, 0, 0]); var rotated = await Bounds([0, 1, -1, 0, 0, 0]);
        Assert.InRange(scaled.Width, plain.Width * 2 - 3, plain.Width * 2 + 3); Assert.InRange(scaled.Height, plain.Height * 2 - 3, plain.Height * 2 + 3);
        Assert.InRange(rotated.Width, plain.Height - 3, plain.Height + 3); Assert.InRange(rotated.Height, plain.Width - 3, plain.Width + 3);
    }

    [Theory]
    [InlineData("<TextObject Size='4'><TextCode X='0' Y='4'>OK<Note>SECRET</Note></TextCode></TextObject>")]
    [InlineData("<TextObject Size='4'><TextCode X='0' Y='4'>OK</TextCode><Note><TextCode>SECRET</TextCode></Note></TextObject>")]
    [InlineData("<PathObject><AbbreviatedData>M 0 0 L 1 0<Note> L 9 9</Note></AbbreviatedData></PathObject>")]
    [InlineData("<TextObject Size='4'><Clips><Clip><Area><Path><AbbreviatedData>M 0 0<Note> L 9 9</Note></AbbreviatedData></Path></Area></Clip></Clips><TextCode X='0' Y='4'>OK</TextCode></TextObject>")]
    [InlineData("<PathObject DrawParam='12'><AbbreviatedData>M 0 0 L 1 0</AbbreviatedData></PathObject>")]
    [InlineData("<PathObject><FillColor Value='1 2 3' ColorSpace='12'/><AbbreviatedData>M 0 0 L 1 0</AbbreviatedData></PathObject>")]
    [InlineData("<ImageObject ResourceID='MEDIA' ImageMask='123' Boundary='0 0 10 10'/>")]
    [InlineData("<PageBlock xmlns:v='urn:vendor' v:ID='private.bin'><TextObject Size='4'><TextCode X='0' Y='4'>OK</TextCode></TextObject></PageBlock>")]
    public async Task SameNamespaceUnknownAnnotationChildren_BlockPartialExportAndRemainPreserved(string artwork)
    {
        var source = await Annotated(artwork, "0 0 20 20");
        Assert.IsType<OfdRawElement>(Assert.Single(source.Pages[0].AnnotationAppearances));
        Assert.DoesNotContain("SECRET", new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(source));
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(source, 0)]));
        using var ofd = await Write(source); var saved = await new OfdReader().ReadAsync(ofd);
        Assert.Equal(source.PreservedEntries["Doc_0/Annots/Page.xml"], saved.PreservedEntries["Doc_0/Annots/Page.xml"]);
        ofd.Position = 0; using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf));
        ofd.Position = 0; using var svg = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(ofd, svg));
    }

    [Fact]
    public async Task HiddenGraphicUnits_KeepXmlAndMixButDoNotPaintOrExtract()
    {
        var source = await Annotated("<PathObject Name='hidden' Visible='false' Boundary='0 0 10 10' Fill='true' Stroke='false'><FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></PathObject><TextObject Name='hidden-text' Visible='false' Size='4'><TextCode X='0' Y='4'>HIDDEN</TextCode></TextObject>", "10 10 20 20");
        Assert.DoesNotContain(source.Pages[0].AnnotationAppearances, element => element is OfdRawElement);
        Assert.DoesNotContain("HIDDEN", new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(source));
        var mixed = OfdDocumentMixer.Mix([new(source, 0)]); Assert.DoesNotContain("HIDDEN", new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(mixed));
        foreach (var package in new[] { source, mixed })
        {
            using var ofd = await Write(package); using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using var doc = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d)); using var page = doc.GetPageReader(0);
            Assert.InRange(Pixel(page.GetImage(), page.GetPageWidth(), 15, 15, 1), 245, 255);
            ofd.Position = 0; using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(ofd, svg); Assert.DoesNotContain("HIDDEN", Encoding.UTF8.GetString(svg.ToArray()));
        }
        using var savedOfd = await Write(source); var saved = await new OfdReader().ReadAsync(savedOfd);
        Assert.Equal(source.PreservedEntries["Doc_0/Annots/Page.xml"], saved.PreservedEntries["Doc_0/Annots/Page.xml"]);
    }

    [Fact]
    public async Task AnnotationClip_StandardPathAttributesAndAreaStartExportAndMix()
    {
        var source = await Annotated("<PathObject Boundary='0 0 20 20' Fill='true' Stroke='false'><Clips><Clip><Area Start='0 0' CTM='1 0 0 1 0 0'><Path Name='clip' Visible='true' Stroke='false' Fill='true' LineWidth='0.35' Alpha='255'><AbbreviatedData>M 0 0 L 10 0 L 10 10 L 0 10 C</AbbreviatedData></Path></Area></Clip></Clips><FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>", "10 10 20 20");
        Assert.Single(source.Pages[0].AnnotationAppearances.OfType<OfdPathElement>());
        var mixed = OfdDocumentMixer.Mix([new(source, 0)]);
        foreach (var package in new[] { source, mixed })
        {
            using var ofd = await Write(package); using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using var doc = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d)); using var page = doc.GetPageReader(0);
            Assert.InRange(Pixel(page.GetImage(), page.GetPageWidth(), 15, 15, 1), 0, 10);
            Assert.InRange(Pixel(page.GetImage(), page.GetPageWidth(), 25, 25, 1), 245, 255);
            ofd.Position = 0; using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(ofd, svg);
            Assert.Contains("clipPath", Encoding.UTF8.GetString(svg.ToArray()));
        }
    }

    [Fact]
    public async Task OrdinaryClip_DoesNotUseNestedExtensionPathText()
    {
        var source = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdPathElement { XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 20, HeightMillimeters = 20, Fill = true, Stroke = false, FillColor = new OfdColor(255,0,0), AbbreviatedData = "M 0 0 L 20 0 L 20 20 L 0 20 C",
            ClippingXml = $"<Clips xmlns='{source.Options.Namespace}'><Clip><Area><Path><AbbreviatedData>M 0 0 L 10 0 L 10 10 L 0 10 C<Note>M 0 0 L 20 0 L 20 20 L 0 20 C</Note></AbbreviatedData></Path></Area></Clip><v:Clip xmlns:v='urn:vendor'><v:Area><v:Path><v:AbbreviatedData>M 0 0 L 1 0 L 1 1 C</v:AbbreviatedData></v:Path></v:Area></v:Clip></Clips>" }); source.Pages.Add(page);
        using var ofd = await Write(source); using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd,pdf);
        using var doc = DocLib.Instance.GetDocReader(pdf.ToArray(),new PageDimensions(2d)); using var raster = doc.GetPageReader(0);
        Assert.InRange(Pixel(raster.GetImage(),raster.GetPageWidth(),15,15,1),0,10);
        Assert.InRange(Pixel(raster.GetImage(),raster.GetPageWidth(),25,25,1),245,255);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AnnotationImages_MissingIdOrPayloadRejectPartialExportsAndMix(bool missingPayload)
    {
        var id = missingPayload ? "MEDIA" : "UNKNOWN";
        var source=await Annotated($"<PathObject Boundary='0 0 5 5'><AbbreviatedData>M 0 0 L 5 0 L 5 5 C</AbbreviatedData></PathObject><ImageObject ResourceID='{id}' Boundary='0 0 20 20'/>","10 10 20 20");
        if (missingPayload)
        {
            var holder=source.Pages[0].Elements.OfType<OfdImageElement>().Single();
            var payload=source.PreservedEntries.Single(entry=>entry.Value.SequenceEqual(holder.Data)).Key; source.PreservedEntries.Remove(payload);
            var pagePath=source.Pages[0].SourceEntryPath!; var pageXml=XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[pagePath]).TrimStart('\uFEFF'));
            pageXml.Descendants().Where(node=>node.Name.LocalName=="ImageObject").Remove(); source.PreservedEntries[pagePath]=Encoding.UTF8.GetBytes(pageXml.ToString());
            using var raw=new MemoryStream(); using(var zip=new ZipArchive(raw,ZipArchiveMode.Create,true))foreach(var entry in source.PreservedEntries){using var output=zip.CreateEntry(entry.Key).Open();output.Write(entry.Value);} raw.Position=0;
            source=await new OfdReader().ReadAsync(raw);
        }
        Assert.IsType<OfdRawElement>(Assert.Single(source.Pages[0].AnnotationAppearances));
        Assert.Throws<NotSupportedException>(()=>OfdDocumentMixer.Mix([new(source,0)]));
        using var ofd=await Write(source); var saved=await new OfdReader().ReadAsync(ofd); Assert.Equal(source.PreservedEntries["Doc_0/Annots/Page.xml"],saved.PreservedEntries["Doc_0/Annots/Page.xml"]);
        ofd.Position=0;using var pdf=new MemoryStream();await Assert.ThrowsAsync<NotSupportedException>(()=>new OfdToPdfConverter().ConvertAsync(ofd,pdf));Assert.Equal(0,pdf.Length);
        ofd.Position=0;using var svg=new MemoryStream();await Assert.ThrowsAsync<NotSupportedException>(()=>new OfdToSvgConverter().ConvertAsync(ofd,svg));Assert.Equal(0,svg.Length);
    }

    [Fact]
    public async Task AnnotationFiles_MissingOwnedPageFileRejectsExportsAndKeepsIndex()
    {
        var source=await Annotated("<ImageObject ResourceID='MEDIA' Boundary='0 0 20 20'/>","10 10 20 20"); source.PreservedEntries.Remove("Doc_0/Annots/Page.xml");
        using var raw=new MemoryStream();using(var zip=new ZipArchive(raw,ZipArchiveMode.Create,true))foreach(var entry in source.PreservedEntries){using var output=zip.CreateEntry(entry.Key).Open();output.Write(entry.Value);}raw.Position=0;
        source=await new OfdReader().ReadAsync(raw);Assert.Equal("UnsupportedAnnotationAppearance",Assert.IsType<OfdRawElement>(Assert.Single(source.Pages[0].AnnotationAppearances)).LocalName);
        Assert.Throws<NotSupportedException>(()=>OfdDocumentMixer.Mix([new(source,0)]));
        using var ofd=await Write(source);var saved=await new OfdReader().ReadAsync(ofd);Assert.Equal(source.PreservedEntries["Doc_0/Annots/Annotations.xml"],saved.PreservedEntries["Doc_0/Annots/Annotations.xml"]);
        ofd.Position=0;using var pdf=new MemoryStream();await Assert.ThrowsAsync<NotSupportedException>(()=>new OfdToPdfConverter().ConvertAsync(ofd,pdf));Assert.Equal(0,pdf.Length);
        ofd.Position=0;using var svg=new MemoryStream();await Assert.ThrowsAsync<NotSupportedException>(()=>new OfdToSvgConverter().ConvertAsync(ofd,svg));Assert.Equal(0,svg.Length);
    }

    [Theory]
    [InlineData("<Area><v:Path xmlns:v='urn:vendor'><v:AbbreviatedData>M 0 0 L 10 0 L 10 10 C</v:AbbreviatedData></v:Path></Area>")]
    [InlineData("<Area><Path><AbbreviatedData><Note>M 0 0 L 10 0 L 10 10 C</Note></AbbreviatedData></Path></Area>")]
    [InlineData("<Area><Path><AbbreviatedData>M 0 0 L 5 0 L 5 5 C</AbbreviatedData></Path></Area><Area><v:Path xmlns:v='urn:vendor'><v:AbbreviatedData>M 0 0 L 40 0 L 40 40 C</v:AbbreviatedData></v:Path></Area>")]
    [InlineData("<Area><Path><AbbreviatedData>M 0 0 L 5 0 L 5 5 C</AbbreviatedData></Path></Area><v:Area xmlns:v='urn:vendor'><v:Path><v:AbbreviatedData>M 0 0 L 40 0 L 40 40 C</v:AbbreviatedData></v:Path></v:Area>")]
    public async Task OrdinaryClip_WithoutSupportedLiteralPathsFailsBothExports(string areas)
    {
        var source = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdPathElement { XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 20, HeightMillimeters = 20, Fill = true, Stroke = false, FillColor = new OfdColor(255,0,0), AbbreviatedData = "M 0 0 L 20 0 L 20 20 C",
            ClippingXml = $"<Clips xmlns='{source.Options.Namespace}'><Clip>{areas}</Clip></Clips>" }); source.Pages.Add(page);
        using var ofd = await Write(source); using var pdf = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd,pdf)); Assert.Equal(0,pdf.Length);
        ofd.Position = 0; using var svg = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(ofd,svg)); Assert.Equal(0,svg.Length);
        ofd.Position = 0; var saved = await new OfdReader().ReadAsync(ofd); Assert.True(XNode.DeepEquals(XElement.Parse(source.Pages[0].Elements.OfType<OfdPathElement>().Single().ClippingXml!), XElement.Parse(saved.Pages[0].Elements.OfType<OfdPathElement>().Single().ClippingXml!)));
    }

    [Theory]
    [InlineData("Payload='private.bin'")]
    [InlineData("xmlns:v='urn:vendor' v:Font='private.bin'")]
    public async Task UnknownPrimitiveAttributes_KeepKnownArtworkAndBlockMix(string attribute)
    {
        var source = await Annotated($"<PathObject {attribute} Boundary='0 0 10 10' Fill='true' Stroke='false'><FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></PathObject>", "10 10 20 20");
        Assert.Single(source.Pages[0].AnnotationAppearances.OfType<OfdPathElement>());
        Assert.Contains(source.Pages[0].AnnotationAppearances, element => element is OfdRawElement);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(source, 0)]));
        using var ofd = await Write(source); var saved = await new OfdReader().ReadAsync(ofd);
        Assert.Equal(source.PreservedEntries["Doc_0/Annots/Page.xml"], saved.PreservedEntries["Doc_0/Annots/Page.xml"]);
        ofd.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var doc = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d)); using var page = doc.GetPageReader(0);
        Assert.InRange(Pixel(page.GetImage(), page.GetPageWidth(), 15, 15, 2), 245, 255);
    }

    private static int Pixel(byte[] pixels, int width, double x, double y, int channel)
    {
        var offset = ((int)Math.Round(y * 72 / 25.4 * 2) * width + (int)Math.Round(x * 72 / 25.4 * 2)) * 4;
        var alpha = pixels[offset + 3]; return (int)Math.Round(pixels[offset + channel] * alpha / 255d + 255 - alpha);
    }
    private static async Task<MemoryStream> Write(OfdDocumentPackage source)
    {
        var stream = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, stream); stream.Position = 0; return stream;
    }
    private static async Task<OfdDocumentPackage> Annotated(string artwork, string boundary, string? matrix = null, string? extraAttributes = null)
    {
        var source = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        using var red = new Image<Rgba32>(1, 1, Color.Red); using var png = new MemoryStream(); red.SaveAsPng(png);
        page.Elements.Add(new OfdImageElement { Data = png.ToArray(), WidthMillimeters = 1, HeightMillimeters = 1 }); source.Pages.Add(page);
        using var input = await Write(source); var read = await new OfdReader().ReadAsync(input); var ns = read.Options.Namespace;
        var document = XDocument.Parse(Encoding.UTF8.GetString(read.PreservedEntries["Doc_0/Document.xml"]).TrimStart('\uFEFF'));
        document.Root!.Add(new XElement(XName.Get("Annotations", ns), "Annots/Annotations.xml")); read.PreservedEntries["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString());
        read.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{read.Pages[0].Id}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        artwork = artwork.Replace("MEDIA", read.Pages[0].Elements.OfType<OfdImageElement>().Single().ResourceId);
        read.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot ID='900'><Appearance Boundary='{boundary}' {extraAttributes} {(matrix is null ? "" : $"CTM='{matrix}'")}>{artwork}</Appearance></Annot></PageAnnot>");
        using var zipStream = new MemoryStream();
        using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create, true)) foreach (var pair in read.PreservedEntries) { using var target = zip.CreateEntry(pair.Key).Open(); target.Write(pair.Value); }
        zipStream.Position = 0; return await new OfdReader().ReadAsync(zipStream);
    }
}
