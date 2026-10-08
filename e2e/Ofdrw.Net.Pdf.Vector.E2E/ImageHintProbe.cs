using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using UglyToad.PdfPig;

internal static class ImageHintProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        byte[] rgb = { 255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0, 0, 0, 0, 255, 255, 255 };
        var hint = XName.Get("PdfInterpolateV1", "https://ofdrw.net/image-hints");
        var packages = new List<OfdDocumentPackage>(); var rows = new List<object>();
        foreach (var (name, declared) in new[] { ("absent", (bool?)null), ("false", (bool?)false), ("true", (bool?)true) })
        {
            using var fixture = new PdfFixture(font);
            var source = fixture.Create(new[] { "q 420 0 0 595 0 0 cm /Im1 Do Q" }, interpolate: declared, imageWidth: 3, imageHeight: 2, imageBytes: rgb);
            File.WriteAllBytes(Path.Combine(output, "image-" + name + "-source.pdf"), source);
            using var input = new MemoryStream(source); using var ofd = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
                Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } }).ConvertWithResultAsync(input, ofd);
            if (result.Pages.Single().IsNative || !result.Pages[0].Diagnostic.Contains("PDFV_ORIGINAL_IMAGE_PAGE")) throw new Exception("Image page misclassified.");
            ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd); packages.Add(package);
            var image = package.Pages.Single().Elements.OfType<OfdImageElement>().Single();
            if (package.Pages[0].Elements.Count != 1 || XElement.Parse(image.SourceXml!).Attribute(hint)!.Value != (declared == true ? "true" : "false")) throw new Exception("Image hint/object lost.");
            using (var pixels = Image.Load<Rgb24>(image.Data))
            {
                if (pixels.Width != 3 || pixels.Height != 2) throw new Exception("Original sampling grid changed.");
                for (var index = 0; index < 6; index++)
                { var pixel = pixels[index % 3, index / 3]; if (pixel.R != rgb[index * 3] || pixel.G != rgb[index * 3 + 1] || pixel.B != rgb[index * 3 + 2]) throw new Exception("Pixel orientation/bytes changed."); }
            }
            File.WriteAllBytes(Path.Combine(output, "image-" + name + ".ofd"), ofd.ToArray());
            var pdfPath = Path.Combine(output, "image-" + name + ".pdf"); ofd.Position = 0;
            await using (var pdf = File.Create(pdfPath)) await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using (var pdf = PdfDocument.Open(pdfPath))
            {
                var exported = pdf.GetPage(1).GetImages().Single();
                if (exported.WidthInSamples != 3 || exported.HeightInSamples != 2 || exported.Interpolate != (declared == true) ||
                    !exported.TryGetBytesAsMemory(out var decoded) || !decoded.Span.SequenceEqual(rgb)) throw new Exception("PDF image samples or interpolation changed.");
            }
            ofd.Position = 0; await using (var svg = File.Create(Path.Combine(output, "image-" + name + ".svg"))) await new OfdToSvgConverter().ConvertAsync(ofd, svg);
            rows.Add(new { name, declared, effective = declared == true, originalRgbSha256 = Convert.ToHexStringLower(SHA256.HashData(rgb)), sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)), result });
        }
        var mixed = OfdDocumentMerger.Merge(new[] { packages[1], packages[2], packages[1] });
        using var combined = new MemoryStream(); await new OfdPackageWriter().WriteAsync(mixed, combined);
        File.WriteAllBytes(Path.Combine(output, "image-flags-mixed.ofd"), combined.ToArray()); combined.Position = 0;
        var mixedPath = Path.Combine(output, "image-flags-mixed.pdf");
        await using (var pdf = File.Create(mixedPath)) await new OfdToPdfConverter().ConvertAsync(combined, pdf);
        using (var pdf = PdfDocument.Open(mixedPath))
            if (!Enumerable.Range(1, 3).Select(index => pdf.GetPage(index).GetImages().Single().Interpolate).SequenceEqual(new[] { false, true, false })) throw new Exception("Same-payload page sampling flags aliased.");
        File.WriteAllText(Path.Combine(output, "image-hint-report.json"), JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true }));
    }
}
