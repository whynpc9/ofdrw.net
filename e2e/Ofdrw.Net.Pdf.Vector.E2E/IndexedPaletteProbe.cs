using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using UglyToad.PdfPig;

internal static class IndexedPaletteProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        const string body = "q 420 0 0 595 0 0 cm /Im1 Do Q";
        const string space = "[/Indexed /DeviceRGB 1 <1f5da6e9f1f9>]";
        byte[] indices = { 0, 1, 0, 1, 1, 0 };
        var rgb = indices.SelectMany(index => index == 0 ? new byte[] { 31, 93, 166 } : new byte[] { 233, 241, 249 }).ToArray();
        var rows = new List<object>();
        foreach (var (name, declared) in new[] { ("absent", (bool?)null), ("false", (bool?)false), ("true", (bool?)true) })
        {
            using var fixture = new PdfFixture(font);
            var source = fixture.Create(new[] { body }, imageBytes: indices, imageWidth: 3, imageHeight: 2, imageColorSpace: space, interpolate: declared);
            var stem = "indexed-" + name; File.WriteAllBytes(Path.Combine(output, stem + "-source.pdf"), source);
            using var input = new MemoryStream(source); using var ofd = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage }).ConvertWithResultAsync(input, ofd);
            if (result.Pages[0].IsNative || !result.Pages[0].Diagnostic.StartsWith("PDFV_ORIGINAL_IMAGE_PAGE")) throw new Exception("Indexed page misclassified.");
            ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
            if (package.Pages[0].Elements.Count != 1) throw new Exception("Indexed partial native content or text overlay.");
            var image = package.Pages[0].Elements.OfType<OfdImageElement>().Single();
            if (XElement.Parse(image.SourceXml!).Attribute(XName.Get("PdfInterpolateV1", "https://ofdrw.net/image-hints"))!.Value != (declared == true ? "true" : "false")) throw new Exception("Indexed interpolation hint changed.");
            using (var decoded = Image.Load<Rgb24>(image.Data))
            {
                if (decoded.Width != 3 || decoded.Height != 2) throw new Exception("Indexed sampling grid changed.");
                for (var i = 0; i < indices.Length; i++) if (!decoded[i % 3, i / 3].Equals(new Rgb24(rgb[i * 3], rgb[i * 3 + 1], rgb[i * 3 + 2]))) throw new Exception("Indexed palette/orientation changed.");
            }
            File.WriteAllBytes(Path.Combine(output, stem + ".ofd"), ofd.ToArray()); ofd.Position = 0;
            var path = Path.Combine(output, stem + ".pdf"); using (var pdf = File.Create(path)) await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using (var pdf = PdfDocument.Open(path))
            {
                var exported = pdf.GetPage(1).GetImages().Single();
                if (exported.Interpolate != (declared == true) || !exported.TryGetBytesAsMemory(out var decoded) || !decoded.Span.SequenceEqual(rgb)) throw new Exception("Indexed PDF samples/interpolation changed.");
            }
            using var strictInput = new MemoryStream(source); using var strictOutput = new MemoryStream(); strictOutput.Write(new byte[] { 7, 8 });
            try { await new PdfVectorToOfdConverter().ConvertAsync(strictInput, strictOutput); throw new Exception("Strict Fail admitted raster image."); }
            catch (NotSupportedException) { if (!strictOutput.ToArray().SequenceEqual(new byte[] { 7, 8 })) throw new Exception("Strict Fail changed destination."); }
            rows.Add(new { name, declared, rgbSha256 = Convert.ToHexStringLower(SHA256.HashData(rgb)), sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)), result });
        }
        File.WriteAllText(Path.Combine(output, "indexed-palette-report.json"), JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true }));
    }
}
