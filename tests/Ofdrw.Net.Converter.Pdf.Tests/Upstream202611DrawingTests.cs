using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using PdfDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed partial class Upstream202611DrawingTests
{
    private const int PixelsPerMillimeter = 6;

    private static async Task<IReadOnlyList<OfdElement>> ReadBlock(byte[] ofd)
    {
        var package = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        return Assert.IsType<OfdRawElement>(package.Pages.Single().Elements[0]).PageBlockChildren!;
    }

    private static IEnumerable<XElement> Nodes(XContainer xml, string name) =>
        xml.Descendants().Where(node => node.Name.LocalName == name);

    private static XDocument SvgFromHtml(string html)
    {
        var start = html.IndexOf("<svg", StringComparison.Ordinal);
        var end = html.IndexOf("</svg>", start, StringComparison.Ordinal) + "</svg>".Length;
        return XDocument.Parse(html[start..end]);
    }

    private static async Task<(byte[] Pdf, XDocument Svg, XDocument Html, byte[] Png)> Render(
        string name, byte[] ofd)
    {
        using var pdf = new MemoryStream();
        using var svg = new MemoryStream();
        using var html = new MemoryStream();
        using var png = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = PixelsPerMillimeter })
            .ConvertAsync(new MemoryStream(ofd), png);
        var svgBytes = svg.ToArray();
        var htmlBytes = html.ToArray();
        var pdfBytes = pdf.ToArray();
        var pngBytes = png.ToArray();

        using var parsed = PdfDocument.Open(pdfBytes);
        Assert.Equal(1, parsed.NumberOfPages);
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } output)
        {
            var dir = Path.Combine(output, name);
            Directory.CreateDirectory(dir);
            File.WriteAllBytes(Path.Combine(dir, "input.ofd"), ofd);
            File.WriteAllBytes(Path.Combine(dir, "output.pdf"), pdfBytes);
            File.WriteAllBytes(Path.Combine(dir, "page-001.svg"), svgBytes);
            File.WriteAllBytes(Path.Combine(dir, "output.html"), htmlBytes);
            File.WriteAllBytes(Path.Combine(dir, "page-001.png"), pngBytes);
        }

        return (pdfBytes, XDocument.Parse(Encoding.UTF8.GetString(svgBytes)),
            SvgFromHtml(Encoding.UTF8.GetString(htmlBytes)), pngBytes);
    }

    private static double Number(XElement node, string name) =>
        double.Parse(node.Attribute(name)!.Value, CultureInfo.InvariantCulture);

    // All samples are at least 1 mm inside a flat region. Allow 3 byte
    // values per channel for backend alpha rounding and RGB quantization.
    private static void Pixel(Image<Rgb24> raster, double x, double y, Rgb24 expected)
    {
        var actual = raster[(int)(x * PixelsPerMillimeter), (int)(y * PixelsPerMillimeter)];
        Assert.True(Math.Abs(actual.R - expected.R) <= 3 && Math.Abs(actual.G - expected.G) <= 3 &&
            Math.Abs(actual.B - expected.B) <= 3, $"({x},{y}) expected {expected}, got {actual}");
    }

    private static void RetainInput(string name, byte[] ofd)
    {
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is not { Length: > 0 } output) return;
        var dir = Path.Combine(output, "fixtures");
        Directory.CreateDirectory(dir);
        File.WriteAllBytes(Path.Combine(dir, name + ".ofd"), ofd);
    }

    private static async Task RejectInAllOutputs(byte[] ofd, string diagnostic)
    {
        Func<Stream, Stream, Task>[] convert =
        [
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output),
            (input, output) => new OfdToImageConverter().ConvertAsync(input, output)
        ];
        foreach (var entry in convert)
        {
            using var sink = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => entry(new MemoryStream(ofd), sink));
            Assert.Contains(diagnostic, error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, sink.Length);
        }
    }
}
