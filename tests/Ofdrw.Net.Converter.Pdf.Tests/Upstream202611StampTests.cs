using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Signatures.Signing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class Upstream202611StampTests
{
    private static readonly byte[] Sentinel = [0x53, 0x45, 0x41, 0x4c];

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SelectedAndRepeatedPagesKeepTheirRidingSealSlices(bool nested)
    {
        var signed = await OfdToPdfResultTests.SignedTwoPageSeal(
            nested ? await NestedSeal() : OfdToPdfResultTests.Picture());
        var (result, pdf) = await Convert(signed);
        Assert.Equal(new[] { 0, 1 }, result.SourcePages);
        Assert.Equal(2, result.RenderedSignatureAppearances);
        Assert.Empty(result.Diagnostics);
        AssertPdfSlices(pdf, [true, false]);

        int[] selection = [1, 1, 0];
        var (selected, selectedPdf) = await Convert(signed, selection);
        Assert.Equal(selection, selected.SourcePages);
        Assert.Equal(3, selected.RenderedSignatureAppearances);
        Assert.Empty(selected.Diagnostics);
        AssertPdfSlices(selectedPdf, [false, false, true]);

        await SaveEvidence(nested ? "riding-seal-nested" : "riding-seal-bitmap", signed, pdf);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ClippedBitmapSealEmbedsOnlyItsSourcePixels(bool strict)
    {
        var signed = await OfdToPdfResultTests.SignedTwoPageSeal(OfdToPdfResultTests.Picture(),
            new OfdNormalStampPlacement(1, 10, 10, 40, 20));
        signed = OfdToPdfResultTests.Rewrite(signed, entries =>
        {
            var path = Assert.Single(entries.Keys, name => name.EndsWith("/Signature.xml", StringComparison.Ordinal));
            var xml = XDocument.Parse(Encoding.UTF8.GetString(entries[path]));
            // This synthetic signature has no cryptographic validity. Freeze its
            // date for a byte-identical before/after rendering fixture.
            xml.Descendants().Single(node => node.Name.LocalName == "SignatureDateTime").Value = "2020-01-01T00:00:00Z";
            entries[path] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        });
        await SaveEvidence(strict ? "bitmap-interpolation-strict" : "bitmap-interpolation-default", signed);
        using var input = new MemoryStream(signed);
        using var output = new MemoryStream();
        var converter = new OfdToPdfConverter();
        if (strict)
        {
            // Inspect the same intermediate PDF that strict image export rasterizes.
            var package = await new OfdReader().ReadAsync(input);
            await converter.ConvertPackageAsync(package, output, [0], CancellationToken.None,
                strictAppearanceBudgets: true, maximumSignatureAppearances: 2);
        }
        else
        {
            await converter.ConvertAsync(input, output, [0]);
        }

        using var document = PdfPigDocument.Open(output.ToArray());
        var images = document.GetPage(1).GetImages().OrderBy(image => image.BoundingBox.Left).ToArray();
        Assert.Equal(2, images.Length);
        Assert.Equal((40, 20), (images[0].WidthInSamples, images[0].HeightInSamples));
        Assert.True(images[0].Interpolate);
        Assert.Equal((20, 20), (images[1].WidthInSamples, images[1].HeightInSamples));
        Assert.False(images[1].Interpolate);
    }

    [Fact]
    public async Task SelectedNestedSliceFitsExactPageBudgetAndExportsJpeg()
    {
        var signed = await OfdToPdfResultTests.SignedTwoPageSeal(await NestedSeal());
        const double ppm = 2.125;
        var options = new OfdToImageOptions
        {
            Format = OfdImageFormat.Jpeg,
            PixelsPerMillimeter = ppm,
            MaxSignatureAppearanceCount = 1,
            PackageLoadOptions = new() { MaxPageCount = 3 }
        };

        using var input = new MemoryStream(signed);
        using var output = new MemoryStream();
        await new OfdToImageConverter(options).ConvertAsync(input, output, pageIndex: 1);
        var jpeg = output.ToArray();
        Assert.Equal(new byte[] { 0xff, 0xd8, 0xff }, jpeg[..3]);
        using var raster = Image.Load<Rgb24>(jpeg);
        Assert.Equal((213, 170), (raster.Width, raster.Height));
        AssertSlice((x, y) => raster[(int)(x * ppm), (int)(y * ppm)], red: false);

        // Two outer pages are admissible, but leave no room for the nested page.
        options.PackageLoadOptions.MaxPageCount = 2;
        input.Position = 0;
        using var untouched = SeededOutput();
        var failure = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdToImageConverter(options).ConvertAsync(input, untouched, pageIndex: 1));
        Assert.Contains("signature appearance expansion budget", failure.Message);
        AssertUnchanged(untouched);
        await SaveEvidence("nested-image-budget", signed);
    }

    [Fact]
    public async Task UnreadableNestedSealWarnsOnSelectedSourcePageAndStrictExportLeavesDestinationUntouched()
    {
        // Valid ZIP container that cannot be parsed as OFD. This reaches nested
        // package loading rather than the missing/unsupported payload fallback.
        using var invalidSeal = new MemoryStream();
        using (var zip = new ZipArchive(invalidSeal, ZipArchiveMode.Create, leaveOpen: true))
        using (var entry = zip.CreateEntry("not-ofd.txt").Open())
            entry.WriteByte(42);
        var signed = await OfdToPdfResultTests.SignedTwoPageSeal(invalidSeal.ToArray());

        var (result, pdf) = await Convert(signed, [1]);
        Assert.Equal(new[] { 1 }, result.SourcePages);
        Assert.Equal(0, result.RenderedSignatureAppearances);
        var warning = Assert.Single(result.Diagnostics);
        Assert.Equal("PDF_SIGNATURE_APPEARANCE_UNREADABLE", warning.Code);
        Assert.Equal(OfdToPdfDiagnosticSeverity.Warning, warning.Severity);
        Assert.Equal(1, warning.PageIndex);
        Assert.False(string.IsNullOrWhiteSpace(warning.SignatureId));
        Assert.Contains("nested OFD package", warning.Message);
        using var document = DocLib.Instance.GetDocReader(pdf, new PageDimensions(1d));
        Assert.Equal(1, document.GetPageCount());

        using var input = new MemoryStream(signed);
        using var output = SeededOutput();
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdToImageConverter().ConvertAsync(input, output, pageIndex: 1));
        AssertUnchanged(output);
        await SaveEvidence("unreadable-nested-seal", signed);
    }

    [Fact]
    public async Task CancellationDuringSignedInputReadLeavesDestinationUntouched()
    {
        var signed = await OfdToPdfResultTests.SignedTwoPageSeal(await NestedSeal());
        using var cancellation = new CancellationTokenSource();
        using var input = new CancelAfterFirstByte(signed, cancellation);
        using var output = SeededOutput();
        var failure = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new OfdToImageConverter().ConvertAsync(input, output, pageIndex: 1, cancellationToken: cancellation.Token));
        Assert.Equal(cancellation.Token, failure.CancellationToken);
        Assert.Equal(1, input.Position);
        AssertUnchanged(output);
        await SaveEvidence("cancelled-signed-input", signed);
    }

    private static async Task<byte[]> NestedSeal()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { Index = 0, WidthMillimeters = 40, HeightMillimeters = 20 };
        package.Pages.Add(page);
        for (var half = 0; half < 2; half++)
        {
            page.Elements.Add(new OfdPathElement
            {
                XMillimeters = half * 20,
                WidthMillimeters = 20,
                HeightMillimeters = 20,
                Fill = true,
                Stroke = false,
                FillColor = half == 0 ? new OfdColor(220, 20, 20) : new OfdColor(20, 20, 220),
                AbbreviatedData = "M 0 0 L 20 0 L 20 20 L 0 20 C"
            });
        }

        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        return output.ToArray();
    }

    private static async Task<(OfdToPdfConversionResult Result, byte[] Pdf)> Convert(
        byte[] signed, IReadOnlyList<int>? pages = null)
    {
        using var input = new MemoryStream(signed);
        using var output = new MemoryStream();
        var result = await new OfdToPdfConverter().ConvertWithResultAsync(input, output, pages);
        return (result, output.ToArray());
    }

    private static void AssertPdfSlices(byte[] pdf, bool[] redPages)
    {
        using var document = DocLib.Instance.GetDocReader(pdf, new PageDimensions(3d * 25.4d / 72d));
        Assert.Equal(redPages.Length, document.GetPageCount());
        for (var index = 0; index < redPages.Length; index++)
        {
            using var page = document.GetPageReader(index);
            var width = page.GetPageWidth();
            var pixels = page.GetImage();
            AssertSlice((x, y) =>
            {
                var offset = ((y * 3 * width) + x * 3) * 4;
                return pixels[offset + 3] == 0
                    ? new Rgb24(255, 255, 255)
                    : new Rgb24(pixels[offset + 2], pixels[offset + 1], pixels[offset]);
            }, redPages[index]);
        }
    }

    private static void AssertSlice(Func<int, int, Rgb24> sample, bool red)
    {
        // 40×20 mm stamp, halved and right-aligned on a 100×80 mm page.
        // On both pages the visible rectangle is x=80–100 and y=30–50 mm.
        foreach (var (x, y) in new[] { (83, 33), (96, 46) })
        {
            var pixel = sample(x, y);
            Assert.True((red ? pixel.R : pixel.B) > 180 && pixel.G < 90 &&
                (red ? pixel.B : pixel.R) < 90, $"Wrong seal color at ({x}, {y}): {pixel}");
        }

        foreach (var (x, y) in new[] { (75, 40), (90, 25), (90, 55) })
        {
            var pixel = sample(x, y);
            Assert.True(pixel.R > 235 && pixel.G > 235 && pixel.B > 235, $"Ink outside seal at ({x}, {y}): {pixel}");
        }
    }

    private static MemoryStream SeededOutput()
    {
        var output = new MemoryStream();
        output.Write(Sentinel);
        output.Position = 1;
        return output;
    }

    private static void AssertUnchanged(MemoryStream output)
    {
        Assert.Equal(1, output.Position);
        Assert.Equal(Sentinel, output.ToArray());
    }

    private static async Task SaveEvidence(string name, byte[] input, byte[]? pdf = null)
    {
        var root = Environment.GetEnvironmentVariable("OFDRW_UPSTREAM_202611_STAMP_EVIDENCE");
        if (string.IsNullOrEmpty(root)) return;
        var directory = Path.Combine(root, name);
        Directory.CreateDirectory(directory);
        await File.WriteAllBytesAsync(Path.Combine(directory, "input.ofd"), input);
        if (pdf is not null)
            await File.WriteAllBytesAsync(Path.Combine(directory, "output.pdf"), pdf);
    }

    private sealed class CancelAfterFirstByte(byte[] signed, CancellationTokenSource cancellation)
        : MemoryStream(signed, writable: false)
    {
        public override bool CanSeek => false;

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var read = Read(buffer, offset, Math.Min(count, 1));
            cancellation.Cancel();
            return Task.FromResult(read);
        }
    }
}
