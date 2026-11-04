using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Signatures.Signing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class OfdToPdfResultTests
{
    private const string Skipped = "PDF_SIGNATURE_APPEARANCE_SKIPPED";
    private const string Unreadable = "PDF_SIGNATURE_APPEARANCE_UNREADABLE";
    private const string MetadataUnreadable = "PDF_SIGNATURE_METADATA_UNREADABLE";

    [Fact]
    public async Task ValidSeal_ReportsRenderedAppearancesAndNoDiagnostics()
    {
        var signed = await SignedTwoPageSeal(Picture());

        var result = await Convert(signed);

        Assert.True(result.RenderedSignatureAppearances >= 2);
        Assert.Empty(result.Diagnostics);
        Assert.Equal(new[] { 0, 1 }, result.SourcePages);
        // Optional visual non-regression input; never written in normal runs.
        var visualDirectory = Environment.GetEnvironmentVariable("OFDRW_SEAL_VISUAL_DIR");
        if (!string.IsNullOrEmpty(visualDirectory))
        {
            Directory.CreateDirectory(visualDirectory);
            await File.WriteAllBytesAsync(Path.Combine(visualDirectory, "seal.ofd"), signed);
        }
    }

    [Fact]
    public async Task UndrawableSeal_ReportsSkippedAppearance()
    {
        // A PNG signature passes payload sniffing; the garbage body fails decoding at draw time.
        var undecodable = new byte[] { 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a }
            .Concat(Enumerable.Repeat((byte)0x42, 64)).ToArray();
        var signed = await SignedTwoPageSeal(undecodable);

        var result = await Convert(signed);

        var diagnostic = Assert.Single(result.Diagnostics, item => item.Code == Skipped && item.PageIndex == 0);
        Assert.Equal(OfdToPdfDiagnosticSeverity.Warning, diagnostic.Severity);
        Assert.NotNull(diagnostic.SignatureId);
        Assert.Equal(0, result.RenderedSignatureAppearances);
    }

    [Fact]
    public async Task UnrepresentableBoundary_ReportsSkippedAppearance()
    {
        var damaged = MutateStamp(await SignedTwoPageSeal(Picture()), stamp => stamp.SetAttributeValue("Boundary", "1e308 1e308 40 20"));

        var result = await Convert(damaged);

        Assert.True(result.Diagnostics.Any(item => item.Code == Skipped && item.Message.Contains("boundary") && item.PageIndex is not null),
            Describe(result));
        Assert.Equal(1, result.RenderedSignatureAppearances); // The other page's stamp still draws.
        await AssertStrictImageExportThrows(damaged);
    }

    [Fact]
    public async Task InvalidClip_ReportsSkippedAppearance()
    {
        var damaged = MutateStamp(await SignedTwoPageSeal(Picture()), stamp => stamp.SetAttributeValue("Clip", "0 0 999 999"));

        var result = await Convert(damaged);

        Assert.Contains(result.Diagnostics, item => item.Code == Skipped && item.Message.Contains("clip"));
        await AssertStrictImageExportThrows(damaged);
    }

    [Fact]
    public async Task MissingPayload_ReportsUnreadableAppearance()
    {
        var signed = await SignedTwoPageSeal(Picture());
        var damaged = Rewrite(signed, entries =>
        {
            var seal = Assert.Single(entries.Keys, path => path.Contains("/Signs/") &&
                !path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && entries[path].Take(4).SequenceEqual(new byte[] { 0x89, 0x50, 0x4e, 0x47 }));
            entries.Remove(seal);
        });

        var result = await Convert(damaged);

        Assert.Contains(result.Diagnostics, item => item.Code == Unreadable && item.Message.Contains("no supported payload") && item.PageIndex is not null);
        Assert.Equal(0, result.RenderedSignatureAppearances);
    }

    [Fact]
    public async Task MalformedSignatureMetadata_ReportsMetadataUnreadable()
    {
        var signed = await SignedTwoPageSeal(Picture());
        var damaged = Rewrite(signed, entries =>
        {
            var list = Assert.Single(entries.Keys, path => path.EndsWith("Signatures.xml", StringComparison.Ordinal));
            entries[list] = Encoding.UTF8.GetBytes("<Broken");
        });

        var result = await Convert(damaged);

        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(MetadataUnreadable, diagnostic.Code);
        Assert.DoesNotContain("<Broken", diagnostic.Message);
    }

    [Fact]
    public async Task NestedSealBudget_UsesCallerOptions()
    {
        var nested = new OfdDocumentPackage();
        for (var index = 0; index < 10; index++) nested.Pages.Add(new OfdPage
        {
            Index = index, WidthMillimeters = 40, HeightMillimeters = 20,
            Elements = { new OfdTextElement { Text = "SEAL 印章", XMillimeters = 2, YMillimeters = 4, WidthMillimeters = 36,
                HeightMillimeters = 10, FontSizeMillimeters = 6, FontName = "Noto Sans CJK SC", FillColor = new OfdColor(200, 0, 0) } }
        });
        using var nestedOfd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(nested, nestedOfd);
        var signed = await SignedTwoPageSeal(nestedOfd.ToArray());
        int mainEntries, nestedEntries;
        using (var zip = new ZipArchive(new MemoryStream(signed))) mainEntries = zip.Entries.Count;
        using (var zip = new ZipArchive(new MemoryStream(nestedOfd.ToArray()))) nestedEntries = zip.Entries.Count;
        Assert.True(nestedEntries > mainEntries + 1, $"nested {nestedEntries} must exceed main {mainEntries} + 1");

        var visualDirectory = Environment.GetEnvironmentVariable("OFDRW_SEAL_VISUAL_DIR");
        if (!string.IsNullOrEmpty(visualDirectory))
        {
            Directory.CreateDirectory(visualDirectory);
            await File.WriteAllBytesAsync(Path.Combine(visualDirectory, "nested-seal.ofd"), signed);
        }
        var permissive = await Convert(signed);
        Assert.Empty(permissive.Diagnostics);
        Assert.Equal(2, permissive.RenderedSignatureAppearances);

        var limited = await Convert(signed, new OfdToPdfOptions
        { PackageLoadOptions = new Ofdrw.Net.Packaging.Archive.OfdPackageLoadOptions { MaxEntryCount = mainEntries + 1 } });

        Assert.Equal(2, limited.Diagnostics.Count(item => item.Code == Unreadable && item.PageIndex is not null));
        Assert.Equal(0, limited.RenderedSignatureAppearances);
    }

    [Fact]
    public async Task DrawGuard_ReportsSkippedPlacement()
    {
        // The stamp Boundary is valid, but a huge page origin makes its page-relative placement unrepresentable.
        var damaged = Rewrite(await SignedTwoPageSeal(Picture()), entries =>
        {
            foreach (var path in entries.Keys.Where(item => item.EndsWith("Document.xml", StringComparison.Ordinal) ||
                         item.EndsWith("Content.xml", StringComparison.Ordinal)).ToArray())
            {
                var xml = XDocument.Parse(Encoding.UTF8.GetString(entries[path]));
                foreach (var box in xml.Descendants().Where(node => node.Name.LocalName == "PhysicalBox")) box.Value = "-1.7e308 0 100 80";
                entries[path] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
            }
        });

        var result = await Convert(damaged);

        Assert.True(2 == result.Diagnostics.Count(item => item.Code == Skipped && item.Message.Contains("placement")), Describe(result));
        Assert.Equal(0, result.RenderedSignatureAppearances);
    }

    [Fact]
    public async Task ConvertAsync_StillReturnsTask()
    {
        using var input = new MemoryStream(await SignedTwoPageSeal(Picture()));
        using var pdf = new MemoryStream();

        await new OfdToPdfConverter().ConvertAsync(input, pdf);

        Assert.True(pdf.Length > 100);
    }

    private static string Describe(OfdToPdfConversionResult result) =>
        $"rendered={result.RenderedSignatureAppearances}; " + string.Join(" | ", result.Diagnostics.Select(item => $"{item} page={item.PageIndex}"));

    internal static async Task<OfdToPdfConversionResult> Convert(byte[] ofd, OfdToPdfOptions? options = null)
    {
        using var input = new MemoryStream(ofd);
        using var pdf = new MemoryStream();
        var converter = options is null ? new OfdToPdfConverter() : new OfdToPdfConverter(options);
        var result = await converter.ConvertWithResultAsync(input, pdf);
        Assert.True(pdf.Length > 100);
        return result;
    }

    private static async Task AssertStrictImageExportThrows(byte[] ofd)
    {
        using var input = new MemoryStream(ofd);
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 3 }).ConvertAsync(input, output, 0));
    }

    private static byte[] MutateStamp(byte[] signed, Action<XElement> mutate) => Rewrite(signed, entries =>
    {
        var path = Assert.Single(entries.Keys, item => item.EndsWith("/Signature.xml", StringComparison.Ordinal));
        var xml = XDocument.Parse(Encoding.UTF8.GetString(entries[path]));
        mutate(xml.Descendants().First(node => node.Name.LocalName == "StampAnnot"));
        entries[path] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
    });

    internal static byte[] Rewrite(byte[] ofd, Action<Dictionary<string, byte[]>> change)
    {
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using (var zip = new ZipArchive(new MemoryStream(ofd)))
            foreach (var entry in zip.Entries)
            {
                using var stream = entry.Open();
                using var copy = new MemoryStream();
                stream.CopyTo(copy);
                entries[entry.FullName] = copy.ToArray();
            }
        change(entries);
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var (name, data) in entries)
            {
                using var stream = zip.CreateEntry(name).Open();
                stream.Write(data);
            }
        return output.ToArray();
    }

    internal static byte[] Picture()
    {
        using var picture = new Image<Rgb24>(40, 20);
        for (var y = 0; y < picture.Height; y++)
        for (var x = 0; x < picture.Width; x++)
            picture[x, y] = x < 20 ? new Rgb24(220, 20, 20) : new Rgb24(20, 20, 220);
        using var image = new MemoryStream();
        picture.SaveAsPng(image);
        return image.ToArray();
    }

    internal static async Task<byte[]> SignedTwoPageSeal(byte[] seal)
    {
        var package = new OfdDocumentPackage();
        for (var index = 0; index < 2; index++) package.Pages.Add(new OfdPage
        { Index = index, WidthMillimeters = 100, HeightMillimeters = 80 });
        using var unsigned = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, unsigned);
        unsigned.Position = 0;
        var options = new OfdSignatureOptions();
        options.StampAppearances.Add(new OfdRidingStampPlacement(OfdStampEdge.Right, 1, 2, 40, 20));
        using var signed = new MemoryStream();
        await new OfdSignatureService().SignAsync(unsigned, signed, new PictureProvider(seal), options);
        return signed.ToArray();
    }

    private sealed class PictureProvider(byte[] picture) : IOfdSignatureProvider
    {
        public string ProviderName => "Result diagnostics fixture";
        public string Company => "Ofdrw.Net Tests";
        public string Version => "1";
        public string SignatureMethod => "urn:ofdrw:r11:test";
        public string SignatureType => "Seal";
        public byte[]? SealData => picture;
        public Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation,
            CancellationToken cancellationToken = default)
            => Task.FromResult(Encoding.UTF8.GetBytes("render-fixture-only"));
    }
}
