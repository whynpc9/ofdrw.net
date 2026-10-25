using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Signatures.Signing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class StampClipRenderingTests
{
    [Fact]
    public async Task NewlySignedRidingSeal_ShowsOnlyTheAssignedSliceOnEachPage()
    {
        var signed = await SignedTwoPageSeal();
        using var reread = new MemoryStream(signed);
        var package = await new OfdReader().ReadAsync(reread);
        Assert.Contains(package.PreservedEntries.Keys, path =>
            path.EndsWith("/Signature.xml", StringComparison.Ordinal));
        using var input = new MemoryStream(signed);
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(input, pdf);
        Assert.True(pdf.Length > 100);

        for (var index = 0; index < 2; index++)
        {
            using var source = new MemoryStream(signed);
            using var output = new MemoryStream();
            await new OfdToImageConverter(new OfdToImageOptions
            { PixelsPerMillimeter = 3 }).ConvertAsync(source, output, index);
            using var pixels = Image.Load<Rgb24>(output.ToArray());
            var seal = pixels[90 * 3, 40 * 3];
            if (index == 0) Assert.True(seal.R > 180 && seal.B < 90);
            else Assert.True(seal.B > 180 && seal.R < 90);
            var outsideClip = pixels[70 * 3, 40 * 3];
            Assert.True(outsideClip.R > 235 && outsideClip.G > 235 && outsideClip.B > 235);
        }
    }

    [Fact]
    public async Task InvalidForeignClip_IsRejectedByStrictImageExport()
    {
        var signed = await SignedTwoPageSeal();
        using var damaged = new MemoryStream();
        damaged.Write(signed);
        using (var zip = new ZipArchive(damaged, ZipArchiveMode.Update, true))
        {
            var entry = Assert.Single(zip.Entries, item =>
                item.FullName.EndsWith("/Signature.xml", StringComparison.Ordinal));
            XDocument xml;
            using (var stream = entry.Open()) xml = XDocument.Load(stream);
            xml.Descendants().First(node => node.Name.LocalName == "StampAnnot")
                .SetAttributeValue("Clip", "0 0 60 20");
            var path = entry.FullName;
            entry.Delete();
            using var replacement = zip.CreateEntry(path).Open();
            xml.Save(replacement);
        }
        damaged.Position = 0;
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdToImageConverter(new OfdToImageOptions
            { PixelsPerMillimeter = 3 }).ConvertAsync(damaged, output, 0));
        Assert.Equal(0, output.Length);
        damaged.Position = 0;
        var reread = await new OfdReader().ReadAsync(damaged);
        Assert.Single(OfdSignatureAppearanceReader.Read(reread)); // tolerant PDF omits the invalid slice, never the full seal
    }

    private static async Task<byte[]> SignedTwoPageSeal()
    {
        var package = new OfdDocumentPackage();
        for (var index = 0; index < 2; index++) package.Pages.Add(new OfdPage
        { Index = index, WidthMillimeters = 100, HeightMillimeters = 80 });
        using var unsigned = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, unsigned);
        unsigned.Position = 0;
        var options = new OfdSignatureOptions();
        options.StampAppearances.Add(new OfdRidingStampPlacement(OfdStampEdge.Right, 1, 2, 40, 20));
        using var picture = new Image<Rgb24>(40, 20);
        for (var y = 0; y < picture.Height; y++)
        for (var x = 0; x < picture.Width; x++)
            picture[x, y] = x < 20 ? new Rgb24(220, 20, 20) : new Rgb24(20, 20, 220);
        using var image = new MemoryStream();
        picture.SaveAsPng(image);
        using var signed = new MemoryStream();
        await new OfdSignatureService().SignAsync(unsigned, signed,
            new PictureProvider(image.ToArray()), options);
        return signed.ToArray();
    }

    private sealed class PictureProvider(byte[] picture) : IOfdSignatureProvider
    {
        public string ProviderName => "R11 render fixture";
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
