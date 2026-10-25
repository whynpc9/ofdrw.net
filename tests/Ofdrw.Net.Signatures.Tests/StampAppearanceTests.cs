using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.Tests;

public sealed class StampAppearanceTests
{
    private static readonly byte[] Picture = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData(OfdStampEdge.Right, "88 30 40 20", "0 0 10 20", "68 30 40 20", "10 0 20 20", "30 0 10 20")]
    [InlineData(OfdStampEdge.Left, "-28 30 40 20", "30 0 10 20", "-8 30 40 20", "10 0 20 20", "0 0 10 20")]
    [InlineData(OfdStampEdge.Bottom, "30 73 40 20", "0 0 40 5", "30 63 40 20", "0 5 40 10", "0 15 40 5")]
    [InlineData(OfdStampEdge.Top, "30 -13 40 20", "0 15 40 5", "30 -3 40 20", "0 5 40 10", "0 0 40 5")]
    public async Task SignedAppearance_GroupsAndWeightsUsePageReferencesAndLocalClips(
        OfdStampEdge edge, string firstBoundary, string firstClip,
        string secondBoundary, string secondClip, string thirdClip)
    {
        using var unsigned = await Source();
        var provider = new PictureProvider(Picture);
        var options = new OfdSignatureOptions { CheckMethod = OfdDigestAlgorithms.Sha256Oid };
        options.StampAppearances.Add(new OfdNormalStampPlacement(2, 10, 12, 15, 14));
        var riding = new OfdRidingStampPlacement(edge, 1, 4, 40, 20)
        { GroupSize = 3, MarginMillimeters = 2 };
        foreach (var weight in new[] { 1d, 2d, 1d }) riding.SliceWeights.Add(weight);
        options.StampAppearances.Add(riding);

        using var signed = new MemoryStream();
        var coverage = await new OfdSignatureService().SignWithResultAsync(unsigned, signed, provider, options);
        Assert.Equal(1, provider.SignCalls);
        signed.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(signed);
        var path = Assert.Single(archive.EntryNames,
            name => name.EndsWith("/Signature.xml", StringComparison.Ordinal));
        var exactXml = archive.GetBytes(path);
        Assert.Equal(exactXml, provider.SignedXml);
        var xml = XDocument.Parse(Encoding.UTF8.GetString(exactXml));
        var stamps = xml.Descendants().Where(node => node.Name.LocalName == "StampAnnot").ToArray();
        Assert.Equal(5, stamps.Length);
        Assert.Equal(5, stamps.Select(stamp => stamp.Attribute("ID")!.Value).Distinct().Count());
        Assert.Equal("10 12 15 14", stamps[0].Attribute("Boundary")!.Value);
        Assert.Null(stamps[0].Attribute("Clip"));
        AssertBox(firstBoundary, stamps[1].Attribute("Boundary")!.Value);
        AssertBox(firstClip, stamps[1].Attribute("Clip")!.Value);
        AssertBox(secondBoundary, stamps[2].Attribute("Boundary")!.Value);
        AssertBox(secondClip, stamps[2].Attribute("Clip")!.Value);
        AssertBox(thirdClip, stamps[3].Attribute("Clip")!.Value);
        AssertBox("0 0 40 20", stamps[4].Attribute("Clip")!.Value); // short tail is a complete group
        var documentPath = XDocument.Parse(archive.ReadUtf8Text("OFD.xml"))
            .Descendants().Single(node => node.Name.LocalName == "DocRoot").Value.TrimStart('/');
        var pageIds = XDocument.Parse(archive.ReadUtf8Text(documentPath)).Descendants()
            .Where(node => node.Name.LocalName == "Page" && node.Attribute("ID") is not null)
            .Select(node => node.Attribute("ID")!.Value).ToArray();
        Assert.Equal(new[] { pageIds[1], pageIds[0], pageIds[1], pageIds[2], pageIds[3] },
            stamps.Select(stamp => stamp.Attribute("PageRef")!.Value));
        Assert.Contains(coverage.ProtectedFiles, name => name.EndsWith("/Seal.esl", StringComparison.Ordinal));
        Assert.Contains(xml.Descendants().Where(node => node.Name.LocalName == "Reference"),
            node => node.Attribute("FileRef")!.Value.EndsWith("/Seal.esl", StringComparison.Ordinal));
        signed.Position = 0;
        Assert.True((await new OfdSignatureVerifier().VerifyAsync(signed)).ReferenceIntegrityValid);

        signed.Position = 0;
        using var appended = new MemoryStream();
        await new OfdSignatureService().SignAsync(signed, appended, new PictureProvider(Picture));
        appended.Position = 0;
        var again = await new OfdPackageLoader().LoadAsync(appended);
        Assert.Equal(exactXml, again.GetBytes(path));
        Assert.Equal(archive.GetBytes(path.Replace("Signature.xml", "SignedValue.dat")),
            again.GetBytes(path.Replace("Signature.xml", "SignedValue.dat")));
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)] [InlineData(4)]
    public async Task InvalidAppearance_DoesNotCallProviderOrWriteOutput(int kind)
    {
        using var unsigned = await Source();
        var provider = new PictureProvider(kind == 0 ? null : Picture);
        var options = new OfdSignatureOptions();
        if (kind == 0) options.StampAppearances.Add(new OfdNormalStampPlacement(1, 1, 1, 10, 10));
        if (kind == 1) options.StampAppearances.Add(new OfdNormalStampPlacement(5, 1, 1, 10, 10));
        if (kind == 2) options.StampAppearances.Add(new OfdNormalStampPlacement(1, 1, 1, double.NaN, 10));
        if (kind == 3) options.StampAppearances.Add(new OfdNormalStampPlacement(1, 95, 1, 10, 10));
        if (kind == 4)
        {
            var riding = new OfdRidingStampPlacement(OfdStampEdge.Right, 1, 4, 40, 20)
            { GroupSize = 3 };
            riding.SliceWeights.Add(1); riding.SliceWeights.Add(0); riding.SliceWeights.Add(1);
            options.StampAppearances.Add(riding);
        }
        using var output = new MemoryStream();
        output.WriteByte(0x7b);
        await Assert.ThrowsAnyAsync<Exception>(() =>
            new OfdSignatureService().SignAsync(unsigned, output, provider, options));
        Assert.Equal(new byte[] { 0x7b }, output.ToArray());
        Assert.Equal(0, provider.SignCalls);
    }

    [Theory]
    [InlineData("./", true)]
    [InlineData("Pages/../", true)]
    [InlineData("../../", false)]
    public async Task PageBaseLoc_ResolvesDotSegmentsWithinPackage(string prefix, bool valid)
    {
        using var source = await Source();
        using var changed = new MemoryStream();
        changed.Write(source.ToArray());
        changed.Position = 0;
        using (var zip = new ZipArchive(changed, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("Doc_0/Document.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            var page = xml.Descendants().First(node => node.Name.LocalName == "Page" && node.Attribute("BaseLoc") is not null);
            page.SetAttributeValue("BaseLoc", prefix + page.Attribute("BaseLoc")!.Value);
            entry.Delete();
            using var output = zip.CreateEntry("Doc_0/Document.xml").Open();
            xml.Save(output);
        }
        changed.Position = 0;
        var provider = new PictureProvider(Picture);
        var options = new OfdSignatureOptions();
        options.StampAppearances.Add(new OfdNormalStampPlacement(1, 1, 1, 10, 10));
        using var signed = new MemoryStream();
        if (valid)
        {
            await new OfdSignatureService().SignAsync(changed, signed, provider, options);
            Assert.Equal(1, provider.SignCalls);
            signed.Position = 0;
            Assert.True((await new OfdSignatureVerifier().VerifyAsync(signed)).ReferenceIntegrityValid);
        }
        else
        {
            await Assert.ThrowsAnyAsync<Exception>(() =>
                new OfdSignatureService().SignAsync(changed, signed, provider, options));
            Assert.Equal(0, provider.SignCalls);
            Assert.Empty(signed.ToArray());
        }
    }

    private static async Task<MemoryStream> Source()
    {
        var package = new OfdDocumentPackage();
        for (var index = 0; index < 4; index++) package.Pages.Add(new OfdPage
        { Index = index, WidthMillimeters = 100, HeightMillimeters = 80 });
        var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0;
        return output;
    }

    private static void AssertBox(string expected, string actual)
    {
        static double[] Numbers(string value) => value.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(item => double.Parse(item, CultureInfo.InvariantCulture)).ToArray();
        var left = Numbers(expected); var right = Numbers(actual);
        Assert.Equal(4, right.Length);
        for (var index = 0; index < 4; index++) Assert.Equal(left[index], right[index], 8);
    }

    private sealed class PictureProvider(byte[]? picture) : IOfdSignatureProvider
    {
        public int SignCalls { get; private set; }
        public byte[]? SignedXml { get; private set; }
        public string ProviderName => "R11 fixture";
        public string Company => "Ofdrw.Net Tests";
        public string Version => "1";
        public string SignatureMethod => "urn:ofdrw:r11:test";
        public string SignatureType => "Seal";
        public byte[]? SealData => picture;
        public Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation,
            CancellationToken cancellationToken = default)
        {
            SignCalls++;
            SignedXml = (byte[])signatureXml.Clone();
            return Task.FromResult(OfdDigestAlgorithms.Compute(OfdDigestAlgorithms.Sha256Oid,
                signatureXml.Concat(Encoding.UTF8.GetBytes(propertyInformation)).ToArray()));
        }
    }
}
