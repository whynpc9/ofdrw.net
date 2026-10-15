using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Audit002RegressionTests
{
    // A real 64x64 PNG with red/green/blue/yellow quadrants, suitable for rendering.
    private static readonly byte[] ImageData = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAYElEQVR4nO3PIREAIBAAMEKgPw6anIQgCZoGeDK8eLe7FVi7EaVijVJNQEBAQEBAQEBAQEBAQEBAQEBAQEBAQCAdiHlKvd1LCQgICAgICAgICAgICAgICAgICAgICAikffJJBTzDv2gMAAAAAElFTkSuQmCC");

    [Fact]
    public Task Audit002_AbsoluteRootImage() => AssertRoundTrip("/image.png", "image.png", 1);

    [Fact]
    public Task Audit002_ParentRelativeRootImage() => AssertRoundTrip("../../../image.png", "image.png", 1);

    [Fact]
    public Task Audit002_NestedImageControl() => AssertRoundTrip("../../Res/image.png", "Doc_0/Res/image.png", 1);

    [Fact]
    public Task Audit002_RepeatedResource() => AssertRoundTrip("/image.png", "image.png", 2);

    private static async Task AssertRoundTrip(string reference, string imagePath, int imageCount)
    {
        await using var input = CreatePackage(reference, imagePath, imageCount);
        var original = await new OfdPackageLoader().LoadAsync(input);
        Assert.Equal(ImageData, original.GetBytes(imagePath));
        input.Position = 0;
        var parsed = await new OfdReader().ReadAsync(input);
        AssertImages(parsed, imageCount);

        await using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(parsed, output);
        output.Position = 0;
        var written = await new OfdPackageLoader().LoadAsync(output);
        Assert.Equal(ImageData, written.GetBytes(imagePath));
        output.Position = 0;
        var reread = await new OfdReader().ReadAsync(output);
        AssertImages(reread, imageCount);

        var artifactDirectory = Environment.GetEnvironmentVariable("OFDRW_AUDIT002_ARTIFACT_DIR");
        if (!string.IsNullOrWhiteSpace(artifactDirectory))
        {
            Directory.CreateDirectory(artifactDirectory);
            var name = imageCount == 2 ? "repeated" : reference.StartsWith('/') ? "absolute-root" : imagePath == "image.png" ? "parent-root" : "nested-control";
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, name + "-input.ofd"), input.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, name + "-roundtrip.ofd"), output.ToArray());
        }
    }

    private static void AssertImages(OfdDocumentPackage package, int count)
    {
        var images = Assert.Single(package.Pages).Elements.OfType<OfdImageElement>().ToArray();
        Assert.Equal(count, images.Length);
        Assert.All(images, image =>
        {
            Assert.Equal(ImageData, image.Data);
            Assert.Equal("image.png", image.FileName);
            Assert.Equal("image/png", image.MediaType);
            Assert.Equal("10", image.ResourceId);
        });
    }

    private static MemoryStream CreatePackage(string reference, string imagePath, int imageCount)
    {
        const string ns = "http://www.ofdspec.org/2016";
        var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Xml(string path, string content)
            {
                using var writer = new StreamWriter(zip.CreateEntry(path).Open());
                writer.Write(content);
            }

            Xml("OFD.xml", $"<OFD xmlns=\"{ns}\" Version=\"1.1\" DocType=\"OFD\"><DocBody><DocInfo><DocID>audit002</DocID></DocInfo><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
            Xml("Doc_0/Document.xml", $"<Document xmlns=\"{ns}\"><CommonData><MaxUnitID>30</MaxUnitID><PageArea><PhysicalBox>0 0 140 100</PhysicalBox></PageArea></CommonData><Pages><Page ID=\"1\" BaseLoc=\"Pages/Page_0/Content.xml\"/></Pages></Document>");
            var images = string.Concat(Enumerable.Range(0, imageCount).Select(index =>
                $"<ImageObject ID=\"{20 + index}\" ResourceID=\"10\" Boundary=\"{10 + index * 65} 20 50 50\" CTM=\"50 0 0 50 0 0\"/>"));
            Xml("Doc_0/Pages/Page_0/Content.xml", $"<Page xmlns=\"{ns}\"><Area><PhysicalBox>0 0 140 100</PhysicalBox></Area><Content><Layer ID=\"2\">{images}</Layer></Content></Page>");
            Xml("Doc_0/Pages/Page_0/PageRes.xml", $"<Res xmlns=\"{ns}\"><MultiMedias><MultiMedia ID=\"10\" Type=\"Image\" Format=\"PNG\"><MediaFile>{reference}</MediaFile></MultiMedia></MultiMedias></Res>");
            using var image = zip.CreateEntry(imagePath).Open();
            image.Write(ImageData);
        }
        stream.Position = 0;
        return stream;
    }
}
