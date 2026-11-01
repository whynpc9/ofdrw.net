using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TextSizeSerializationTests
{
    [Theory]
    [InlineData(double.Epsilon)]
    [InlineData(0.0001)]
    [InlineData(0.0004)]
    [InlineData(0.0004999999)]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public async Task NewTextWithUnwritableSizeFailsBeforeDestinationIsTouched(double size)
    {
        using var destination = new UntouchableStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdPackageWriter().WriteAsync(Package(size), destination));
        Assert.Contains("writer precision", error.Message);
    }

    [Theory]
    [InlineData("later-page")]
    [InlineData("composite")]
    [InlineData("annotation")]
    public async Task UnwritableSizeInEveryGeneratedContainerRejectsTheWholeWrite(string container)
    {
        var package = Package(4);
        var tiny = (OfdTextElement)Package(0.0004).Pages[0].Elements[0];
        if (container == "later-page")
            package.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 210, HeightMillimeters = 297, Elements = { tiny } });
        else if (container == "composite")
        {
            var resource = new OfdCompositeResource { WidthMillimeters = 100, HeightMillimeters = 50, Elements = { tiny } };
            package.CompositeResources = new(); package.CompositeResources.Items.Add(resource);
            package.Pages[0].Elements.Add(new OfdCompositeElement { Resource = resource, WidthMillimeters = 100, HeightMillimeters = 50 });
        }
        else
        {
            var annotation = new OfdAnnotation { Creator = "N03", Type = OfdAnnotationType.Stamp,
                LastModifiedDate = new DateTime(2026, 11, 1),
                Appearance = new() { WidthMillimeters = 100, HeightMillimeters = 50, Elements = { tiny } } };
            package.Pages[0].Annotations = new(); package.Pages[0].Annotations.Items.Add(annotation);
        }
        using var destination = new UntouchableStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(package, destination));
        Assert.Contains("writer precision", error.Message);
    }

    [Theory]
    [InlineData(0.0004999999999999999, "0.001")]
    [InlineData(0.0005, "0.001")]
    [InlineData(0.0005000000000000001, "0.001")]
    [InlineData(0.001, "0.001")]
    [InlineData(4, "4")]
    [InlineData(4.1234, "4.123")]
    public async Task WritableSizeKeepsExistingPrecisionAndReadsBackPositive(double size, string literal)
    {
        using var destination = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(Package(size), destination);
        destination.Position = 0;
        using (var zip = new ZipArchive(destination, ZipArchiveMode.Read, leaveOpen: true))
        using (var page = zip.Entries.Single(entry => entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal)).Open())
        {
            var text = XDocument.Load(page).Descendants().Single(node => node.Name.LocalName == "TextObject");
            Assert.Equal(literal, text.Attribute("Size")!.Value);
            Assert.Equal("10.123 12.346 60 8", text.Attribute("Boundary")!.Value);
        }
        destination.Position = 0;
        var read = await new OfdReader().ReadAsync(destination);
        var output = Assert.IsType<OfdTextElement>(Assert.Single(Assert.Single(read.Pages).Elements));
        Assert.Equal("中文 Alpha", output.Text);
        Assert.Equal(double.Parse(literal, CultureInfo.InvariantCulture), output.FontSizeMillimeters);
        Assert.True(output.FontSizeMillimeters > 0);
    }

    [Fact]
    public async Task RetainedTextKeepsItsOriginalPositiveSubMillimeterLiteral()
    {
        using var destination = new MemoryStream();
        var package = Package(0.0001);
        ((OfdTextElement)package.Pages[0].Elements[0]).SourceXml =
            "<TextObject xmlns='http://www.ofdspec.org/2016' ID='1' Boundary='10 12 60 8' Font='2' Size='0.0001'><TextCode X='0' Y='1'>中文 Alpha</TextCode></TextObject>";
        await new OfdPackageWriter().WriteAsync(package, destination);
        destination.Position = 0;
        var read = await new OfdReader().ReadAsync(destination);
        Assert.Equal(0.0001, Assert.IsType<OfdTextElement>(Assert.Single(Assert.Single(read.Pages).Elements)).FontSizeMillimeters);
    }

    private static OfdDocumentPackage Package(double size)
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297,
            Elements = { new OfdTextElement { Text = "中文 Alpha", FontSizeMillimeters = size,
                XMillimeters = 10.1234, YMillimeters = 12.3456, WidthMillimeters = 60, HeightMillimeters = 8 } } });
        return package;
    }

    // Even querying stream capabilities/position would cross the output boundary.
    private sealed class UntouchableStream : Stream
    {
        public override bool CanRead => throw new InvalidOperationException("Destination touched.");
        public override bool CanSeek => throw new InvalidOperationException("Destination touched.");
        public override bool CanWrite => throw new InvalidOperationException("Destination touched.");
        public override long Length => throw new InvalidOperationException("Destination touched.");
        public override long Position { get => throw new InvalidOperationException("Destination touched."); set => throw new InvalidOperationException("Destination touched."); }
        public override void Flush() => throw new InvalidOperationException("Destination touched.");
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Destination touched.");
        public override long Seek(long offset, SeekOrigin origin) => throw new InvalidOperationException("Destination touched.");
        public override void SetLength(long value) => throw new InvalidOperationException("Destination touched.");
        public override void Write(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Destination touched.");
    }
}
