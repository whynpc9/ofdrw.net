using System.IO.Compression;
using System.Text;
using System.Xml;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

/// <summary>Small, deterministic hostile archives exercise the loader before any large allocation.</summary>
public sealed class BadPackageCorpusTests
{
    [Theory]
    [InlineData("../escape.bin")]
    [InlineData("folder/../../escape.bin")]
    [InlineData("folder\\..\\escape.bin")]
    [InlineData("/absolute.bin")]
    public async Task RejectsUnsafePaths(string name)
    {
        using var input = Archive((name, new byte[] { 1 }));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(input));
    }

    [Fact]
    public async Task RejectsCaseInsensitiveDuplicate()
    {
        using var input = Archive(("Doc/A.bin", new byte[] { 1 }), ("doc/a.bin", new byte[] { 2 }));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(input));
    }

    [Fact]
    public async Task RejectsEntryCountAboveBudget()
    {
        using var input = Archive(("a", new byte[] { 1 }), ("b", new byte[] { 2 }));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(input,
            new OfdPackageLoadOptions { MaxEntryCount = 1 }));
    }

    [Fact]
    public async Task RejectsAggregateExpansionAboveBudget()
    {
        using var input = Archive(("a", new byte[16]), ("b", new byte[16]));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(input,
            new OfdPackageLoadOptions { MaxEntryUncompressedBytes = 16, MaxTotalUncompressedBytes = 31 }));
    }

    [Fact]
    public async Task RejectsSmallCompressionBombByRatio()
    {
        using var input = Archive(("zeros", new byte[8192]));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(input,
            new OfdPackageLoadOptions { MaxCompressionRatio = 2 }));
    }

    [Fact]
    public async Task AcceptsExactEntryAndAggregateLimits()
    {
        using var input = Archive(("a", new byte[16]), ("b", new byte[16]));
        var result = await new OfdPackageLoader().LoadAsync(input,
            new OfdPackageLoadOptions { MaxEntryCount = 2, MaxEntryUncompressedBytes = 16,
                MaxTotalUncompressedBytes = 32, MaxCompressionRatio = 1000 });
        Assert.Equal(2, result.EntryNames.Count());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task RejectsFixedMalformedZipSeeds(int seed)
    {
        byte[] bytes = seed switch
        {
            0 => [],
            1 => [0x50, 0x4b, 0x03, 0x04],
            _ => [0, 1, 2, 3, 4, 5, 6, 7]
        };
        using var input = new MemoryStream(bytes);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(input));
    }

    [Fact]
    public async Task ReaderRejectsMalformedRootXml()
    {
        using var input = Archive(("OFD.xml", Encoding.UTF8.GetBytes("<ofd:OFD")));
        await Assert.ThrowsAsync<XmlException>(() => new OfdReader().ReadAsync(input));
    }

    [Fact]
    public async Task ReaderRejectsMissingDocumentReference()
    {
        const string root = "<ofd:OFD xmlns:ofd=\"http://www.ofdspec.org\"><ofd:DocBody><ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot></ofd:DocBody></ofd:OFD>";
        using var input = Archive(("OFD.xml", Encoding.UTF8.GetBytes(root)));
        await Assert.ThrowsAsync<KeyNotFoundException>(() => new OfdReader().ReadAsync(input));
    }

    private static MemoryStream Archive(params (string Name, byte[] Data)[] entries)
    {
        var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, data) in entries)
            {
                using var stream = zip.CreateEntry(name, CompressionLevel.Optimal).Open();
                stream.Write(data);
            }
        }
        output.Position = 0;
        return output;
    }
}
