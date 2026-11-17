using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using Ofdrw.Net.Packaging.Archive;
using Xunit;

namespace Ofdrw.Net.Packaging.Tests;

/// <summary>Limit failures are classified by an internal marker, not by message text.</summary>
public class PackageLoaderLimitMarkerTests
{
    private static byte[] Zip(int entries)
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
            for (var i = 0; i < entries; i++)
            {
                using var w = new StreamWriter(zip.CreateEntry("e" + i + ".xml").Open());
                w.Write("<a/>");
            }
        return ms.ToArray();
    }

    [Fact]
    public async Task ConfiguredLimitsAreMarked()
    {
        var options = new OfdPackageLoadOptions { MaxEntryCount = 1 };
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdPackageLoader().LoadAsync(new MemoryStream(Zip(2)), options));
        Assert.True(OfdPackageLoader.IsLimitFailure(error));

        var bytes = Zip(1);
        error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdPackageLoader().LoadAsync(new MemoryStream(bytes), new OfdPackageLoadOptions { MaxInputBytes = bytes.Length - 1 }));
        Assert.True(OfdPackageLoader.IsLimitFailure(error));
    }

    [Fact]
    public void MalformedInputIsNotMarkedEvenWhenTheTextMentionsLimit()
    {
        Assert.False(OfdPackageLoader.IsLimitFailure(new InvalidDataException("unexpected limit marker in malformed ZIP")));
    }
}
