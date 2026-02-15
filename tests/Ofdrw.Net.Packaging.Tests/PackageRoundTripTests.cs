using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class PackageRoundTripTests
{
    [Fact]
    public async Task Writer_ShouldGenerateRequiredEntries_AndReaderCanParse()
    {
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = "hello ofd",
                    XMillimeters = 10,
                    YMillimeters = 12,
                    WidthMillimeters = 60,
                    HeightMillimeters = 8,
                    FontSizeMillimeters = 4
                }
            }
        });

        var package = builder.Build();
        var writer = new OfdPackageWriter();

        await using var ms = new MemoryStream();
        await writer.WriteAsync(package, ms);
        ms.Position = 0;

        var loader = new OfdPackageLoader();
        var archive = await loader.LoadAsync(ms);

        Assert.Contains("OFD.xml", archive.EntryNames);
        Assert.Contains("Doc_0/Document.xml", archive.EntryNames);
        Assert.Contains("Doc_0/Pages/Page_0/Content.xml", archive.EntryNames);

        ms.Position = 0;
        var reader = new OfdReader();
        var parsed = await reader.ReadAsync(ms);

        Assert.Single(parsed.Pages);
        var firstText = parsed.Pages[0].Elements.OfType<OfdTextElement>().First();
        Assert.Equal("hello ofd", firstText.Text);
    }
}
