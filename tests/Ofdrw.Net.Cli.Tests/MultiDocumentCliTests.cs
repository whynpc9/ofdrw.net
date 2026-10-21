using System.Text;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Cli.Tests;

public sealed class MultiDocumentCliTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "ofd-multidoc-cli-" + Guid.NewGuid().ToString("N"));
    public MultiDocumentCliTests() => Directory.CreateDirectory(directory);

    [Fact]
    public async Task DocumentOption_SelectsExportsAndWatermarkWithoutChangingOtherDocument()
    {
        var source = Path.Combine(directory, "source.ofd");
        var output = Path.Combine(directory, "watermarked.ofd");
        var reordered = Path.Combine(directory, "reordered.ofd");
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        using (var zip = MultiDocumentFixture.Zip(entries)) await File.WriteAllBytesAsync(source, zip.ToArray());
        var svg = Path.Combine(directory, "second.svg");
        var html = Path.Combine(directory, "second.html");
        Assert.Equal(0, await global::Cli.RunAsync(["ofd-to-svg", "--input", source, "--output", svg, "--document", "2"]));
        Assert.Contains("SECOND DOC1", await File.ReadAllTextAsync(svg));
        Assert.Equal(0, await global::Cli.RunAsync(["ofd-to-html", "--input", source, "--output", html, "--document", "2"]));
        Assert.Contains("SECOND DOC1", await File.ReadAllTextAsync(html));
        Assert.Equal(0, await global::Cli.RunAsync(["watermark", source, output, "--document", "2", "--pages", "1", "--text", "SELECTED"]));
        using var saved = new MemoryStream(await File.ReadAllBytesAsync(output));
        var written = MultiDocumentFixture.Unzip(saved);
        Assert.All(entries.Where(entry => entry.Key.StartsWith("Doc_0/", StringComparison.Ordinal)), entry => Assert.Equal(entry.Value, written[entry.Key]));
        saved.Position = 0;
        var selected = await new OfdReader().ReadAsync(saved, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Contains("SELECTED", selected.Pages[0].Elements.OfType<Ofdrw.Net.Core.Models.OfdTextElement>().Select(text => text.Text));
        Assert.Equal(0, await global::Cli.RunAsync(["reorder", source, reordered, "--document", "2", "--pages", "2,1"]));
        using var reorderedStream = new MemoryStream(await File.ReadAllBytesAsync(reordered));
        var reorderedEntries = MultiDocumentFixture.Unzip(reorderedStream);
        Assert.All(entries.Where(entry => entry.Key.StartsWith("Doc_0/", StringComparison.Ordinal)), entry => Assert.Equal(entry.Value, reorderedEntries[entry.Key]));
        reorderedStream.Position = 0;
        var reorderedPackage = await new OfdReader().ReadAsync(reorderedStream, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Contains("PAGE2", reorderedPackage.Pages[0].Elements.OfType<Ofdrw.Net.Core.Models.OfdTextElement>().Single().Text);
        Assert.Equal(1, await global::Cli.RunAsync(["ofd-to-svg", "--input", source, "--output", svg, "--document", "3"]));
    }

    public void Dispose() => Directory.Delete(directory, true);
}
