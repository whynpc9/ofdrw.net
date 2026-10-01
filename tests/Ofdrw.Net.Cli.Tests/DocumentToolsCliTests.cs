using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using System.IO.Compression;
using System.Xml.Linq;
using System.Text;

namespace Ofdrw.Net.Cli.Tests;

public sealed class DocumentToolsCliTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofd-document-tools-" + Guid.NewGuid().ToString("N"));
    public DocumentToolsCliTests() => Directory.CreateDirectory(_directory);
    private string PathFor(string name) => Path.Combine(_directory, name);
    [Fact]
    public async Task Commands_UseOneBasedPagesAndAtomicInPlaceOutput()
    {
        var input = PathFor("source.ofd"); var output = PathFor("output.ofd");
        var source = new OfdDocumentPackage();
        for (var i = 0; i < 2; i++)
        {
            var page = new OfdPage { Index = i, WidthMillimeters = 100 + i, HeightMillimeters = 150 };
            page.Elements.Add(new OfdTextElement { Text = "PAGE" + (i + 1) }); source.Pages.Add(page);
        }
        await using (var stream = File.Create(input)) await new OfdPackageWriter().WriteAsync(source, stream);
        Assert.Equal(0, await global::Cli.RunAsync(["watermark", input, input, "--pages", "2", "--text", "DRAFT"]));
        await using (var stream = File.OpenRead(input))
        {
            var read = await new OfdReader().ReadAsync(stream);
            Assert.Single(read.Pages[0].Elements); Assert.Equal(2, read.Pages[1].Elements.Count);
        }
        Assert.Equal(0, await global::Cli.RunAsync(["split", input, output, "--pages", "2"]));
        await using (var stream = File.OpenRead(output)) Assert.Equal(101, Assert.Single((await new OfdReader().ReadAsync(stream)).Pages).WidthMillimeters);
        Assert.Equal(0, await global::Cli.RunAsync(["mix", output, input, "2", input, "1"]));
        await using (var stream = File.OpenRead(output))
        {
            var page = Assert.Single((await new OfdReader().ReadAsync(stream)).Pages);
            Assert.Equal(101, page.WidthMillimeters); Assert.Equal(new[] { "PAGE2", "DRAFT", "PAGE1" }, page.Elements.OfType<OfdTextElement>().Select(text => text.Text));
        }
        Assert.Equal(0, await global::Cli.RunAsync(["clean-signatures", output, output]));
        Assert.Equal(0, await global::Cli.RunAsync(["verify-signatures", output]));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("watermark", "--pages", "0")]
    [InlineData("split", "--pages", "0")]
    [InlineData("watermark", "--unknown", "value")]
    public async Task BadCommand_PreservesExistingOutput(string command, string option, string value)
    {
        var input = PathFor("source.ofd"); var output = PathFor("output.ofd");
        await using (var stream = File.Create(input)) await new OfdPackageWriter().WriteAsync(new OfdDocumentPackage(), stream);
        await File.WriteAllTextAsync(output, "unchanged");
        Assert.Equal(1, await global::Cli.RunAsync([command, input, output, option, value]));
        Assert.Equal("unchanged", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task Verify_MissingDeclaredListMustNotReportNoSignatures()
    {
        var input = PathFor("broken.ofd");
        await using (var stream = File.Create(input)) await new OfdPackageWriter().WriteAsync(new OfdDocumentPackage(), stream);
        using (var zip = ZipFile.Open(input, ZipArchiveMode.Update))
        {
            var rootEntry = zip.GetEntry("OFD.xml")!;
            XDocument root; using (var stream = rootEntry.Open()) root = XDocument.Load(stream);
            root.Root!.Elements().Single().Add(new XElement(root.Root.Name.Namespace + "Signatures", "missing.xml")); rootEntry.Delete();
            using var output = zip.CreateEntry("OFD.xml").Open(); var bytes = Encoding.UTF8.GetBytes(root.ToString()); output.Write(bytes);
        }
        Assert.Equal(3, await global::Cli.RunAsync(["verify-signatures", input]));
    }
    [Fact]
    public async Task Clean_RejectsOrdinaryZipWithoutReplacingExistingOutput()
    {
        var input = PathFor("ordinary.zip"); var output = PathFor("existing.ofd");
        using (var zip = ZipFile.Open(input, ZipArchiveMode.Create))
        { using var file = zip.CreateEntry("ordinary.txt").Open(); file.Write(new byte[] { 1 }); }
        await File.WriteAllTextAsync(output, "unchanged");
        Assert.Equal(1, await global::Cli.RunAsync(["clean-signatures", input, output]));
        Assert.Equal("unchanged", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
