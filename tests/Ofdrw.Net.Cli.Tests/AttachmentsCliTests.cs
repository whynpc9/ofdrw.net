using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class AttachmentsCliTests : IDisposable
{
    private const string Ns = "http://www.ofdspec.org/2016";
    private static readonly byte[] Payload = [0, 1, 2, 3, 255];
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-attachments-" + Guid.NewGuid().ToString("N"));

    public AttachmentsCliTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public async Task List_Human()
    {
        var (code, stdout, _) = await Run("attachments", "list", await Seeded());

        Assert.Equal(0, code);
        var lines = stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        Assert.Contains(lines, line => line.StartsWith("7\toriginal.bin\t", StringComparison.Ordinal) && line.TrimEnd().EndsWith("internal", StringComparison.Ordinal));
        Assert.Contains(lines, line => line.TrimEnd().EndsWith("external", StringComparison.Ordinal));
    }

    [Fact]
    public async Task List_Json()
    {
        var (code, stdout, _) = await Run("attachments", "list", await Seeded(), "--json");

        Assert.Equal(0, code);
        using var json = JsonDocument.Parse(stdout);
        Assert.Equal("ofdrw-cli-attachments/1", json.RootElement.GetProperty("schema").GetString());
        var items = json.RootElement.GetProperty("attachments").EnumerateArray().ToArray();
        Assert.Equal(2, items.Length);
        var internalItem = Assert.Single(items, item => !item.GetProperty("external").GetBoolean());
        Assert.Equal(5, internalItem.GetProperty("bytes").GetInt32());
    }

    [Fact]
    public async Task Extract_WritesPayload()
    {
        var output = PathFor("x.bin");

        Assert.Equal(0, (await Run("attachments", "extract", await Seeded(), "--id", "7", "--output", output)).Code);

        Assert.Equal(Payload, await File.ReadAllBytesAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task Extract_External_Fails()
    {
        var output = PathFor("x.bin");

        Assert.Equal(1, (await Run("attachments", "extract", await Seeded(), "--id", "8", "--output", output)).Code);

        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task Extract_Unknown_Fails() =>
        Assert.Equal(1, (await Run("attachments", "extract", await Seeded(), "--id", "99", "--output", PathFor("x.bin"))).Code);

    [Fact]
    public async Task Remove_Logical()
    {
        var output = PathFor("out.ofd");

        Assert.Equal(0, (await Run("attachments", "remove", await Seeded(), output, "--id", "7")).Code);

        Assert.Single((await Read(output)).Attachments);
        Assert.Contains("Doc_0/Payload/original.bin", EntryNames(output));
    }

    [Fact]
    public async Task Remove_Physical()
    {
        var output = PathFor("out.ofd");

        Assert.Equal(0, (await Run("attachments", "remove", await Seeded(), output, "--id", "7", "--physical")).Code);

        var remaining = Assert.Single((await Read(output)).Attachments);
        Assert.True(remaining.IsExternal);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", EntryNames(output));
    }

    [Fact]
    public async Task Remove_PhysicalExternal_Fails()
    {
        var output = PathFor("out.ofd");
        await File.WriteAllTextAsync(output, "unchanged");

        Assert.Equal(1, (await Run("attachments", "remove", await Seeded(), output, "--id", "8", "--physical")).Code);

        Assert.Equal("unchanged", await File.ReadAllTextAsync(output));
    }

    [Fact]
    public async Task Remove_InPlace()
    {
        var input = await Seeded();

        Assert.Equal(0, (await Run("attachments", "remove", input, input, "--id", "7")).Code);

        Assert.Single((await Read(input)).Attachments);
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task UnknownSubcommand_Fails() =>
        Assert.Equal(1, (await Run("attachments", "frob", await Seeded())).Code);

    [Fact]
    public async Task MultiDocument_ListAndExtractSecondDocument()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var document = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_1/Document.xml"]));
        var ns = document.Root!.Name.Namespace;
        document.Root.Element(ns + "Pages")!.AddAfterSelf(new XElement(ns + "Attachments", "Attachs/List.xml"));
        entries["Doc_1/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        entries["Doc_1/Attachs/List.xml"] = Encoding.UTF8.GetBytes(new XElement(ns + "Attachments",
            new XElement(ns + "Attachment", new XAttribute("ID", "21"), new XAttribute("Name", "second.bin"),
                new XElement(ns + "FileLoc", "second.bin"))).ToString(SaveOptions.DisableFormatting));
        entries["Doc_1/Attachs/second.bin"] = [9, 8, 7];
        var input = PathFor("multi.ofd");
        using (var zip = MultiDocumentFixture.Zip(entries)) await File.WriteAllBytesAsync(input, zip.ToArray());

        var second = await Run("attachments", "list", input, "--document", "2", "--json");
        Assert.Equal(0, second.Code);
        using (var json = JsonDocument.Parse(second.Stdout))
            Assert.Equal("21", Assert.Single(json.RootElement.GetProperty("attachments").EnumerateArray()).GetProperty("id").GetString());
        var first = await Run("attachments", "list", input, "--document", "1", "--json");
        Assert.Equal(0, first.Code);
        using (var json = JsonDocument.Parse(first.Stdout))
            Assert.Empty(json.RootElement.GetProperty("attachments").EnumerateArray());

        var extracted = PathFor("second.bin");
        Assert.Equal(0, (await Run("attachments", "extract", input, "--document", "2", "--id", "21", "--output", extracted)).Code);
        Assert.Equal(new byte[] { 9, 8, 7 }, await File.ReadAllBytesAsync(extracted));

        var output = PathFor("multi-out.ofd");
        Assert.Equal(1, (await Run("attachments", "remove", input, output, "--document", "2", "--id", "21")).Code);
        Assert.False(File.Exists(output));
    }

    private string PathFor(string name) => Path.Combine(_directory, name);

    // Same package as R18AttachmentPhysicalCleanupTests.Seed(external: true).
    private async Task<string> Seeded()
    {
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
            ["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'><CommonData><MaxUnitID>2</MaxUnitID><PageArea><PhysicalBox>0 0 210 297</PhysicalBox></PageArea></CommonData><Pages><Page ID='1' BaseLoc='Pages/Content.xml'/></Pages><Attachments>Metadata/List.xml</Attachments></Document>"),
            ["Doc_0/Pages/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer ID='2'/></Content></Page>"),
            ["Doc_0/Metadata/List.xml"] = Utf8($"<Attachments xmlns='{Ns}'><Attachment ID='7' Name='original.bin'><FileLoc>../Payload/original.bin</FileLoc></Attachment><Attachment ID='8' Name='external.bin' External='true'><FileLoc>https://invalid.example/never-fetch.bin</FileLoc></Attachment></Attachments>"),
            ["Doc_0/Payload/original.bin"] = Payload
        };
        var path = PathFor("seeded.ofd");
        using (var zip = MultiDocumentFixture.Zip(entries)) await File.WriteAllBytesAsync(path, zip.ToArray());
        return path;
    }

    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);

    private static async Task<Ofdrw.Net.Core.Models.OfdDocumentPackage> Read(string path)
    {
        await using var stream = File.OpenRead(path);
        return await new OfdReader().ReadAsync(stream);
    }

    private static string[] EntryNames(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.Select(entry => entry.FullName).ToArray();
    }

    private static async Task<(int Code, string Stdout, string Stderr)> Run(params string[] args)
    {
        var savedOut = Console.Out; var savedError = Console.Error;
        using var stdout = new StringWriter(); using var stderr = new StringWriter();
        try
        {
            Console.SetOut(stdout); Console.SetError(stderr);
            return (await global::Cli.RunAsync(args), stdout.ToString(), stderr.ToString());
        }
        finally { Console.SetOut(savedOut); Console.SetError(savedError); }
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
