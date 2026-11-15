using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class TerminalTextTests : IDisposable
{
    private const string Ns = "http://www.ofdspec.org/2016";
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-terminal-" + Guid.NewGuid().ToString("N"));

    public TerminalTextTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public void PlainAndCjkTextIsUnchanged()
    {
        const string text = "abc 中文 é 🙂 \\u001B";
        Assert.Same(text, TerminalText.Escape(text));
        Assert.Equal(string.Empty, TerminalText.Escape(null));
        Assert.Equal(string.Empty, TerminalText.Escape(string.Empty));
    }

    [Theory]
    [InlineData(0x09)]
    [InlineData(0x0A)]
    [InlineData(0x0D)]
    [InlineData(0x1B)]
    [InlineData(0x7F)]
    [InlineData(0x9B)]
    [InlineData(0x200F)]
    [InlineData(0x202E)]
    [InlineData(0x2066)]
    [InlineData(0x2028)]
    public void ControlAndBidiCharactersBecomeVisible(int code)
    {
        var actual = TerminalText.Escape("a" + (char)code + "b");

        Assert.Equal("a\\u" + code.ToString("X4", CultureInfo.InvariantCulture) + "b", actual);
        Assert.DoesNotContain(actual, char.IsControl);
    }

    [Fact]
    public async Task ErrorMessagesEscapeEntryNames()
    {
        var path = Path.Combine(_directory, "unsafe.ofd");
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            zip.CreateEntry("OFD.xml");
            zip.CreateEntry("../a\u001Bb\rc");
        }

        var (code, stdout, stderr) = await Run("verify-signatures", path);

        Assert.Equal(1, code);
        Assert.Empty(stdout);
        Assert.Equal(@"Error: OFD package entry has an unsafe path: ../a\u001Bb\u000Dc" + Environment.NewLine, stderr);
    }

    [Fact]
    public async Task AttachmentRowsKeepFiveColumns()
    {
        var (code, stdout, stderr) = await Run("attachments", "list", await Seeded());

        Assert.Equal(0, code);
        Assert.Empty(stderr);
        var line = Assert.Single(stdout.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries));
        Assert.Equal(
            new[] { @"7\u2066x", @"a\u0009b\u202Ec", @"application/octet-stream\u000Alabel", "5", "internal" },
            line.Split('\t'));
    }

    [Fact]
    public async Task JsonOutputIsUnchanged()
    {
        var (code, stdout, stderr) = await Run("attachments", "list", await Seeded(), "--json");

        Assert.Equal(0, code);
        Assert.Empty(stderr);
        using var json = JsonDocument.Parse(stdout);
        var attachment = Assert.Single(json.RootElement.GetProperty("attachments").EnumerateArray());
        Assert.Equal("7\u2066x", attachment.GetProperty("id").GetString());
        Assert.Equal("a\tb\u202Ec", attachment.GetProperty("name").GetString());
        Assert.Equal("application/octet-stream\nlabel", attachment.GetProperty("mediaType").GetString());
    }

    [Fact]
    public async Task ExtractedTextPayloadIsUnchanged()
    {
        var path = await Seeded();
        var output = Path.Combine(_directory, "extracted.txt");
        const string text = "中文a\tb\u202Ec";

        var (code, stdout, stderr) = await Run("extract-text", path);
        var saved = await Run("extract-text", path, "--output", output);

        Assert.Equal(0, code);
        Assert.Empty(stderr);
        Assert.Equal(text + Environment.NewLine, stdout);
        Assert.Equal(0, saved.Code);
        Assert.Equal(Encoding.UTF8.GetBytes(text), await File.ReadAllBytesAsync(output));
    }

    private async Task<string> Seeded()
    {
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
            ["Doc_0/Document.xml"] = Utf8($"""
                <Document xmlns='{Ns}'>
                  <CommonData><MaxUnitID>100</MaxUnitID><PageArea><PhysicalBox>0 0 210 297</PhysicalBox></PageArea></CommonData>
                  <Pages><Page ID='1' BaseLoc='Pages/Content.xml'/></Pages>
                  <Attachments>Metadata/List.xml</Attachments>
                </Document>
                """),
            ["Doc_0/Pages/Content.xml"] = Utf8($"""
                <Page xmlns='{Ns}'><Content><Layer ID='2'>
                  <TextObject ID='100' Boundary='20 30 100 50' Size='4'>
                    <TextCode X='2' Y='8'>中文a&#9;b&#x202E;c</TextCode>
                  </TextObject>
                </Layer></Content></Page>
                """),
            ["Doc_0/Metadata/List.xml"] = Utf8($"""
                <Attachments xmlns='{Ns}'>
                  <Attachment ID='7&#x2066;x' Name='a&#9;b&#x202E;c' MediaType='application/octet-stream&#10;label'>
                    <FileLoc>../Payload/original.bin</FileLoc>
                  </Attachment>
                </Attachments>
                """),
            ["Doc_0/Payload/original.bin"] = [0, 1, 2, 3, 255]
        };
        var path = Path.Combine(_directory, "seeded.ofd");
        using var zip = MultiDocumentFixture.Zip(entries);
        await File.WriteAllBytesAsync(path, zip.ToArray());
        return path;
    }

    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);

    private static async Task<(int Code, string Stdout, string Stderr)> Run(params string[] args)
    {
        var savedOut = Console.Out;
        var savedError = Console.Error;
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        try
        {
            Console.SetOut(stdout);
            Console.SetError(stderr);
            return (await global::Cli.RunAsync(args), stdout.ToString(), stderr.ToString());
        }
        finally
        {
            Console.SetOut(savedOut);
            Console.SetError(savedError);
        }
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
