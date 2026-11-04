using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class LocateCliTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-locate-" + Guid.NewGuid().ToString("N"));

    public LocateCliTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public async Task Locate_RequiresFind() =>
        Assert.Equal(1, (await Run("locate", await Simple())).Code);

    [Fact]
    public async Task Locate_BadTrustedFont()
    {
        var (code, _, stderr) = await Run("locate", await Simple(), "--find", "world", "--trusted-font", "abc");

        Assert.Equal(1, code);
        Assert.Contains("Trusted font must be SHA256:BYTECOUNT.", stderr);
    }

    [Fact]
    public async Task Locate_Json_Shape()
    {
        var (input, identity) = await RealFont();
        var (code, stdout, _) = await Run("locate", input, "--find", "中文", "--trusted-font", identity, "--json");

        using var json = JsonDocument.Parse(stdout);
        var root = json.RootElement;
        Assert.Equal("ofdrw-cli-locate/1", root.GetProperty("schema").GetString());
        Assert.Equal(ExpectedCode(root), code);
        Assert.Equal(JsonValueKind.Array, root.GetProperty("diagnostics").ValueKind);
    }

    [Fact]
    public async Task Locate_NoMatch()
    {
        var (input, identity) = await RealFont();
        var (code, stdout, _) = await Run("locate", input, "--find", "zzz", "--trusted-font", identity, "--json");

        using var json = JsonDocument.Parse(stdout);
        Assert.Empty(json.RootElement.GetProperty("hits").EnumerateArray());
        Assert.Equal(json.RootElement.GetProperty("complete").GetBoolean() ? 2 : 3, code);
    }

    [Fact]
    public async Task Locate_RealFont_FindsHitWithBounds()
    {
        var (input, identity) = await RealFont();

        var (code, stdout, _) = await Run("locate", input, "--find", "中文", "--trusted-font", identity, "--json");

        Assert.Equal(0, code);
        using var json = JsonDocument.Parse(stdout);
        var hit = Assert.Single(json.RootElement.GetProperty("hits").EnumerateArray());
        Assert.Equal(1, hit.GetProperty("page").GetInt32());
        Assert.Equal("中文", hit.GetProperty("text").GetString());
        Assert.Equal(JsonValueKind.Object, hit.GetProperty("bounds").ValueKind);
        Assert.True(hit.GetProperty("canReplace").GetBoolean());
        var human = await Run("locate", input, "--find", "中文", "--trusted-font", identity);
        Assert.Contains("page 1 \"中文\" box ", human.Stdout);
        Assert.Contains("replaceable=true", human.Stdout);
        Assert.Contains("Complete: true", human.Stdout);
    }

    // The SDK rejects sources it cannot measure (here: an embedded but untrusted face) instead of
    // reporting them as incomplete; the CLI surfaces that as an error.
    [Fact]
    public async Task Locate_UntrustedFont_IsError()
    {
        var (input, _) = await RealFont();
        var (code, _, stderr) = await Run("locate", input, "--find", "中文");

        Assert.Equal(1, code);
        Assert.Contains("unsupported", stderr);
    }

    [Fact]
    public async Task Locate_WritesNothing()
    {
        var input = await Simple();
        var before = Directory.GetFileSystemEntries(_directory).OrderBy(x => x).ToArray();

        await Run("locate", input, "--find", "world");

        Assert.Equal(before, Directory.GetFileSystemEntries(_directory).OrderBy(x => x).ToArray());
    }

    private static int ExpectedCode(JsonElement root) =>
        !root.GetProperty("complete").GetBoolean() ? 3 : root.GetProperty("hits").GetArrayLength() > 0 ? 0 : 2;

    private async Task<string> Simple()
    {
        var path = Path.Combine(_directory, "simple.ofd");
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage
        {
            Index = 0, WidthMillimeters = 210, HeightMillimeters = 297,
            // No literal space: the SDK locator rejects unescaped source whitespace (UnescapedSourceWhitespace).
            Elements = { new OfdTextElement { Text = "helloworld", XMillimeters = 20, YMillimeters = 20, WidthMillimeters = 60, HeightMillimeters = 8 } }
        });
        await using (var output = File.Create(path)) await new OfdPackageWriter().WriteAsync(package, output);
        return path;
    }

    private async Task<(string Path, string Identity)> RealFont()
    {
        // The pinned trusted-font fixture shared with the PDF tests.
        using var archive = System.IO.Compression.ZipFile.OpenRead(Path.Combine(AppContext.BaseDirectory, "fonts", "trusted-fonts.zip"));
        using var entry = archive.GetEntry("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf")!.Open();
        using var copy = new MemoryStream();
        await entry.CopyToAsync(copy);
        var font = copy.ToArray();
        var path = Path.Combine(_directory, "real.ofd");
        await File.WriteAllBytesAsync(path, RealFontOfd(font));
        return (path, $"{Convert.ToHexString(SHA256.HashData(font)).ToLowerInvariant()}:{font.Length}");
    }

    // Raw package with the trusted face embedded, as in KeywordLocatorRealFontTests (CJK advance 4 mm at size 4).
    private static byte[] RealFontOfd(byte[] font)
    {
        const string ns = "http://www.ofdspec.org/2016";
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Encoding.UTF8.GetBytes($"<o:OFD xmlns:o=\"{ns}\" Version=\"1.0\" DocType=\"OFD-H\"><o:DocBody><o:DocInfo/><o:DocRoot>Doc_0/Document.xml</o:DocRoot></o:DocBody></o:OFD>"),
            ["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes($"<o:Document xmlns:o=\"{ns}\"><o:CommonData><o:MaxUnitID>1000</o:MaxUnitID><o:PageArea><o:PhysicalBox>0 0 200 200</o:PhysicalBox></o:PageArea><o:DocumentRes>DocumentRes.xml</o:DocumentRes></o:CommonData><o:Pages><o:Page ID=\"1\" BaseLoc=\"Pages/Page_0/Content.xml\"/></o:Pages></o:Document>"),
            ["Doc_0/DocumentRes.xml"] = Encoding.UTF8.GetBytes($"<o:Res xmlns:o=\"{ns}\"><o:Fonts><o:Font ID=\"10\" FontName=\"OriginalOFL\"><o:FontFile>Res/font.ttf</o:FontFile></o:Font></o:Fonts></o:Res>"),
            ["Doc_0/Pages/Page_0/Content.xml"] = Encoding.UTF8.GetBytes($"<o:Page xmlns:o=\"{ns}\"><o:Area><o:PhysicalBox>0 0 200 200</o:PhysicalBox></o:Area><o:Content><o:Layer ID=\"20\"><o:TextObject ID=\"100\" Boundary=\"20 30 100 50\" Font=\"10\" Size=\"4\"><o:TextCode X=\"2\" Y=\"8\" DeltaX=\"4\">中文</o:TextCode></o:TextObject></o:Layer></o:Content></o:Page>"),
            ["Doc_0/Res/font.ttf"] = font
        };
        using var zip = MultiDocumentFixture.Zip(entries);
        return zip.ToArray();
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
