using System.Runtime.Versioning;
using System.Text;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class OutputPermissionTests : IDisposable
{
    private const string Ns = "http://www.ofdspec.org/2016";
    private const UnixFileMode OwnerOnly = UnixFileMode.UserRead | UnixFileMode.UserWrite;
    private static readonly byte[] Payload = [0, 1, 2, 3, 255];
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-permissions-" + Guid.NewGuid().ToString("N"));

    public OutputPermissionTests() => Directory.CreateDirectory(_directory);

    [UnixFact]
    [UnsupportedOSPlatform("windows")]
    public async Task InPlaceEdit_KeepsOwnerOnlyMode()
    {
        var input = await Seeded();
        File.SetUnixFileMode(input, OwnerOnly);

        var (code, _, stderr) = await Run("attachments", "remove", input, input, "--id", "7");

        Assert.True(code == 0, stderr);
        Assert.Equal(OwnerOnly, File.GetUnixFileMode(input));
    }

    [UnixFact]
    [UnsupportedOSPlatform("windows")]
    public async Task Replace_KeepsExistingWiderMode()
    {
        const UnixFileMode existingMode = OwnerOnly | UnixFileMode.GroupRead;
        var input = await Seeded();
        var output = Path.Combine(_directory, "existing.bin");
        await File.WriteAllTextAsync(output, "old contents");
        File.SetUnixFileMode(output, existingMode);

        var (code, _, stderr) = await Run("attachments", "extract", input, "--id", "7", "--output", output);

        Assert.True(code == 0, stderr);
        Assert.Equal(existingMode, File.GetUnixFileMode(output));
        Assert.Equal(Payload, await File.ReadAllBytesAsync(output));
    }

    [UnixFact]
    [UnsupportedOSPlatform("windows")]
    public async Task Replace_WithSwappedTemporaryPath_DoesNotChangeUnrelatedMode()
    {
        var unrelated = Path.Combine(_directory, "protected.bin");
        await File.WriteAllBytesAsync(unrelated, Payload);
        File.SetUnixFileMode(unrelated, OwnerOnly);

        var outputDirectory = Path.Combine(_directory, "output");
        Directory.CreateDirectory(outputDirectory);
        var output = Path.Combine(outputDirectory, "existing.bin");
        await File.WriteAllBytesAsync(output, []);
        File.SetUnixFileMode(output, OwnerOnly | UnixFileMode.GroupRead);

        await using var staged = new AtomicOutput(output);
        await staged.Stream.WriteAsync(Payload);
        var temporaryPath = Assert.Single(Directory.GetFiles(outputDirectory, ".ofdrw-*.tmp"));
        File.Delete(temporaryPath);
        File.CreateSymbolicLink(temporaryPath, unrelated);

        await staged.CommitAsync(CancellationToken.None);

        Assert.Equal(OwnerOnly, File.GetUnixFileMode(unrelated));
    }

    [Fact]
    public async Task NewOutput_IsCreatedAndReadable()
    {
        var input = await Seeded();
        var output = Path.Combine(_directory, "new.bin");

        var (code, _, stderr) = await Run("attachments", "extract", input, "--id", "7", "--output", output);

        Assert.True(code == 0, stderr);
        Assert.Equal(Payload, await File.ReadAllBytesAsync(output));
        if (!OperatingSystem.IsWindows())
        {
            var control = Path.Combine(_directory, "default.bin");
            await File.WriteAllBytesAsync(control, Payload);
            Assert.Equal(File.GetUnixFileMode(control), File.GetUnixFileMode(output));
        }
    }

    private async Task<string> Seeded()
    {
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Encoding.UTF8.GetBytes($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
            ["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes($"<Document xmlns='{Ns}'><CommonData><MaxUnitID>2</MaxUnitID><PageArea><PhysicalBox>0 0 210 297</PhysicalBox></PageArea></CommonData><Pages><Page ID='1' BaseLoc='Pages/Content.xml'/></Pages><Attachments>Metadata/List.xml</Attachments></Document>"),
            ["Doc_0/Pages/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{Ns}'><Content><Layer ID='2'/></Content></Page>"),
            ["Doc_0/Metadata/List.xml"] = Encoding.UTF8.GetBytes($"<Attachments xmlns='{Ns}'><Attachment ID='7' Name='original.bin'><FileLoc>../Payload/original.bin</FileLoc></Attachment><Attachment ID='8' Name='external.bin' External='true'><FileLoc>https://invalid.example/never-fetch.bin</FileLoc></Attachment></Attachments>"),
            ["Doc_0/Payload/original.bin"] = Payload
        };
        var path = Path.Combine(_directory, "seeded.ofd");
        using var zip = MultiDocumentFixture.Zip(entries);
        await File.WriteAllBytesAsync(path, zip.ToArray());
        return path;
    }

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

    private sealed class UnixFactAttribute : FactAttribute
    {
        public UnixFactAttribute()
        {
            if (OperatingSystem.IsWindows())
                Skip = "Unix permission bits are not available on Windows.";
        }
    }
}
