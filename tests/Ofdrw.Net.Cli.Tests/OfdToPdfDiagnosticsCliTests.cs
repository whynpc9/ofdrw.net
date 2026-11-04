using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Signatures.Signing;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class OfdToPdfDiagnosticsCliTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-pdf-diag-" + Guid.NewGuid().ToString("N"));

    public OfdToPdfDiagnosticsCliTests() => Directory.CreateDirectory(directory);

    [Fact]
    public async Task UndrawableSeal_IsReportedOnStderrWithExitZero()
    {
        var input = Path.Combine(directory, "sealed.ofd");
        var output = Path.Combine(directory, "sealed.pdf");
        // A PNG signature passes payload sniffing; the garbage body fails decoding at draw time.
        await File.WriteAllBytesAsync(input, await SignedSeal(new byte[] { 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a }
            .Concat(Enumerable.Repeat((byte)0x42, 64)).ToArray()));

        var saved = Console.Error;
        using var errors = new StringWriter();
        int code;
        try { Console.SetError(errors); code = await global::Cli.RunAsync(["ofd-to-pdf", "-i", input, "-o", output]); }
        finally { Console.SetError(saved); }

        Assert.Equal(0, code);
        Assert.Contains("Warning: PDF_SIGNATURE_APPEARANCE_SKIPPED: ", errors.ToString());
        Assert.True(new FileInfo(output).Length > 100);
    }

    private static async Task<byte[]> SignedSeal(byte[] seal)
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 100, HeightMillimeters = 80 });
        using var unsigned = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, unsigned);
        unsigned.Position = 0;
        var options = new OfdSignatureOptions();
        options.StampAppearances.Add(new OfdNormalStampPlacement(1, 10, 10, 40, 20));
        using var signed = new MemoryStream();
        await new OfdSignatureService().SignAsync(unsigned, signed, new PictureProvider(seal), options);
        return signed.ToArray();
    }

    private sealed class PictureProvider(byte[] picture) : IOfdSignatureProvider
    {
        public string ProviderName => "CLI diagnostics fixture";
        public string Company => "Ofdrw.Net Tests";
        public string Version => "1";
        public string SignatureMethod => "urn:ofdrw:r11:test";
        public string SignatureType => "Seal";
        public byte[]? SealData => picture;
        public Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation,
            CancellationToken cancellationToken = default)
            => Task.FromResult(Encoding.UTF8.GetBytes("cli-fixture-only"));
    }

    public void Dispose() => Directory.Delete(directory, true);
}
