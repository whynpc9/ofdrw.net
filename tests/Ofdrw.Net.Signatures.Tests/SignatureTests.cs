using System.Text;
using System.IO.Compression;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.Tests;

public sealed class SignatureTests
{
    [Fact]
    public void Sm3_ShouldMatchPublishedAbcVector()
    {
        var digest = OfdDigestAlgorithms.Compute(
            OfdDigestAlgorithms.Sm3Oid,
            Encoding.ASCII.GetBytes("abc"));

        Assert.Equal(
            "66c7f0f462eeedd9d1f2d46bdc10e4e2" +
            "4167c4875cf2f7a2297da02b8f4ba8e0",
            Convert.ToHexString(digest).ToLowerInvariant());
    }

    [Fact]
    public async Task Verifier_ShouldValidateUpstreamSm3References()
    {
        var samplePath = Path.Combine(
            ResolveRepositoryRoot(),
            "e2e",
            "Ofdrw.Net.Converter.Pdf.E2E",
            "testdata",
            "upstream-ofdrw",
            "999.ofd");

        await using var input = File.OpenRead(samplePath);
        var report = await new OfdSignatureVerifier().VerifyAsync(input);

        var signature = Assert.Single(report.Signatures);
        Assert.Equal(OfdDigestAlgorithms.Sm3Oid, signature.CheckMethod);
        Assert.Equal(20, signature.References.Count);
        Assert.True(report.ReferenceIntegrityValid);
        Assert.Equal(
            OfdCryptographicVerificationStatus.Unsupported,
            signature.CryptographicStatus);
        Assert.False(report.FullyValid);
    }

    [Fact]
    public async Task Verifier_ShouldRejectTamperedProtectedEntry()
    {
        var samplePath = Path.Combine(
            ResolveRepositoryRoot(),
            "e2e",
            "Ofdrw.Net.Converter.Pdf.E2E",
            "testdata",
            "upstream-ofdrw",
            "999.ofd");

        await using var tampered = await CreateTamperedCopyAsync(samplePath);
        var report = await new OfdSignatureVerifier().VerifyAsync(tampered);

        Assert.False(report.ReferenceIntegrityValid);
        Assert.Contains(
            report.Signatures.SelectMany(signature => signature.References),
            reference => reference.EntryExists && !reference.DigestMatches);
    }

    [Fact]
    public async Task SignerAndRegisteredVerifier_ShouldProduceFullyValidOfd()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = "signed",
                    FontName = "Arial",
                    FontSizeMillimeters = 4,
                    XMillimeters = 10,
                    YMillimeters = 10,
                    WidthMillimeters = 40,
                    HeightMillimeters = 8
                }
            }
        });

        await using var unsigned = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, unsigned);
        unsigned.Position = 0;
        await using var signed = new MemoryStream();
        var provider = new TestSignatureProvider();
        await new OfdSignatureService().SignAsync(
            unsigned,
            signed,
            provider,
            new OfdSignatureOptions
            {
                CheckMethod = OfdDigestAlgorithms.Sha256Oid
            });

        signed.Position = 0;
        var report = await new OfdSignatureVerifier(
                new[] { new TestSignedValueVerifier() })
            .VerifyAsync(signed);

        Assert.True(report.ReferenceIntegrityValid);
        Assert.True(report.FullyValid);
        Assert.Equal(
            OfdCryptographicVerificationStatus.Valid,
            Assert.Single(report.Signatures).CryptographicStatus);
    }

    private static byte[] CreateSignedValue(
        byte[] signatureXml,
        string propertyInformation)
    {
        var propertyBytes = Encoding.UTF8.GetBytes(propertyInformation);
        var input = new byte[signatureXml.Length + propertyBytes.Length];
        Buffer.BlockCopy(signatureXml, 0, input, 0, signatureXml.Length);
        Buffer.BlockCopy(
            propertyBytes,
            0,
            input,
            signatureXml.Length,
            propertyBytes.Length);
        return OfdDigestAlgorithms.Compute(
            OfdDigestAlgorithms.Sha256Oid,
            input);
    }

    private static string ResolveRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate repository root.");
    }

    private static async Task<MemoryStream> CreateTamperedCopyAsync(
        string samplePath)
    {
        var output = new MemoryStream();
        using (var target = new ZipArchive(
                   output,
                   ZipArchiveMode.Create,
                   leaveOpen: true))
        await using (var sourceStream = File.OpenRead(samplePath))
        using (var source = new ZipArchive(
                   sourceStream,
                   ZipArchiveMode.Read,
                   leaveOpen: false))
        {
            var modified = false;
            foreach (var entry in source.Entries)
            {
                var copy = target.CreateEntry(
                    entry.FullName,
                    CompressionLevel.Optimal);
                await using var copyStream = copy.Open();
                await using var entryStream = entry.Open();
                await entryStream.CopyToAsync(copyStream);
                if (!modified &&
                    entry.FullName.EndsWith(
                        "/Content.xml",
                        StringComparison.OrdinalIgnoreCase))
                {
                    await copyStream.WriteAsync(new byte[] { 0x20 });
                    modified = true;
                }
            }

            Assert.True(modified);
        }

        output.Position = 0;
        return output;
    }

    private sealed class TestSignatureProvider : IOfdSignatureProvider
    {
        public string ProviderName => "Ofdrw.Net.Tests";
        public string Company => "Ofdrw.Net";
        public string Version => "1";
        public string SignatureMethod => "urn:ofdrw-net:test-signature";
        public string SignatureType => "Sign";
        public byte[]? SealData => null;

        public Task<byte[]> SignAsync(
            byte[] signatureXml,
            string propertyInformation,
            CancellationToken cancellationToken = default)
        {
            return Task.FromResult(
                CreateSignedValue(signatureXml, propertyInformation));
        }
    }

    private sealed class TestSignedValueVerifier : IOfdSignedValueVerifier
    {
        public string SignatureMethod => "urn:ofdrw-net:test-signature";

        public Task<bool> VerifyAsync(
            byte[] signatureXml,
            byte[] signedValue,
            string propertyInformation,
            CancellationToken cancellationToken = default)
        {
            return Task.FromResult(
                CreateSignedValue(signatureXml, propertyInformation)
                    .SequenceEqual(signedValue));
        }
    }
}
