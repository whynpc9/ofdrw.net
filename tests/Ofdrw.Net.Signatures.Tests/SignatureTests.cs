using System.Text;
using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;
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

    [Theory]
    [InlineData("")]
    [InlineData("Doc_0/Signs/")]
    public async Task Audit007_RootSignaturesFirstAndAppend_AndNestedControl(string prefix)
    {
        using var unsigned = await CreateAudit007Package(prefix, occupied: false);
        using var first = new MemoryStream();
        var service = new OfdSignatureService();
        var options = new OfdSignatureOptions { CheckMethod = OfdDigestAlgorithms.Sha256Oid };
        await service.SignAsync(unsigned, first, new TestSignatureProvider(), options);
        first.Position = 0;
        var firstArchive = await new OfdPackageLoader().LoadAsync(first);
        Assert.DoesNotContain(firstArchive.EntryNames, name => name.StartsWith('/'));
        var firstSignature = firstArchive.GetBytes(prefix + "Sign_0/Signature.xml");
        var firstValue = firstArchive.GetBytes(prefix + "Sign_0/SignedValue.dat");
        await AssertAudit007Signatures(first, prefix, 1);
        first.Position = 0;
        using var appended = new MemoryStream();
        await service.SignAsync(first, appended, new TestSignatureProvider(), options);
        appended.Position = 0;
        var appendedArchive = await new OfdPackageLoader().LoadAsync(appended);
        Assert.Equal(firstSignature, appendedArchive.GetBytes(prefix + "Sign_0/Signature.xml"));
        Assert.Equal(firstValue, appendedArchive.GetBytes(prefix + "Sign_0/SignedValue.dat"));
        Assert.Contains(prefix + "Sign_1/Signature.xml", appendedArchive.EntryNames);
        await AssertAudit007Signatures(appended, prefix, 2);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Doc_0/Signs/")]
    public async Task Audit007_OccupiedSignDirectory(string prefix)
    {
        using var unsigned = await CreateAudit007Package(prefix, occupied: true);
        using var signed = new MemoryStream();
        await new OfdSignatureService().SignAsync(unsigned, signed, new TestSignatureProvider(),
            new OfdSignatureOptions { CheckMethod = OfdDigestAlgorithms.Sha256Oid });
        signed.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(signed);
        Assert.Equal(new byte[] { 7, 8, 9 }, archive.GetBytes(prefix + "Sign_0/keep.bin"));
        Assert.DoesNotContain(prefix + "Sign_0/Signature.xml", archive.EntryNames);
        Assert.Contains(prefix + "Sign_1/Signature.xml", archive.EntryNames);
        await AssertAudit007Signatures(signed, prefix, 1);
    }

    private static async Task<MemoryStream> CreateAudit007Package(string prefix, bool occupied)
    {
        var package = new OfdDocumentPackage { DocumentEntryPath = "Document.xml" };
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 });
        package.PreservedDocBodyElements.Add(new XElement(XName.Get("Signatures", package.Options.Namespace),
            "/" + prefix + "Signatures.xml").ToString(SaveOptions.DisableFormatting));
        package.PreservedEntries[prefix + "Signatures.xml"] = Encoding.UTF8.GetBytes(
            "<Signatures xmlns=\"http://www.ofdspec.org/2016\"><MaxSignId>0</MaxSignId></Signatures>");
        if (occupied) package.PreservedEntries[prefix + "Sign_0/keep.bin"] = [7, 8, 9];
        var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0;
        return output;
    }

    private static async Task AssertAudit007Signatures(MemoryStream output, string prefix, int count)
    {
        output.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(output);
        Assert.DoesNotContain(archive.EntryNames, name => name.StartsWith('/'));
        var list = XDocument.Parse(archive.ReadUtf8Text(prefix + "Signatures.xml"));
        var ns = list.Root!.Name.Namespace;
        Assert.Equal(count, list.Root.Elements(ns + "Signature").Count());
        foreach (var record in list.Root.Elements(ns + "Signature"))
        {
            var relative = record.Attribute("BaseLoc")!.Value;
            Assert.False(relative.StartsWith('/'));
            var signature = XDocument.Parse(archive.ReadUtf8Text(prefix + relative));
            Assert.StartsWith("/" + prefix, signature.Root!.Element(ns + "SignedValue")!.Value);
            Assert.All(signature.Descendants(ns + "Reference"), reference =>
                Assert.StartsWith("/", reference.Attribute("FileRef")!.Value));
        }
        output.Position = 0;
        var report = await new OfdSignatureVerifier(new[] { new TestSignedValueVerifier() }).VerifyAsync(output);
        Assert.Equal(count, report.Signatures.Count);
        Assert.True(report.ReferenceIntegrityValid);
        Assert.True(report.FullyValid);
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
