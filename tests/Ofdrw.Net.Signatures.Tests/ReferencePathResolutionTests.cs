using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.Tests;

public sealed class ReferencePathResolutionTests
{
    private const string UnregisteredMethod = "urn:ofdrw:plan032:unregistered";
    private const string SignaturePath = "Signs/Signature.xml";
    private const string SignedValuePath = "Signs/SignedValue.dat";
    private const string ContentPath = "content.txt";

    [Theory]
    [InlineData("Signature.xml", "/content.txt", "SignedValue.dat")]
    [InlineData("./Signature.xml", "/content.txt", "SignedValue.dat")]
    [InlineData("../Signs/Signature.xml", "/content.txt", "SignedValue.dat")]
    [InlineData("/Signs/Signature.xml", "/content.txt", "SignedValue.dat")]
    [InlineData("Signature.xml", "/./content.txt", "SignedValue.dat")]
    [InlineData("Signature.xml", "content.txt", "SignedValue.dat")]
    [InlineData("Signature.xml", "/content.txt", "../Signs/SignedValue.dat")]
    [InlineData("Signature.xml", "/content.txt", "./SignedValue.dat")]
    public async Task EquivalentReferencesVerifyAlike(string baseLoc, string fileRef, string signedValue)
    {
        await using var package = Package(baseLoc, fileRef, signedValue);
        var report = await new OfdSignatureVerifier().VerifyAsync(package);

        var result = Assert.Single(report.Signatures);
        Assert.True(result.ReferenceIntegrityValid);
        Assert.Equal(OfdCryptographicVerificationStatus.Unsupported, result.CryptographicStatus);
        Assert.Equal(SignaturePath, result.SignaturePath);
        Assert.Equal(SignedValuePath, result.SignedValuePath);
        var reference = Assert.Single(result.References);
        Assert.Equal(fileRef, reference.FileReference);
        Assert.Equal(ContentPath, reference.ResolvedEntryName);
        Assert.False(report.FullyValid);
    }

    [Theory]
    [InlineData("../../Signature.xml", "SignedValue.dat", "BaseLoc")]
    [InlineData("Signature.xml", "../../SignedValue.dat", "SignedValue")]
    public async Task SignatureLocationsThatLeaveThePackageAreErrors(
        string baseLoc, string signedValue, string field)
    {
        await using var package = Package(baseLoc, "/content.txt", signedValue);
        var report = await new OfdSignatureVerifier().VerifyAsync(package);

        var result = Assert.Single(report.Signatures);
        Assert.Equal(OfdCryptographicVerificationStatus.Error, result.CryptographicStatus);
        Assert.Contains("package root", result.CryptographicMessage);
        Assert.Contains(field, result.CryptographicMessage);
        Assert.False(report.FullyValid);
    }

    [Theory]
    [InlineData("/Signs/../content.txt")]
    [InlineData("/../content.txt")]
    [InlineData(@"\Signs\..\content.txt")]
    public async Task FileReferencesWithParentSegmentsAreUnsafe(string fileRef)
    {
        await using var package = Package("Signature.xml", fileRef, "SignedValue.dat");
        var report = await new OfdSignatureVerifier().VerifyAsync(package);

        var result = Assert.Single(report.Signatures);
        var reference = Assert.Single(result.References);
        Assert.Equal("The reference path is unsafe.", reference.Error);
        Assert.False(reference.EntryExists);
        Assert.False(result.ReferenceIntegrityValid);
        Assert.False(report.FullyValid);
    }

    [Theory]
    [InlineData("./Signature.xml", "SignedValue.dat", "Signs/./Signature.xml", "BaseLoc")]
    [InlineData("Signature.xml", ".//SignedValue.dat", "Signs/.//SignedValue.dat", "SignedValue")]
    public async Task SignatureLocationsWithConflictingEntriesAreErrors(
        string baseLoc, string signedValue, string extraEntry, string field)
    {
        var plugin = new RecordingVerifier();
        await using var package = Package(baseLoc, "/content.txt", signedValue, plugin.SignatureMethod, extraEntry);
        var report = await new OfdSignatureVerifier(new[] { plugin }).VerifyAsync(package);

        var result = Assert.Single(report.Signatures);
        Assert.Equal(OfdCryptographicVerificationStatus.Error, result.CryptographicStatus);
        Assert.Contains("ambiguous", result.CryptographicMessage);
        Assert.Contains(field, result.CryptographicMessage);
        Assert.Null(plugin.PropertyInformation);
        Assert.False(report.FullyValid);
    }

    [Fact]
    public async Task MissingCanonicalSignatureDoesNotFallBackToLiteralEntry()
    {
        var plugin = new RecordingVerifier();
        await using var package = Package("./Signature.xml", "/content.txt", "SignedValue.dat",
            plugin.SignatureMethod, "Signs/./Signature.xml");
        using (var zip = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            zip.GetEntry(SignaturePath)!.Delete();
        }
        package.Position = 0;

        var report = await new OfdSignatureVerifier(new[] { plugin }).VerifyAsync(package);

        var result = Assert.Single(report.Signatures);
        Assert.Equal(OfdCryptographicVerificationStatus.Missing, result.CryptographicStatus);
        Assert.Equal(SignaturePath, result.SignaturePath);
        Assert.Null(plugin.PropertyInformation);
        Assert.False(report.FullyValid);
    }

    [Fact]
    public async Task FileReferenceWithConflictingEntriesFailsIntegrity()
    {
        var plugin = new RecordingVerifier();
        await using var package = Package("Signature.xml", "/./content.txt", "SignedValue.dat",
            plugin.SignatureMethod, "./content.txt");
        var report = await new OfdSignatureVerifier(new[] { plugin }).VerifyAsync(package);

        var result = Assert.Single(report.Signatures);
        var reference = Assert.Single(result.References);
        Assert.Equal("The reference path is ambiguous.", reference.Error);
        Assert.False(reference.DigestMatches);
        Assert.False(report.ReferenceIntegrityValid);
        Assert.False(report.FullyValid);
    }

    [Fact]
    public async Task PassesCanonicalPropertyAndOriginalBytesToThePlugin()
    {
        var plugin = new RecordingVerifier();
        await using var package = Package(@"\Signs\.\Signature.xml", "content.txt", @"\Signs\SignedValue.dat",
            plugin.SignatureMethod, "Signs/content.txt");
        byte[] signatureXml;
        byte[] signedValue;
        using (var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
        {
            signatureXml = ReadEntry(zip, SignaturePath);
            signedValue = ReadEntry(zip, SignedValuePath);
        }
        package.Position = 0;

        var report = await new OfdSignatureVerifier(new[] { plugin }).VerifyAsync(package);

        Assert.True(report.FullyValid);
        Assert.Equal("/" + SignaturePath, plugin.PropertyInformation);
        Assert.Equal(signatureXml, plugin.SignatureXml);
        Assert.Equal(signedValue, plugin.SignedValue);
        Assert.Equal(ContentPath, Assert.Single(Assert.Single(report.Signatures).References).ResolvedEntryName);
    }

    private static byte[] ReadEntry(ZipArchive zip, string path)
    {
        using var source = zip.GetEntry(path)!.Open();
        using var buffer = new MemoryStream();
        source.CopyTo(buffer);
        return buffer.ToArray();
    }

    private sealed class RecordingVerifier : IOfdSignedValueVerifier
    {
        public string SignatureMethod => "urn:ofdrw:plan032:recording";
        public string? PropertyInformation { get; private set; }
        public byte[]? SignatureXml { get; private set; }
        public byte[]? SignedValue { get; private set; }

        public Task<bool> VerifyAsync(byte[] signatureXml, byte[] signedValue, string propertyInformation,
            CancellationToken cancellationToken = default)
        {
            SignatureXml = signatureXml;
            SignedValue = signedValue;
            PropertyInformation = propertyInformation;
            return Task.FromResult(true);
        }
    }

    private static MemoryStream Package(
        string baseLoc, string fileRef, string signedValue,
        string signatureMethod = UnregisteredMethod, string? extraEntry = null)
    {
        XNamespace ns = "http://www.ofdspec.org/2016";
        const string content = "protected content";
        var checkValue = Convert.ToBase64String(OfdDigestAlgorithms.Compute(
            OfdDigestAlgorithms.Sha256Oid, Encoding.UTF8.GetBytes(content)));
        var signature = new XElement(ns + "Signature",
            new XElement(ns + "SignedInfo",
                new XElement(ns + "SignatureMethod", signatureMethod),
                new XElement(ns + "References",
                    new XAttribute("CheckMethod", OfdDigestAlgorithms.Sha256Oid),
                    new XElement(ns + "Reference", new XAttribute("FileRef", fileRef),
                        new XElement(ns + "CheckValue", checkValue)))),
            new XElement(ns + "SignedValue", signedValue));

        var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            var ofd = new XElement(ns + "OFD",
                new XElement(ns + "DocBody", new XElement(ns + "Signatures", "Signs/Signatures.xml")));
            var signatures = new XElement(ns + "Signatures", new XElement(ns + "Signature",
                new XAttribute("ID", "1"), new XAttribute("BaseLoc", baseLoc)));
            WriteEntry(zip, "OFD.xml", ofd.ToString());
            WriteEntry(zip, ContentPath, content);
            WriteEntry(zip, "Signs/Signatures.xml", signatures.ToString());
            WriteEntry(zip, SignaturePath, "\uFEFF" + signature.ToString().ReplaceLineEndings("\r\n"));
            WriteEntry(zip, SignedValuePath, "test signed value");
            if (extraEntry is not null)
            {
                WriteEntry(zip, extraEntry, "distinct entry");
            }
        }
        output.Position = 0;
        return output;
    }

    private static void WriteEntry(ZipArchive zip, string name, string text)
    {
        using var target = zip.CreateEntry(name).Open();
        target.Write(Encoding.UTF8.GetBytes(text));
    }
}
