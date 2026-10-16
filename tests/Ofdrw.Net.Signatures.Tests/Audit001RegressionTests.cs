using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.Tests;

public sealed class Audit001RegressionTests
{
    [Fact]
    public async Task Audit001_ValidAndMissingList()
    {
        await using var package = CreatePackage("missing");
        var report = await CreateVerifier().VerifyAsync(package);

        Assert.True(report.HasSignatureDeclarations);
        Assert.Equal(OfdCryptographicVerificationStatus.Valid, Assert.Single(report.Signatures).CryptographicStatus);
        Assert.Contains(report.Issues, issue => issue.Contains("missing"));
        Assert.False(report.SignatureDeclarationsComplete);
        Assert.False(report.ReferenceIntegrityValid);
        Assert.False(report.FullyValid);
    }

    [Theory]
    [InlineData("malformed")]
    [InlineData("empty")]
    [InlineData("blank")]
    [InlineData("duplicate-missing")]
    public async Task Audit001_MalformedEmptyAndBlankDeclarations(string defect)
    {
        await using var package = CreatePackage(defect);
        var report = await CreateVerifier().VerifyAsync(package);

        Assert.True(Assert.Single(report.Signatures).ReferenceIntegrityValid);
        Assert.False(report.SignatureDeclarationsComplete);
        Assert.False(report.ReferenceIntegrityValid);
        Assert.False(report.FullyValid);
        Assert.NotEmpty(report.Issues);
    }

    [Theory]
    [InlineData("complete", 2)]
    [InlineData("duplicate", 1)]
    [InlineData("single", 1)]
    public async Task Audit001_CompleteAndDuplicateLists(string kind, int expectedSignatures)
    {
        await using var package = CreatePackage(kind);
        var report = await CreateVerifier().VerifyAsync(package);

        Assert.Equal(expectedSignatures, report.Signatures.Count);
        Assert.True(report.SignatureDeclarationsComplete);
        Assert.True(report.ReferenceIntegrityValid);
        Assert.True(report.FullyValid);
        Assert.Empty(report.Issues);
    }

    [Fact]
    public async Task Audit001_UnregisteredVerifierRemainsUnsupported()
    {
        await using var package = CreatePackage("single");
        var report = await new OfdSignatureVerifier().VerifyAsync(package);

        Assert.True(report.ReferenceIntegrityValid);
        Assert.False(report.FullyValid);
        Assert.True(report.SignatureDeclarationsComplete);
        Assert.Equal(OfdCryptographicVerificationStatus.Unsupported, Assert.Single(report.Signatures).CryptographicStatus);
    }

    private static OfdSignatureVerifier CreateVerifier() => new(new[] { new DigestVerifier() });

    private static MemoryStream CreatePackage(string kind)
    {
        XNamespace ns = "http://www.ofdspec.org/2016";
        var declarations = new List<string> { "Signs/First.xml" };
        if (kind == "complete") declarations.Add("Signs/Second.xml");
        else if (kind == "duplicate") declarations.Add("Signs/First.xml");
        else if (kind == "blank") declarations.Add("   ");
        else if (kind != "single") declarations.Add("Signs/Broken.xml");
        if (kind == "duplicate-missing") declarations.Add("Signs/Broken.xml");

        var root = new XDocument(new XElement(ns + "OFD", declarations.Select(location =>
            new XElement(ns + "DocBody", new XElement(ns + "Signatures", location)))));
        var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "OFD.xml", Encoding.UTF8.GetBytes(root.ToString()));
            var content = Encoding.UTF8.GetBytes("Protected content 中英文");
            WriteEntry(archive, "content.txt", content);
            WriteSignature(archive, ns, "First", content);
            if (kind == "complete") WriteSignature(archive, ns, "Second", content);
            if (kind == "malformed") WriteEntry(archive, "Signs/Broken.xml", Encoding.UTF8.GetBytes("<Signatures>"));
            if (kind == "empty") WriteEntry(archive, "Signs/Broken.xml", Encoding.UTF8.GetBytes(new XElement(ns + "Signatures").ToString()));
        }

        output.Position = 0;
        return output;
    }

    private static void WriteSignature(ZipArchive archive, XNamespace ns, string id, byte[] content)
    {
        var list = new XElement(ns + "Signatures", new XElement(ns + "Signature",
            new XAttribute("ID", id), new XAttribute("BaseLoc", id + "/Signature.xml")));
        var signature = new XElement(ns + "Signature",
            new XElement(ns + "SignedInfo",
                new XElement(ns + "SignatureMethod", DigestVerifier.Method),
                new XElement(ns + "References", new XAttribute("CheckMethod", OfdDigestAlgorithms.Sha256Oid),
                    new XElement(ns + "Reference", new XAttribute("FileRef", "/content.txt"),
                        new XElement(ns + "CheckValue", Convert.ToBase64String(OfdDigestAlgorithms.Compute(OfdDigestAlgorithms.Sha256Oid, content)))))),
            new XElement(ns + "SignedValue", "SignedValue.dat"));
        var signatureBytes = Encoding.UTF8.GetBytes(signature.ToString());
        WriteEntry(archive, "Signs/" + id + ".xml", Encoding.UTF8.GetBytes(list.ToString()));
        WriteEntry(archive, "Signs/" + id + "/Signature.xml", signatureBytes);
        WriteEntry(archive, "Signs/" + id + "/SignedValue.dat", OfdDigestAlgorithms.Compute(OfdDigestAlgorithms.Sha256Oid, signatureBytes));
    }

    private static void WriteEntry(ZipArchive archive, string name, byte[] bytes)
    {
        using var stream = archive.CreateEntry(name).Open();
        stream.Write(bytes);
    }

    private sealed class DigestVerifier : IOfdSignedValueVerifier
    {
        internal const string Method = "urn:ofdrw:audit001:test-digest";
        public string SignatureMethod => Method;
        public Task<bool> VerifyAsync(byte[] signatureXml, byte[] signedValue, string propertyInformation, CancellationToken cancellationToken = default) =>
            Task.FromResult(OfdDigestAlgorithms.Compute(OfdDigestAlgorithms.Sha256Oid, signatureXml).SequenceEqual(signedValue));
    }
}
