using System.IO.Compression;
using System.Reflection;
using Ofdrw.Net.Packaging.Archive;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.Tests;

public sealed class Audit012RegressionTests
{
    [Fact]
    public async Task Audit012_NormalControl()
    {
        await using var input = CreatePackage();
        var report = await new OfdSignatureVerifier(new[] { new TestVerifier() }).VerifyAsync(input);
        Assert.True(report.FullyValid);
    }

    [Fact]
    public async Task Audit012_CancellationPropagation_Plugin()
    {
        using var cancellation = new CancellationTokenSource();
        await using var input = CreatePackage();
        var verifier = new OfdSignatureVerifier(new[] { new TestVerifier(token =>
        {
            cancellation.Cancel();
            token.ThrowIfCancellationRequested();
            return true;
        }) });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => verifier.VerifyAsync(input, cancellation.Token));
    }

    [Fact]
    public async Task Audit012_ReferenceVisitBudget()
    {
        await using var input = CreatePackage(references: 3);
        var report = await Verify(input, visits: 2);
        AssertIncomplete(report, "reference visit");
        Assert.Equal(2, Assert.Single(report.Signatures).References.Count);
        Assert.All(report.Signatures[0].References, reference => Assert.True(reference.DigestMatches));
    }

    [Theory]
    [InlineData(6, false)]
    [InlineData(7, true)]
    public async Task Audit012_DigestByteBudget(int bytes, bool complete)
    {
        await using var input = CreatePackage(references: 2, secondEntry: true);
        var report = await Verify(input, bytes: bytes);
        Assert.Equal(complete, report.FullyValid);
        Assert.Equal(complete, report.VerificationWorkComplete);
        Assert.Equal(complete, report.Signatures[0].References[1].DigestMatches);
        if (!complete) AssertIncomplete(report, "digest byte");
    }

    [Fact]
    public async Task Audit012_DigestByteBudget_CacheHitStillComparesExpectedValue()
    {
        await using var input = CreatePackage(references: 3, tamperLast: true);
        var report = await Verify(input, bytes: 3);
        Assert.True(report.VerificationWorkComplete);
        Assert.True(report.SignatureDeclarationsComplete);
        Assert.Empty(report.Issues);
        Assert.Equal(new[] { true, true, false }, report.Signatures[0].References.Select(reference => reference.DigestMatches));
        Assert.False(report.FullyValid);
    }

    [Fact]
    public async Task Audit012_DigestByteBudget_CacheSharedAcrossSignaturesAndAlgorithmAliases()
    {
        await using var input = CreatePackage(signatures: 2, algorithms: new[] { " SHA-256 ", OfdDigestAlgorithms.Sha256Oid });
        var report = await Verify(input, bytes: 3);
        Assert.True(report.FullyValid);
        Assert.Equal(2, report.Signatures.Count);
    }

    [Fact]
    public async Task Audit012_DigestByteBudget_DifferentAlgorithmsChargedSeparately()
    {
        await using var input = CreatePackage(signatures: 2, algorithms: new[] { "sha256", "sha1" });
        var report = await Verify(input, bytes: 3);
        AssertIncomplete(report, "digest byte");
        Assert.True(report.Signatures[0].ReferenceIntegrityValid);
        Assert.False(report.Signatures[1].ReferenceIntegrityValid);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit012_SignatureRecordBudget(bool separateLists)
    {
        await using var input = CreatePackage(signatures: 3, repeatedRecords: true, separateLists: separateLists);
        var report = await Verify(input, records: 2);
        AssertIncomplete(report, "signature record");
        Assert.Equal(2, report.Signatures.Count);
        Assert.All(report.Signatures, signature => Assert.True(signature.ReferenceIntegrityValid));
    }

    [Fact]
    public async Task Audit012_UnsupportedStillBounded()
    {
        await using var input = CreatePackage(references: 3);
        var report = await new OfdSignatureVerifier(null, Options(visits: 2)).VerifyAsync(input);
        AssertIncomplete(report, "reference visit");
        Assert.Equal(2, report.Signatures[0].References.Count);
    }

    [Fact]
    public async Task Audit012_UnsupportedControl()
    {
        await using var input = CreatePackage();
        var report = await new OfdSignatureVerifier(null, Options()).VerifyAsync(input);
        Assert.True(report.VerificationWorkComplete);
        Assert.True(report.ReferenceIntegrityValid);
        Assert.False(report.FullyValid);
        Assert.Equal(OfdCryptographicVerificationStatus.Unsupported, report.Signatures[0].CryptographicStatus);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit012_ReferenceVisitBudget_InvalidReferencesAlsoCount(bool unsupportedDigest)
    {
        await using var input = CreatePackage(references: 3, missingReference: !unsupportedDigest,
            algorithms: unsupportedDigest ? new[] { "unknown" } : null);
        var report = await Verify(input, visits: 2);
        AssertIncomplete(report, "reference visit");
        Assert.Equal(2, report.Signatures[0].References.Count);
    }

    [Fact]
    public async Task Audit012_ValidBoundary()
    {
        await using var input = CreatePackage(signatures: 2, references: 2);
        var report = await Verify(input, records: 2, visits: 4, bytes: 3);
        Assert.True(report.FullyValid);
        Assert.True(report.VerificationWorkComplete);
        Assert.Empty(report.Issues);
    }

    [Fact]
    public async Task Audit012_CacheDoesNotCrossVerificationOperations()
    {
        var verifier = new OfdSignatureVerifier(new[] { new TestVerifier() }, Options(bytes: 3));
        await using var first = CreatePackage();
        Assert.True((await verifier.VerifyAsync(first)).FullyValid);
        await using var second = CreatePackage(content: "xyz", expectedContent: "abc");
        var report = await verifier.VerifyAsync(second);
        Assert.True(report.VerificationWorkComplete);
        Assert.False(report.ReferenceIntegrityValid);
    }

    [Fact]
    public async Task Audit012_CancellationPropagation_ReferenceLoop()
    {
        // Enter the actual reference loop directly to distinguish cancellation there
        // from cancellation in the package loader. No large input or scheduling race.
        using var cancellation = new CancellationTokenSource();
        await using var input = CreatePackage();
        using var zip = new ZipArchive(input, ZipArchiveMode.Read, true);
        using var xmlStream = zip.GetEntry("Signs/0/Signature.xml")!.Open();
        using var xmlBytes = new MemoryStream();
        await xmlStream.CopyToAsync(xmlBytes);
        var archive = new OfdPackageArchive(new Dictionary<string, byte[]>
        {
            ["content.txt"] = Encoding.UTF8.GetBytes("abc")
        });
        var result = new OfdSignatureVerificationResult();
        var budgetType = typeof(OfdSignatureVerifier).Assembly.GetType(
            "Ofdrw.Net.Signatures.Verification.OfdSignatureVerificationBudget")!;
        var budget = Activator.CreateInstance(budgetType, BindingFlags.Instance | BindingFlags.NonPublic,
            null, new object[] { Options() }, null)!;
        var method = typeof(OfdSignatureVerifier).GetMethod("VerifySignatureAsync", BindingFlags.Instance | BindingFlags.NonPublic)!;
        cancellation.Cancel();
        var task = (Task)method.Invoke(new OfdSignatureVerifier(),
            new object[] { archive, xmlBytes.ToArray(), "Signs/0/Signature.xml", result, budget, cancellation.Token })!;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
        Assert.Empty(result.References);
    }

    [Fact]
    public async Task Audit012_CancellationPropagation_PreCancelled()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await using var input = CreatePackage(references: 3);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new OfdSignatureVerifier().VerifyAsync(input, cancellation.Token));
    }

    [Fact]
    public async Task Audit012_CancellationPropagation_PluginReturnsAfterCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        await using var input = CreatePackage(signatures: 2);
        var verifier = new OfdSignatureVerifier(new[] { new TestVerifier(_ =>
        {
            cancellation.Cancel();
            return true;
        }) });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => verifier.VerifyAsync(input, cancellation.Token));
    }

    [Fact]
    public async Task Audit012_PluginFailureRemainsError()
    {
        await using var input = CreatePackage();
        var verifier = new OfdSignatureVerifier(new[] { new TestVerifier(_ => throw new InvalidDataException("plugin failure")) });
        var report = await verifier.VerifyAsync(input);
        Assert.Equal(OfdCryptographicVerificationStatus.Error, report.Signatures[0].CryptographicStatus);
        Assert.False(report.FullyValid);
    }

    [Theory]
    [InlineData(0, 1, 1)]
    [InlineData(-1, 1, 1)]
    [InlineData(1, 0, 1)]
    [InlineData(1, -1, 1)]
    [InlineData(1, 1, 0)]
    [InlineData(1, 1, -1)]
    public void Audit012_OptionsRequirePositiveLimits(int records, int visits, long bytes) =>
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdSignatureVerifier(null, Options(records, visits, bytes)));

    [Fact]
    public async Task Audit012_OptionsAreCopied()
    {
        var options = Options();
        var verifier = new OfdSignatureVerifier(new[] { new TestVerifier() }, options);
        options.MaxSignatureRecords = 0;
        options.MaxReferenceVisits = 0;
        options.MaxDigestBytes = 0;
        await using var input = CreatePackage();
        Assert.True((await verifier.VerifyAsync(input)).FullyValid);
    }

    private static void AssertIncomplete(OfdSignatureVerificationReport report, string limit)
    {
        Assert.False(report.VerificationWorkComplete);
        Assert.False(report.SignatureDeclarationsComplete);
        Assert.False(report.ReferenceIntegrityValid);
        Assert.False(report.FullyValid);
        Assert.Contains(report.Issues, issue => issue.Contains("incomplete") && issue.Contains(limit));
    }

    private static OfdSignatureVerificationOptions Options(int records = 10, int visits = 10, long bytes = 100) => new()
    {
        MaxSignatureRecords = records, MaxReferenceVisits = visits, MaxDigestBytes = bytes
    };

    private static Task<OfdSignatureVerificationReport> Verify(Stream input, int records = 10, int visits = 10, long bytes = 100) =>
        new OfdSignatureVerifier(new[] { new TestVerifier() }, Options(records, visits, bytes)).VerifyAsync(input);

    private static MemoryStream CreatePackage(int signatures = 1, int references = 1, bool secondEntry = false,
        bool tamperLast = false, string[]? algorithms = null, bool missingReference = false,
        string content = "abc", string? expectedContent = null, bool repeatedRecords = false, bool separateLists = false)
    {
        XNamespace ns = "http://www.ofdspec.org/2016";
        var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
        {
            Write(zip, "OFD.xml", new XElement(ns + "OFD", Enumerable.Range(0, separateLists ? signatures : 1).Select(i =>
                new XElement(ns + "DocBody", new XElement(ns + "Signatures", $"Signs/List{i}.xml")))).ToString());
            Write(zip, "content.txt", content);
            Write(zip, "second.txt", "defg");
            for (var list = 0; list < (separateLists ? signatures : 1); list++)
            {
                Write(zip, $"Signs/List{list}.xml", new XElement(ns + "Signatures",
                    Enumerable.Range(separateLists ? list : 0, separateLists ? 1 : signatures).Select(i =>
                        new XElement(ns + "Signature", new XAttribute("ID", i),
                            new XAttribute("BaseLoc", $"{(repeatedRecords ? 0 : i)}/Signature.xml")))).ToString());
            }
            for (var i = 0; i < signatures; i++)
            {
                var algorithm = algorithms?[i] ?? "sha256";
                var refs = Enumerable.Range(0, references).Select(j =>
                {
                    var other = secondEntry && j == 1;
                    var digest = OfdDigestAlgorithms.Compute(OfdDigestAlgorithms.IsSupported(algorithm) ? algorithm : "sha256",
                        Encoding.UTF8.GetBytes(other ? "defg" : expectedContent ?? content));
                    if (tamperLast && j == references - 1) digest[0] ^= 1;
                    return new XElement(ns + "Reference", new XAttribute("FileRef", missingReference ? "missing.txt" :
                        other ? "/second.txt" : j % 2 == 0 ? "/CONTENT.TXT" : "\\content.txt"),
                        new XElement(ns + "CheckValue", Convert.ToBase64String(digest)));
                });
                Write(zip, $"Signs/{i}/Signature.xml", new XElement(ns + "Signature",
                    new XElement(ns + "SignedInfo", new XElement(ns + "SignatureMethod", TestVerifier.Method),
                        new XElement(ns + "References", new XAttribute("CheckMethod", algorithm), refs)),
                    new XElement(ns + "SignedValue", "SignedValue.dat")).ToString());
                Write(zip, $"Signs/{i}/SignedValue.dat", "test-plugin-value");
            }
        }
        output.Position = 0;
        return output;
    }

    private static void Write(ZipArchive zip, string name, string value)
    {
        using var stream = zip.CreateEntry(name).Open();
        stream.Write(Encoding.UTF8.GetBytes(value));
    }

    private sealed class TestVerifier(Func<CancellationToken, bool>? verify = null) : IOfdSignedValueVerifier
    {
        internal const string Method = "urn:ofdrw:audit012:test";
        public string SignatureMethod => Method;
        public Task<bool> VerifyAsync(byte[] signatureXml, byte[] signedValue, string propertyInformation, CancellationToken cancellationToken = default) =>
            Task.FromResult(verify?.Invoke(cancellationToken) ?? true);
    }
}
