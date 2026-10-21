using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.Tests;

public sealed class SignatureScopeTests
{
    private const string Method = "urn:ofdrw-net:r13:test";

    [Fact]
    public async Task OptionalUnreferencedEntryCanBeExcludedWithoutClaimingItsIntegrity()
    {
        using var source = await CreatePackage();
        var provider = new Provider();
        using var signed = new MemoryStream();
        var selection = await new OfdSignatureService().SignWithResultAsync(source, signed, provider,
            new OfdSignatureOptions
            {
                CheckMethod = OfdDigestAlgorithms.Sha256Oid,
                ProtectFile = path => path != "/Extras/optional.bin"
            });

        Assert.True(selection.RequiredFileCoverageValid);
        Assert.Contains("/OFD.xml", selection.ProtectedFiles);
        Assert.Contains("/Doc_0/Document.xml", selection.ProtectedFiles);
        Assert.Contains("/Doc_0/Res/opaque.bin", selection.ProtectedFiles);
        Assert.Contains("/Extras/optional.bin", selection.UnprotectedFiles);
        Assert.Contains("/Doc_0/Signs/Signatures.xml", selection.UnprotectedFiles);
        Assert.Equal(1, provider.Calls);
        var report = await Verify(signed);
        Assert.True(report.ReferenceIntegrityValid);
        Assert.True(report.FullyValid);

        using var optionalChanged = await Rewrite(signed, "Extras/optional.bin", _ => [9, 8, 7]);
        Assert.True((await Verify(optionalChanged)).FullyValid);
        using var requiredChanged = await Rewrite(signed, "Doc_0/Res/opaque.bin", _ => [9, 8, 7]);
        Assert.False((await Verify(requiredChanged)).ReferenceIntegrityValid);
    }

    [Theory]
    [InlineData("/OFD.xml")]
    [InlineData("/Doc_0/Document.xml")]
    [InlineData("/Doc_0/Res/opaque.bin")]
    public async Task RequiredFileCannotBeOmittedBeforeProviderOrOutput(string excluded)
    {
        using var source = await CreatePackage();
        using var target = new MemoryStream();
        var provider = new Provider();
        var options = new OfdSignatureOptions { ProtectFile = path => path != excluded };
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignWithResultAsync(source, target, provider, options));
        Assert.Contains(excluded, error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, target.Length);
    }

    [Fact]
    public async Task ReferencedSharedResourceOutsideDocumentDirectoryCannotBeOmitted()
    {
        using var source = await CreatePackage();
        using var referenced = await Rewrite(source, "Doc_0/Document.xml", bytes =>
        {
            var document = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            document.Root!.Add(new XElement(document.Root.Name.Namespace + "VendorResource",
                "/Shared/external.bin"));
            return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        });
        using var target = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(referenced, target, provider,
                new OfdSignatureOptions
                {
                    ProtectFile = path => path != "/Shared/external.bin"
                }));
        Assert.Contains("/Shared/external.bin", error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, target.Length);
    }

    [Fact]
    public async Task DefaultSelectionAndProtectedListRetainExistingLockBehavior()
    {
        using var source = await CreatePackage();
        using var first = new MemoryStream();
        var service = new OfdSignatureService();
        var defaultSelection = await service.SignWithResultAsync(source, first, new Provider());
        Assert.Contains("/Extras/optional.bin", defaultSelection.ProtectedFiles);
        Assert.Equal(new[]
        {
            "/Doc_0/Signs/Sign_0/Signature.xml",
            "/Doc_0/Signs/Sign_0/SignedValue.dat",
            "/Doc_0/Signs/Signatures.xml"
        }, defaultSelection.UnprotectedFiles);

        source.Position = 0;
        using var locked = new MemoryStream();
        await service.SignAsync(source, locked, new Provider(),
            new OfdSignatureOptions { ProtectSignatureList = true });
        locked.Position = 0;
        using var attempted = new MemoryStream();
        var nextProvider = new Provider();
        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            service.SignAsync(locked, attempted, nextProvider));
        Assert.Equal(0, nextProvider.Calls);
        Assert.Equal(0, attempted.Length);
    }

    [Fact]
    public async Task Ofd2RelativeAndParametersAreSignedAndBaseBytesRemainUnchanged()
    {
        using var source = await CreatePackage(version2: true);
        var service = new OfdSignatureService();
        using var first = new MemoryStream();
        await service.SignAsync(source, first, new Provider());
        first.Position = 0;
        var original = await new OfdPackageLoader().LoadAsync(first);
        var oldXml = original.GetBytes("Doc_0/Signs/Sign_0/Signature.xml");
        var oldValue = original.GetBytes("Doc_0/Signs/Sign_0/SignedValue.dat");

        var options = new OfdSignatureOptions { RelativeSignatureId = "1" };
        options.Parameters.Add(new OfdSignatureParameter("purpose", "second approval", "text"));
        first.Position = 0;
        using var second = new MemoryStream();
        await service.SignAsync(first, second, new Provider(), options);
        second.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(second);
        Assert.Equal(oldXml, archive.GetBytes("Doc_0/Signs/Sign_0/Signature.xml"));
        Assert.Equal(oldValue, archive.GetBytes("Doc_0/Signs/Sign_0/SignedValue.dat"));
        var list = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Signs/Signatures.xml"));
        var ns = list.Root!.Name.Namespace;
        Assert.Equal("1", list.Root.Elements(ns + "Signature").Last().Attribute("Relative")?.Value);
        var current = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Signs/Sign_1/Signature.xml"));
        Assert.Equal(ns, current.Root!.Name.Namespace);
        var signedInfo = current.Root!.Element(ns + "SignedInfo")!;
        Assert.True(signedInfo.Element(ns + "Parameters") is not null, current.ToString());
        var parameter = Assert.Single(signedInfo.Element(ns + "Parameters")!.Elements(ns + "Parameter"));
        Assert.Equal("purpose", parameter.Attribute("Name")?.Value);
        Assert.Equal("text", parameter.Attribute("Type")?.Value);
        Assert.Equal("second approval", parameter.Value);
        Assert.True(await IsFullyValid(second));

        using var badRelation = await Rewrite(second, "Doc_0/Signs/Signatures.xml", bytes =>
            Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace("Relative=\"1\"", "Relative=\"missing\"")));
        var relationReport = await Verify(badRelation);
        Assert.True(relationReport.ReferenceIntegrityValid);
        Assert.False(relationReport.FullyValid);
        Assert.False(relationReport.Signatures[1].RelativeIntegrityValid);

        using var changedParameter = await Rewrite(second, "Doc_0/Signs/Sign_1/Signature.xml", bytes =>
            Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace("second approval", "changed approval")));
        var changedReport = await Verify(changedParameter);
        Assert.True(changedReport.ReferenceIntegrityValid);
        Assert.False(changedReport.FullyValid);
        Assert.Equal(OfdCryptographicVerificationStatus.Invalid, changedReport.Signatures[1].CryptographicStatus);
    }

    [Fact]
    public async Task Ofd2ExtensionsRejectVersionOneAndMissingBaseBeforeProvider()
    {
        using var versionOne = await CreatePackage();
        using var output = new MemoryStream();
        var provider = new Provider();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdSignatureService().SignAsync(
            versionOne, output, provider,
            new OfdSignatureOptions { RelativeSignatureId = "1" }));
        Assert.Equal(0, provider.Calls);

        using var versionTwo = await CreatePackage(version2: true);
        using var output2 = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdSignatureService().SignAsync(
            versionTwo, output2, provider,
            new OfdSignatureOptions { RelativeSignatureId = "missing" }));
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, output2.Length);

        versionTwo.Position = 0;
        using var output3 = new MemoryStream();
        var duplicates = new OfdSignatureOptions();
        duplicates.Parameters.Add(new OfdSignatureParameter("role", "one"));
        duplicates.Parameters.Add(new OfdSignatureParameter("role", "two"));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdSignatureService().SignAsync(
            versionTwo, output3, provider, duplicates));
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, output3.Length);
    }

    private static async Task<MemoryStream> CreatePackage(bool version2 = false)
    {
        var package = new OfdDocumentPackage();
        if (version2) package.Options.Namespace = OfdConstants.StandardNamespace;
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 });
        package.PreservedEntries["Extras/optional.bin"] = [1, 2, 3];
        package.PreservedEntries["Doc_0/Res/opaque.bin"] = [4, 5, 6];
        package.PreservedEntries["Shared/external.bin"] = [7, 8, 9];
        var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0;
        if (!version2) return output;
        using (output)
            return await Rewrite(output, "OFD.xml", bytes =>
                Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace("Version=\"1.0\"", "Version=\"2.0\"")));
    }

    private static async Task<MemoryStream> Rewrite(MemoryStream source, string path, Func<byte[], byte[]> change)
    {
        source.Position = 0;
        var output = new MemoryStream();
        using (var input = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var stream = entry.Open();
                using var buffer = new MemoryStream();
                await stream.CopyToAsync(buffer);
                var bytes = entry.FullName == path ? change(buffer.ToArray()) : buffer.ToArray();
                using var destination = target.CreateEntry(entry.FullName).Open();
                await destination.WriteAsync(bytes);
            }
        }
        output.Position = 0;
        return output;
    }

    private static async Task<OfdSignatureVerificationReport> Verify(MemoryStream stream)
    {
        stream.Position = 0;
        return await new OfdSignatureVerifier(new[] { new Verifier() }).VerifyAsync(stream);
    }

    private static async Task<bool> IsFullyValid(MemoryStream stream) => (await Verify(stream)).FullyValid;

    private sealed class Provider : IOfdSignatureProvider
    {
        public int Calls { get; private set; }
        public string ProviderName => "R13 test";
        public string Company => "Ofdrw.Net";
        public string Version => "1";
        public string SignatureMethod => Method;
        public string SignatureType => "Sign";
        public byte[]? SealData => null;

        public Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation,
            CancellationToken cancellationToken = default)
        {
            Calls++;
            return Task.FromResult(Digest(signatureXml, propertyInformation));
        }
    }

    private sealed class Verifier : IOfdSignedValueVerifier
    {
        public string SignatureMethod => Method;
        public Task<bool> VerifyAsync(byte[] signatureXml, byte[] signedValue,
            string propertyInformation, CancellationToken cancellationToken = default) =>
            Task.FromResult(signedValue.SequenceEqual(Digest(signatureXml, propertyInformation)));
    }

    private static byte[] Digest(byte[] xml, string path) =>
        OfdDigestAlgorithms.Compute(OfdDigestAlgorithms.Sha256Oid,
            xml.Concat(Encoding.UTF8.GetBytes(path)).ToArray());
}
