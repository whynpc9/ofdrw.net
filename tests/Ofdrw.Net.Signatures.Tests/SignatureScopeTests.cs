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
    public async Task DeclaredAnnotationFilesCanBeExcludedWhilePageContentStaysProtected()
    {
        using var source = await CreatePackage();
        using var signed = new MemoryStream();
        var selection = await new OfdSignatureService().SignWithResultAsync(source, signed,
            new Provider(), new OfdSignatureOptions
            {
                ProtectFile = path => !path.StartsWith("/Doc_0/Annots/", StringComparison.Ordinal)
            });
        Assert.Contains("/Doc_0/Annots/Annotations.xml", selection.UnprotectedFiles);
        Assert.Contains("/Doc_0/Annots/Page_0.xml", selection.UnprotectedFiles);
        Assert.Contains("/Doc_0/Document.xml", selection.ProtectedFiles);
        Assert.True(await IsFullyValid(signed));
        using var annotationChanged = await Rewrite(signed, "Doc_0/Annots/Page_0.xml",
            _ => Encoding.UTF8.GetBytes("<PageAnnot changed=\"true\"/>"));
        Assert.True(await IsFullyValid(annotationChanged));
    }

    [Fact]
    public async Task PageUnderAnnotationDirectoryRemainsRequired()
    {
        using var source = await CreatePackage();
        using var moved = await AddReferencedPage(source, "Doc_0/Annots/VisibleContent.xml");
        using var output = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(moved, output, provider,
                new OfdSignatureOptions
                {
                    ProtectFile = path => path != "/Doc_0/Annots/VisibleContent.xml"
                }));
        Assert.Contains("/Doc_0/Annots/VisibleContent.xml", error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("PageRes.xml")]
    [InlineData("Res/logo.png")]
    [InlineData("Res/typeface.ttf")]
    public async Task PageResourcesUnderSignatureNamedDirectoryRemainRequired(string resource)
    {
        using var source = await CreatePackage();
        var additions = new Dictionary<string, byte[]>
        {
            ["Doc_0/Signs/Visible/PageRes.xml"] = Encoding.UTF8.GetBytes(
                "<Res BaseLoc=\"Res\"><MultiMedia ID=\"2\" MediaFile=\"logo.png\"/>" +
                "<Font ID=\"3\" FontName=\"Fixture\"><FontFile>typeface.ttf</FontFile></Font></Res>"),
            ["Doc_0/Signs/Visible/Res/logo.png"] = [1, 2, 3],
            ["Doc_0/Signs/Visible/Res/typeface.ttf"] = [4, 5, 6]
        };
        using var moved = await AddReferencedPage(source, "Doc_0/Signs/Visible/Content.xml", additions);
        var excluded = "/Doc_0/Signs/Visible/" + resource;
        using var output = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(moved, output, provider,
                new OfdSignatureOptions { ProtectFile = path => path != excluded }));
        Assert.Contains(excluded, error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, output.Length);

        moved.Position = 0;
        using var signed = new MemoryStream();
        var selection = await new OfdSignatureService().SignWithResultAsync(moved, signed,
            new Provider(), new OfdSignatureOptions
            {
                ProtectFile = path => path != "/Extras/optional.bin"
            });
        Assert.Contains(excluded, selection.ProtectedFiles);
        using var changed = await Rewrite(signed, excluded.TrimStart('/'), _ => [9, 9, 9]);
        Assert.False((await Verify(changed)).ReferenceIntegrityValid);
    }

    [Fact]
    public async Task AnnotationAndPageResourceSharingAnAnnotsPayloadKeepsItProtected()
    {
        using var source = await CreatePackage();
        var ns = OfdConstants.Namespace;
        using var shared = await AddEntries(source, new Dictionary<string, byte[]>
        {
            ["Doc_0/Pages/Page_0/PageRes.xml"] = Encoding.UTF8.GetBytes(
                "<Res xmlns=\"" + ns + "\" BaseLoc=\"../../Annots/Res\">" +
                "<MultiMedia ID=\"2\" MediaFile=\"shared.png\"/></Res>"),
            ["Doc_0/Annots/Res/shared.png"] = [1, 2, 3],
            ["Doc_0/Annots/Page_0.xml"] = Encoding.UTF8.GetBytes(
                "<PageAnnot xmlns=\"" + ns + "\"><Annot ID=\"4\"><Appearance>" +
                "<ImageObject ResourceID=\"2\"/></Appearance></Annot></PageAnnot>")
        });
        const string excluded = "/Doc_0/Annots/Res/shared.png";
        using var output = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(shared, output, provider,
                new OfdSignatureOptions { ProtectFile = path => path != excluded }));
        Assert.Contains(excluded, error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, output.Length);

        shared.Position = 0;
        using var signed = new MemoryStream();
        var scope = await new OfdSignatureService().SignWithResultAsync(shared, signed,
            new Provider(), new OfdSignatureOptions
            {
                ProtectFile = path => path != "/Extras/optional.bin"
            });
        Assert.Contains(excluded, scope.ProtectedFiles);
        using var changed = await Rewrite(signed, excluded.TrimStart('/'), _ => [9, 9, 9]);
        Assert.False((await Verify(changed)).ReferenceIntegrityValid);
    }

    [Fact]
    public async Task ReferencedUtf16PageWithoutXmlSuffixIsScannedForSharedDependencies()
    {
        using var source = await CreatePackage();
        source.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(source);
        var page = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Pages/Page_0/Content.xml"));
        page.Root!.Add(new XElement(page.Root.Name.Namespace + "VendorResource",
            "/Shared/page-external.bin"));
        var pageBytes = Encoding.BigEndianUnicode.GetBytes(
            "<?xml version=\"1.0\" encoding=\"utf-16BE\"?>" +
            page.ToString(SaveOptions.DisableFormatting));
        using var moved = await AddReferencedPage(source, "Doc_0/Pages/Page_0/Content.dat",
            new Dictionary<string, byte[]> { ["Shared/page-external.bin"] = [1, 2, 3] }, pageBytes);
        using var output = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(moved, output, provider,
                new OfdSignatureOptions
                {
                    ProtectFile = path => path != "/Shared/page-external.bin"
                }));
        Assert.Contains("/Shared/page-external.bin", error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task DeclaredUtf32AttachmentIndexWithoutXmlSuffixProtectsExternalPayload()
    {
        using var source = await CreatePackage();
        using var withIndex = await AddEntries(source, new Dictionary<string, byte[]>
        {
            ["Doc_0/Attachs/Index.dat"] = new UTF32Encoding(true, false).GetBytes(
                "<?xml version=\"1.0\" encoding=\"utf-32BE\"?>" +
                "<Attachments xmlns=\"" + OfdConstants.Namespace + "\">" +
                "<Attachment ID=\"A1\" Name=\"shared\"><FileLoc>/Shared/attachment.bin</FileLoc>" +
                "</Attachment></Attachments>"),
            ["Shared/attachment.bin"] = [1, 2, 3]
        });
        using var referenced = await Rewrite(withIndex, "Doc_0/Document.xml", bytes =>
        {
            var document = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            document.Root!.Add(new XElement(document.Root.Name.Namespace + "Attachments",
                "Attachs/Index.dat"));
            return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        });
        using var output = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(referenced, output, provider,
                new OfdSignatureOptions { ProtectFile = path => path != "/Shared/attachment.bin" }));
        Assert.Contains("/Shared/attachment.bin", error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, output.Length);

        referenced.Position = 0;
        using var signed = new MemoryStream();
        var scope = await new OfdSignatureService().SignWithResultAsync(referenced, signed,
            new Provider(), new OfdSignatureOptions
            {
                ProtectFile = path => path != "/Extras/optional.bin"
            });
        Assert.Contains("/Doc_0/Attachs/Index.dat", scope.ProtectedFiles);
        Assert.Contains("/Shared/attachment.bin", scope.ProtectedFiles);
        using var changed = await Rewrite(signed, "Shared/attachment.bin", _ => [9, 9, 9]);
        Assert.False((await Verify(changed)).ReferenceIntegrityValid);

        using var malformed = await Rewrite(referenced, "Doc_0/Attachs/Index.dat", _ => [1, 2, 3]);
        using var rejected = new MemoryStream();
        var malformedProvider = new Provider();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdSignatureService().SignAsync(
            malformed, rejected, malformedProvider, new OfdSignatureOptions { ProtectFile = _ => true }));
        Assert.Equal(0, malformedProvider.Calls);
        Assert.Equal(0, rejected.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RootXmlAttributeReferencesRemainRequiredEvenForLongPaths(bool longPath)
    {
        var path = "Shared/" + (longPath ? new string('x', 1100) : "payload") + ".bin";
        using var source = await CreatePackage(rootReferencedPath: path);
        using var target = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(source, target, provider,
                new OfdSignatureOptions { ProtectFile = candidate => candidate != "/" + path }));
        Assert.Contains("/" + path, error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, target.Length);
    }

    [Theory]
    [InlineData("PublicRes/PublicRes.xml", "Res", "PublicRes/Res", "logo.png")]
    [InlineData("PublicRes/PublicRes.xml", "Res", "PublicRes/Res", "typeface.ttf")]
    [InlineData("Doc_0/PublicRes.xml", "../Shared", "Shared", "logo.png")]
    [InlineData("Doc_0/PublicRes.xml", "../Shared", "Shared", "typeface.ttf")]
    [InlineData("Shared/PublicRes.xml", "Res", "Shared/Res", "logo.png")]
    [InlineData("Shared/PublicRes.xml", "Res", "Shared/Res", "typeface.ttf")]
    [InlineData("Shared/PublicRes.dat", "Res", "Shared/Res", "logo.png")]
    [InlineData("Shared/PublicRes.dat", "Res", "Shared/Res", "typeface.ttf")]
    public async Task ComposedSharedMediaAndFontLocationsCannotBeOmitted(
        string resourceXmlPath, string baseLocation, string payloadDirectory, string fileName)
    {
        var excluded = "/" + payloadDirectory + "/" + fileName;
        using var source = await CreatePackage(resourceXmlPath: resourceXmlPath,
            resourceBaseLoc: baseLocation, resourcePayloadDirectory: payloadDirectory);
        using var target = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(source, target, provider,
                new OfdSignatureOptions { ProtectFile = path => path != excluded }));
        Assert.Contains(excluded, error.Message);
        Assert.Equal(0, provider.Calls);
        Assert.Equal(0, target.Length);

        source.Position = 0;
        using var signed = new MemoryStream();
        var selection = await new OfdSignatureService().SignWithResultAsync(source, signed,
            new Provider(), new OfdSignatureOptions
            {
                ProtectFile = path => path != "/Extras/optional.bin"
            });
        Assert.Contains(excluded, selection.ProtectedFiles);
        using var modified = await Rewrite(signed, excluded.TrimStart('/'), _ => [9, 9, 9]);
        Assert.False((await Verify(modified)).ReferenceIntegrityValid);
    }

    [Fact]
    public async Task ReferencedResourceTableIsParsedWithoutXmlSuffixOrUtf16Bom()
    {
        using var source = await CreatePackage(resourceXmlPath: "Shared/PublicRes.dat",
            resourceBaseLoc: "Res", resourcePayloadDirectory: "Shared/Res",
            resourceUtf16Be: true);
        using var target = new MemoryStream();
        var provider = new Provider();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdSignatureService().SignAsync(source, target, provider,
                new OfdSignatureOptions
                {
                    ProtectFile = path => path != "/Shared/Res/typeface.ttf"
                }));
        Assert.Contains("/Shared/Res/typeface.ttf", error.Message);
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

    private static async Task<MemoryStream> CreatePackage(bool version2 = false,
        string? rootReferencedPath = null, string? resourceXmlPath = null,
        string? resourceBaseLoc = null, string? resourcePayloadDirectory = null,
        bool resourceUtf16Be = false)
    {
        var package = new OfdDocumentPackage();
        if (version2) package.Options.Namespace = OfdConstants.StandardNamespace;
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 });
        package.PreservedEntries["Extras/optional.bin"] = [1, 2, 3];
        package.PreservedEntries["Doc_0/Res/opaque.bin"] = [4, 5, 6];
        package.PreservedEntries["Shared/external.bin"] = [7, 8, 9];
        package.PreservedEntries["Doc_0/Annots/Annotations.xml"] =
            Encoding.UTF8.GetBytes("<Annotations xmlns=\"" + package.Options.Namespace +
                "\"><Page PageID=\"1\"><FileLoc>Page_0.xml</FileLoc></Page></Annotations>");
        package.PreservedEntries["Doc_0/Annots/Page_0.xml"] =
            Encoding.UTF8.GetBytes("<PageAnnot xmlns=\"" + package.Options.Namespace + "\"/>");
        if (rootReferencedPath is not null)
        {
            package.PreservedEntries[rootReferencedPath] = [3, 2, 1];
            package.PreservedEntries["Shared/manifest.xml"] = Encoding.UTF8.GetBytes(
                "<Root FileLoc=\"/" + rootReferencedPath + "\"/>");
        }
        var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0;
        var withAnnotations = await Rewrite(output, "Doc_0/Document.xml", bytes =>
        {
            var document = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            document.Root!.Add(new XElement(document.Root.Name.Namespace + "Annotations",
                "Annots/Annotations.xml"));
            if (resourceXmlPath is not null)
                document.Root.Add(new XElement(document.Root.Name.Namespace + "PublicRes",
                    "/" + resourceXmlPath));
            return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        });
        output.Dispose();
        var withResources = withAnnotations;
        if (resourceXmlPath is not null)
        {
            var resourceXml = "<Res xmlns=\"http://www.ofdspec.org/2016\" BaseLoc=\"" +
                resourceBaseLoc + "\">" +
                "<MultiMedias><MultiMedia ID=\"2\" MediaFile=\"logo.png\"/></MultiMedias>" +
                "<Fonts><Font ID=\"3\" FontName=\"Fixture\"><FontFile>typeface.ttf</FontFile></Font></Fonts></Res>";
            withResources = await AddEntries(withAnnotations, new Dictionary<string, byte[]>
            {
                [resourceXmlPath] = resourceUtf16Be
                    ? Encoding.BigEndianUnicode.GetBytes("<?xml version=\"1.0\" encoding=\"utf-16BE\"?>" + resourceXml)
                    : Encoding.UTF8.GetBytes(resourceXml),
                [resourcePayloadDirectory + "/logo.png"] = [1, 2, 3],
                [resourcePayloadDirectory + "/typeface.ttf"] = [4, 5, 6]
            });
            withAnnotations.Dispose();
        }
        if (!version2) return withResources;
        using (withResources)
            return await Rewrite(withResources, "OFD.xml", bytes =>
                Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(bytes).Replace("Version=\"1.0\"", "Version=\"2.0\"")));
    }

    private static async Task<MemoryStream> AddEntries(MemoryStream source,
        IReadOnlyDictionary<string, byte[]> additions)
    {
        source.Position = 0;
        var output = new MemoryStream();
        using (var input = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            var written = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in input.Entries)
            {
                written.Add(entry.FullName);
                using var stream = entry.Open();
                using var destination = target.CreateEntry(entry.FullName).Open();
                if (additions.TryGetValue(entry.FullName, out var replacement))
                    await destination.WriteAsync(replacement);
                else
                    await stream.CopyToAsync(destination);
            }
            foreach (var addition in additions)
            {
                if (written.Contains(addition.Key)) continue;
                using var destination = target.CreateEntry(addition.Key).Open();
                await destination.WriteAsync(addition.Value);
            }
        }
        output.Position = 0;
        return output;
    }

    private static async Task<MemoryStream> AddReferencedPage(MemoryStream source,
        string pagePath, IReadOnlyDictionary<string, byte[]>? otherEntries = null,
        byte[]? pageBytes = null)
    {
        source.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(source);
        var additions = new Dictionary<string, byte[]>
        {
            [pagePath] = pageBytes ?? archive.GetBytes("Doc_0/Pages/Page_0/Content.xml")
        };
        if (otherEntries is not null)
            foreach (var item in otherEntries) additions.Add(item.Key, item.Value);
        using var withPage = await AddEntries(source, additions);
        return await Rewrite(withPage, "Doc_0/Document.xml", bytes =>
        {
            var document = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            var page = document.Descendants().First(element => element.Name.LocalName == "Page");
            page.SetAttributeValue("BaseLoc", pagePath.Substring("Doc_0/".Length));
            return Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        });
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
