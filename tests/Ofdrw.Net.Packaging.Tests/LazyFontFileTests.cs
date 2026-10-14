using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class LazyFontFileTests
{
    [Theory]
    [InlineData(false, false)] [InlineData(false, true)]
    [InlineData(true, false)] [InlineData(true, true)]
    public async Task AuthenticatedExistingLiteralDoesNotHashUnusedName(bool attribute, bool relocate)
    {
        var fixture = await Imported(font =>
        {
            if (!attribute) return;
            var child = font.Element(font.Name.Namespace + "FontFile")!;
            font.SetAttributeValue("FontFile", child.Value); child.Remove();
        });
        var (entries, material, catalog) = Catalog(fixture.Source);
        var destination = relocate ? "Doc_0/MovedRes.xml" : ResourcePath;
        var before = material.ValidationWorkBytes;
        if (attribute && relocate)
        {
            // Existing FontFile attributes are opaque for relocation in the
            // current catalog; naming must not broaden that contract.
            var error = Assert.Throws<NotSupportedException>(() => catalog.PlanFont(fixture.Source.Fonts[0].Id,
                fixture.Source.Fonts[0], destination, Ns));
            Assert.Contains("opaque resource", error.Message); Assert.Equal(before, material.ValidationWorkBytes); return;
        }
        catalog.PlanFont(fixture.Source.Fonts[0].Id, fixture.Source.Fonts[0], destination, Ns);
        Assert.Equal(before, material.ValidationWorkBytes);
        catalog.Apply(); catalog.Flush();
        var expected = relocate ? "/" + fixture.PayloadPath : fixture.Literal;
        var resource = XDocument.Load(new MemoryStream(entries[destination]));
        Assert.Equal(expected, resource.Descendants(Ns + "FontFile").Single().Value);
        Assert.Same(fixture.Bytes, entries[fixture.PayloadPath]);
    }

    [Theory]
    [InlineData(".ttf", ".ttf")] [InlineData(".otf", ".otf")] [InlineData(".bin", ".ttf")]
    public void MissingLiteralKeepsExactContentHashFilename(string inputExtension, string outputExtension)
    {
        var bytes = new byte[] { 1, 2, 3, 4 };
        var source = new OfdDocumentPackage();
        var (entries, material, catalog) = Catalog(source);
        var font = new OfdFontResource { Id = "1", FontName = "fixture", FileName = "font" + inputExtension, Data = bytes };
        catalog.PlanFont(font.Id, font, ResourcePath, Ns);
        Assert.Equal(bytes.Length, material.ValidationWorkBytes);
        catalog.Apply(); catalog.Flush();
        var name = "Font_" + Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant() + outputExtension;
        Assert.Equal(name, XDocument.Load(new MemoryStream(entries[ResourcePath])).Descendants(Ns + "FontFile").Single().Value);
        Assert.Same(bytes, entries["Doc_0/Res/" + name]);
    }

    [Theory]
    [InlineData("")] [InlineData(" ")] [InlineData("absent.ttf")]
    public async Task PresentLiteralKeepsBaselinePathSemanticsWithoutInventingFallback(string literal)
    {
        var fixture = await Imported(font => font.Element(font.Name.Namespace + "FontFile")!.Value = literal);
        var (entries, material, catalog) = Catalog(fixture.Source);
        var before = material.ValidationWorkBytes;
        catalog.PlanFont(fixture.Source.Fonts[0].Id, fixture.Source.Fonts[0], ResourcePath, Ns);
        Assert.Equal(before, material.ValidationWorkBytes);
        catalog.Apply(); catalog.Flush();
        Assert.Equal(string.IsNullOrWhiteSpace(literal) ? string.Empty : literal,
            XDocument.Load(new MemoryStream(entries[ResourcePath])).Descendants(Ns + "FontFile").Single().Value);
        // This preserves existing catalog behavior, including malformed/empty
        // literals; it does not extend Reader support or certify such input.
        Assert.DoesNotContain(entries.Keys.Except(fixture.Source.PreservedEntries.Keys), key => key.Contains("Font_", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task ExistingStructuralErrorPriorityAndProtectedRoleChecksRemain(bool relocateOpaque)
    {
        var fixture = await Imported(font =>
        {
            font.SetAttributeValue("FontFile", "/OFD.xml");
            if (relocateOpaque) font.Add(new XElement(XName.Get("Keep", "urn:vendor")));
        });
        var (_, _, catalog) = Catalog(fixture.Source);
        var error = Assert.Throws<NotSupportedException>(() => catalog.PlanFont(fixture.Source.Fonts[0].Id,
            fixture.Source.Fonts[0], relocateOpaque ? "Doc_0/MovedRes.xml" : ResourcePath, Ns));
        Assert.Contains(relocateOpaque ? "opaque resource" : "Conflicting payload", error.Message);

        var protectedFixture = await Imported(font => font.Element(font.Name.Namespace + "FontFile")!.Value = "/OFD.xml");
        var (_, _, protectedCatalog) = Catalog(protectedFixture.Source);
        Assert.Throws<InvalidDataException>(() => protectedCatalog.PlanFont(protectedFixture.Source.Fonts[0].Id,
            protectedFixture.Source.Fonts[0], ResourcePath, Ns));
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public async Task ImportedSamePathStillRequiresReferenceOrExactPayloadAgreement(int kind)
    {
        var fixture = await Imported(_ => { }, twoFonts: true);
        var first = fixture.Source.Fonts[0]; var second = fixture.Source.Fonts[1];
        second.Data = kind == 0 ? first.Data : first.Data.ToArray();
        if (kind == 2) second.Data[0] ^= 1;
        var (_, material, catalog) = Catalog(fixture.Source);
        catalog.PlanFont(first.Id, first, ResourcePath, Ns); var before = material.ValidationWorkBytes;
        if (kind == 2)
            Assert.Throws<InvalidDataException>(() => catalog.PlanFont(second.Id, second, ResourcePath, Ns));
        else
        {
            catalog.PlanFont(second.Id, second, ResourcePath, Ns);
            Assert.Equal(kind == 0 ? 0 : 2L * first.Data.Length, material.ValidationWorkBytes - before);
        }
    }

    private const string ResourcePath = "Doc_0/PublicRes.xml";
    private static readonly XNamespace Ns = OfdConstants.Namespace;
    private static (Dictionary<string, byte[]> Entries, OfdModelMaterialization Material, OfdResourceCatalog Catalog) Catalog(OfdDocumentPackage source)
    {
        var entries = new Dictionary<string, byte[]>(source.PreservedEntries, StringComparer.OrdinalIgnoreCase);
        var material = new OfdModelMaterialization(new OfdModelLimits());
        var catalog = new OfdResourceCatalog(entries, source.DocumentEntryPath ?? "Doc_0/Document.xml", Ns,
            [ResourcePath], material, default, source.SourceIdentity, source, ["OFD.xml"]);
        return (entries, material, catalog);
    }

    private static async Task<(OfdDocumentPackage Source, byte[] Bytes, string Literal, string PayloadPath)> Imported(Action<XElement> change, bool twoFonts = false)
    {
        var bytes = new byte[4096]; bytes[0] = 1; bytes[100] = 7;
        var source = new OfdDocumentPackage(); source.Pages.Add(new OfdPage());
        source.Fonts.Add(new OfdFontResource { Id = "100", FontName = "fixture", Data = bytes });
        if (twoFonts) source.Fonts.Add(new OfdFontResource { Id = "101", FontName = "second", Data = bytes });
        using var written = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, written); written.Position = 0;
        var entries = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        using (var zip = new ZipArchive(written, ZipArchiveMode.Read, true))
            foreach (var entry in zip.Entries)
            { using var input = entry.Open(); using var output = new MemoryStream(); await input.CopyToAsync(output); entries.Add(entry.FullName, output.ToArray()); }
        var resource = XDocument.Load(new MemoryStream(entries[ResourcePath]));
        var originalLiteral = resource.Descendants(Ns + "FontFile").First().Value;
        foreach (var font in resource.Descendants(Ns + "Font")) change(font);
        entries[ResourcePath] = Encoding.UTF8.GetBytes(resource.ToString(SaveOptions.DisableFormatting));
        using var modified = new MemoryStream();
        using (var zip = new ZipArchive(modified, ZipArchiveMode.Create, true))
            foreach (var entry in entries) { using var stream = zip.CreateEntry(entry.Key).Open(); await stream.WriteAsync(entry.Value); }
        modified.Position = 0; source = await new OfdReader().ReadAsync(modified);
        var payloadPath = "Doc_0/Res/" + originalLiteral;
        // The catalog already supports legacy FontFile attributes and present
        // missing/empty paths. Supply its payload explicitly; Reader support is
        // deliberately not changed by this naming optimization.
        bytes = source.PreservedEntries[payloadPath]; foreach (var font in source.Fonts) font.Data = bytes;
        return (source, bytes, originalLiteral, payloadPath);
    }
}
