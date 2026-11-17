using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

/// <summary>A2: content-changing entry points refuse archival originals; plain roundtrip and explicit derivation stay available.</summary>
public sealed class ArchivalEditProtectionTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofd-archival-edit-" + Guid.NewGuid().ToString("N"));
    public ArchivalEditProtectionTests() => Directory.CreateDirectory(_directory);
    public void Dispose() { try { Directory.Delete(_directory, true); } catch (IOException) { } }
    private string PathFor(string name) => Path.Combine(_directory, name);

    public static TheoryData<string> PackageOperations() => new()
    {
        "OfdDocumentEditor.ReorderPages", "OfdDocumentEditor.RemovePages", "OfdDocumentEditor.CropPage",
        "OfdDocumentSplitter.Split", "OfdDocumentMerger.Merge", "OfdDocumentMixer.Mix",
        "OfdWatermark.AddText", "OfdWatermark.AddImage", "OfdNamedAreaEditor.FillText", "OfdNamedAreaEditor.FillImage"
    };

    private static OfdDocumentPackage? Run(string operation, OfdDocumentPackage package) => operation switch
    {
        "OfdDocumentEditor.ReorderPages" => Do(() => OfdDocumentEditor.ReorderPages(package, [1, 0])),
        "OfdDocumentEditor.RemovePages" => Do(() => OfdDocumentEditor.RemovePages(package, [1])),
        "OfdDocumentEditor.CropPage" => Do(() => OfdDocumentEditor.CropPage(package, package.Pages[0], 5, 5, 90, 90)),
        "OfdDocumentSplitter.Split" => OfdDocumentSplitter.Split(package, [1]),
        "OfdDocumentMerger.Merge" => OfdDocumentMerger.Merge([package, package]),
        "OfdDocumentMixer.Mix" => OfdDocumentMixer.Mix([new OfdMixSource(package, 0), new OfdMixSource(package, 1)]),
        "OfdWatermark.AddText" => Do(() => OfdWatermark.AddText(package, [0], "DRAFT")),
        "OfdWatermark.AddImage" => Do(() => OfdWatermark.AddImage(package, [0], DocumentToolTests.Png, "image/png")),
        "OfdNamedAreaEditor.FillText" => Do(() => new OfdNamedAreaEditor().FillText(package, "name", "Alice")),
        "OfdNamedAreaEditor.FillImage" => Do(() => new OfdNamedAreaEditor().FillImage(package, "name", DocumentToolTests.Png, "image/png", 1, 1)),
        _ => throw new ArgumentOutOfRangeException(nameof(operation))
    };
    private static OfdDocumentPackage? Do(Action action) { action(); return null; }

    [Theory]
    [MemberData(nameof(PackageOperations))]
    public async Task PackageOperation_RejectsArchivalOriginalWithoutMutation(string operation)
    {
        var archival = await Load(await Bytes(Source(OfdDocTypes.Archival)));
        var before = await Bytes(archival);
        var error = Assert.Throws<OfdArchivalOriginalException>(() => Run(operation, archival));
        Assert.Equal(operation, error.Operation);
        Assert.Contains("OFD-A", error.Message);
        Assert.Contains("derive-ofd", error.Message);
        Assert.Equal(Text(before), Text(await Bytes(archival)));
    }

    [Theory]
    [MemberData(nameof(PackageOperations))]
    public async Task PackageOperation_StillEditsOrdinaryOfd(string operation)
    {
        var ordinary = await Load(await Bytes(Source(OfdDocTypes.Ordinary)));
        var before = Text(await Bytes(ordinary));
        var output = Run(operation, ordinary) ?? ordinary;
        var saved = await Load(await Bytes(output));
        Assert.Equal(OfdDocTypes.Ordinary, saved.Options.DocType);
        Assert.NotEqual(before, Text(await Bytes(saved)));
    }

    [Fact]
    public async Task ExplicitOrdinaryDeclarationBeforeEditingStaysAllowed()
    {
        var package = await Load(await Bytes(Source(OfdDocTypes.Archival)));
        package.Options.DocType = OfdDocTypes.Ordinary;
        OfdDocumentEditor.RemovePages(package, [1]);
        var saved = await Load(await Bytes(package));
        Assert.Equal(OfdDocTypes.Ordinary, saved.Options.DocType);
        Assert.Single(saved.Pages);
    }

    [Fact]
    public async Task UnchangedRoundTripKeepsExplicitArchivalDeclaration()
    {
        var first = await Load(await Bytes(Source(OfdDocTypes.Archival)));
        var second = await Load(await Bytes(first));
        Assert.Equal(OfdDocTypes.Archival, second.Options.DocType);
        Assert.Equal("OFD-A", RootDocType(await Bytes(second)));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task MultiInputOperations_RejectArchivalInputAtAnyPosition(int archivalPosition)
    {
        var sources = new List<OfdDocumentPackage>();
        for (var i = 0; i < 3; i++)
            sources.Add(await Load(await Bytes(Source(i == archivalPosition ? OfdDocTypes.Archival : OfdDocTypes.Ordinary))));
        Assert.Equal("OfdDocumentMerger.Merge", Assert.Throws<OfdArchivalOriginalException>(() => OfdDocumentMerger.Merge(sources)).Operation);
        Assert.Throws<OfdArchivalOriginalException>(() => OfdDocumentMerger.MergeWithResult(sources, null, CancellationToken.None));
        Assert.Equal("OfdDocumentMixer.Mix", Assert.Throws<OfdArchivalOriginalException>(
            () => OfdDocumentMixer.Mix(sources.Select(source => new OfdMixSource(source, 0)))).Operation);
    }

    [Fact]
    public async Task TextReplacerRejectsArchivalOriginalBeforeMeasuring()
    {
        var archival = await Load(await Bytes(Source(OfdDocTypes.Archival)));
        var replacer = new OfdTextReplacer(new UnusedGeometryProvider());
        Assert.Equal("OfdTextReplacer.Replace", Assert.Throws<OfdArchivalOriginalException>(() => replacer.Replace(archival, "FIRST", "LAST")).Operation);
        Assert.Throws<OfdArchivalOriginalException>(() => replacer.Replace(archival, [0], "FIRST", "LAST"));
        Assert.Throws<OfdArchivalOriginalException>(() => replacer.Replace(archival, Array.Empty<OfdTextReplacement>()));
    }

    [Theory]
    [InlineData("OFD-A", true)]
    [InlineData("OFD", false)]
    public async Task SignatureCleaner_RejectsArchivalOriginalBeforeWriting(string docType, bool rejected)
    {
        var bytes = await Bytes(Source(docType));
        using var input = new MemoryStream(bytes);
        using var output = new MemoryStream();
        if (rejected)
        {
            var error = await Assert.ThrowsAsync<OfdArchivalOriginalException>(() => OfdSignatureCleaner.CleanAsync(input, output));
            Assert.Equal("OfdSignatureCleaner.CleanAsync", error.Operation);
            Assert.Equal(0, output.Length);
        }
        else
        {
            await OfdSignatureCleaner.CleanAsync(input, output);
            Assert.Equal("OFD", RootDocType(output.ToArray()));
        }
    }

    [Fact]
    public async Task Derivation_RewritesOnlyRootDeclarationAndRecordsIdentities()
    {
        var input = PathFor("archival.ofd"); var output = PathFor("derived.ofd");
        await File.WriteAllBytesAsync(input, await Bytes(Source(OfdDocTypes.Archival)));
        var sourceHash = Sha(input);
        var result = await OfdOrdinaryDerivation.CreateAsync(input, output);

        Assert.Equal(OfdOrdinaryDerivation.OperationName, result.Operation);
        Assert.Equal(Path.GetFullPath(input), result.SourcePath);
        Assert.Equal(sourceHash, result.SourceSha256);
        Assert.Equal(new FileInfo(input).Length, result.SourceLength);
        Assert.Equal("OFD-A", result.SourceDocType);
        Assert.Equal(Path.GetFullPath(output), result.OutputPath);
        Assert.Equal(Sha(output), result.OutputSha256);
        Assert.Equal(new FileInfo(output).Length, result.OutputLength);
        Assert.Equal("OFD", result.OutputDocType);
        Assert.Equal("OFD.xml", result.ChangedEntry);
        Assert.False(result.InheritsArchivalConclusion);
        Assert.Equal(sourceHash, Sha(input));

        var before = Entries(await File.ReadAllBytesAsync(input));
        var after = Entries(await File.ReadAllBytesAsync(output));
        Assert.Equal(before.Keys.OrderBy(x => x, StringComparer.Ordinal), after.Keys.OrderBy(x => x, StringComparer.Ordinal));
        Assert.Equal(new[] { "OFD.xml" }, before.Keys.Where(key => !before[key].SequenceEqual(after[key])));
        var originalRoot = Encoding.UTF8.GetString(before["OFD.xml"]);
        Assert.Equal(originalRoot.Replace("DocType=\"OFD-A\"", "DocType=\"OFD\""), Encoding.UTF8.GetString(after["OFD.xml"]));

        // The derivative is an ordinary OFD and can be edited.
        var derived = await Load(await File.ReadAllBytesAsync(output));
        OfdDocumentEditor.RemovePages(derived, [1]);
        Assert.Single((await Load(await Bytes(derived))).Pages);
    }

    public static TheoryData<string> DerivationRejections() => new()
    {
        "not-archival", "signed-entry", "signed-declaration", "encrypted", "same-path", "alias", "existing-output", "non-literal-declaration"
    };

    [Theory]
    [MemberData(nameof(DerivationRejections))]
    public async Task Derivation_RefusesWithoutOutputAndKeepsSource(string scenario)
    {
        var input = PathFor("source.ofd"); var output = PathFor("derived.ofd");
        await File.WriteAllBytesAsync(input, await Bytes(Source(scenario == "not-archival" ? OfdDocTypes.Ordinary : OfdDocTypes.Archival)));
        switch (scenario)
        {
            case "signed-entry": AddEntry(input, "Doc_0/Signs/Signatures.xml", "<Signatures xmlns='http://www.ofdspec.org'/>"); break;
            case "signed-declaration": EditRoot(input, root => root.Element(root.Name.Namespace + "DocBody")!.Add(new XElement(root.Name.Namespace + "Signatures", "Doc_0/Signs/Signatures.xml"))); break;
            case "encrypted": AddEntry(input, "Encryptions.xml", "<Encryptions xmlns='http://www.ofdspec.org'/>"); break;
            case "non-literal-declaration": ReplaceRootText(input, "DocType=\"OFD-A\"", "DocType=\"OFD&#45;A\""); break;
            case "existing-output": await File.WriteAllTextAsync(output, "keep"); break;
        }
        if (scenario == "same-path") output = input;
        if (scenario == "alias")
        {
            output = PathFor("link.ofd");
            try { File.CreateSymbolicLink(output, input); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) { return; } // Symlinks unavailable on this host.
        }
        var sourceHash = Sha(input);
        var files = Directory.GetFiles(_directory).OrderBy(x => x, StringComparer.Ordinal).ToArray();

        var failure = await Record.ExceptionAsync(() => OfdOrdinaryDerivation.CreateAsync(input, output));

        Assert.NotNull(failure);
        Assert.IsNotType<OfdArchivalOriginalException>(failure);
        Assert.True(scenario switch
        {
            "not-archival" => failure is InvalidOperationException && failure.Message.Contains("does not declare DocType=\"OFD-A\""),
            "signed-entry" or "signed-declaration" => failure is NotSupportedException && failure.Message.Contains("signed"),
            "encrypted" => failure is NotSupportedException && failure.Message.Contains("encrypted"),
            "non-literal-declaration" => failure is NotSupportedException,
            _ => failure is IOException
        }, failure.ToString());
        Assert.Equal(sourceHash, Sha(input));
        Assert.Equal(files, Directory.GetFiles(_directory).OrderBy(x => x, StringComparer.Ordinal).ToArray());
        if (scenario == "existing-output") Assert.Equal("keep", await File.ReadAllTextAsync(output));
    }

    [Fact]
    public async Task Derivation_CancellationLeavesNoOutput()
    {
        var input = PathFor("source.ofd"); var output = PathFor("derived.ofd");
        await File.WriteAllBytesAsync(input, await Bytes(Source(OfdDocTypes.Archival)));
        var sourceHash = Sha(input);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfdOrdinaryDerivation.CreateAsync(input, output, cancellationToken: cancellation.Token));
        Assert.Equal(new[] { input }, Directory.GetFiles(_directory));
        Assert.Equal(sourceHash, Sha(input));
    }

    private sealed class UnusedGeometryProvider : IOfdGlyphGeometryProvider
    {
        public IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits limits, CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("The archival guard must reject before any measurement.");
    }

    private static OfdDocumentPackage Source(string docType)
    {
        var flow = new FlowDocument();
        flow.Blocks.Add(new Paragraph("FIRST"));
        flow.Blocks.Add(new AreaHolderBlock("name", 18));
        flow.Blocks.Add(new Paragraph("SECOND") { PageBreakBefore = true });
        var package = flow.Render();
        package.Options.DocType = docType;
        return package;
    }

    private static async Task<byte[]> Bytes(OfdDocumentPackage package)
    {
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        return stream.ToArray();
    }

    private static async Task<OfdDocumentPackage> Load(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return await new OfdReader().ReadAsync(stream);
    }

    private static string Text(byte[] package) =>
        string.Join("\n", Entries(package).Where(pair => pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(pair => pair.Key, StringComparer.Ordinal).Select(pair => pair.Key + "=" + Encoding.UTF8.GetString(pair.Value)));

    private static string? RootDocType(byte[] package) =>
        XDocument.Parse(Encoding.UTF8.GetString(Entries(package)["OFD.xml"])).Root!.Attribute("DocType")?.Value;

    private static Dictionary<string, byte[]> Entries(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var stream = entry.Open(); using var copy = new MemoryStream(); stream.CopyTo(copy); return copy.ToArray();
        }, StringComparer.Ordinal);
    }

    private static void AddEntry(string path, string name, string content)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        using var stream = zip.CreateEntry(name).Open();
        stream.Write(Encoding.UTF8.GetBytes(content));
    }

    private static void EditRoot(string path, Action<XElement> edit)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = zip.GetEntry("OFD.xml")!;
        XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
        edit(xml.Root!);
        entry.Delete();
        using var target = zip.CreateEntry("OFD.xml").Open();
        target.Write(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
    }

    private static void ReplaceRootText(string path, string oldValue, string newValue)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = zip.GetEntry("OFD.xml")!;
        string text; using (var reader = new StreamReader(entry.Open())) text = reader.ReadToEnd();
        Assert.Contains(oldValue, text);
        entry.Delete();
        using var target = zip.CreateEntry("OFD.xml").Open();
        target.Write(Encoding.UTF8.GetBytes(text.Replace(oldValue, newValue)));
    }

    private static string Sha(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
