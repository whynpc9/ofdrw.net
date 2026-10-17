using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class FontIdentityRegressionTests
{
    [Fact]
    public void NativeEmbedsThePhysicalBoldFaceSelectedForLayout()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(System.IO.Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        var regular = File.ReadAllBytes(System.IO.Path.Combine(root.FullName,
            "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/style-metrics.ttf"));
        var boldFace = new Ofdrw.Net.Core.Fonts.OpenTypeFace(regular);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(boldFace.Table("OS/2"), 62, 32);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(boldFace.Table("OS/2"), 4, 700);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(boldFace.Table("head"), 44, 1);
        var bold = boldFace.Build();
        var faces = new[] { regular, bold };
        using var stream = new MemoryStream(bold);
        var family = SixLabors.Fonts.FontDescription.LoadDescription(stream).FontFamilyInvariantCulture;
        var directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ofdrw-bold-face-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            for (var i = 0; i < faces.Length; i++) File.WriteAllBytes(System.IO.Path.Combine(directory, i + ".ttf"), faces[i]);
            var options = new DocxConversionOptions(); options.FontDirectories.Add(directory);
            var paragraph = new BuiltInParagraphModel();
            paragraph.Inlines.Add(new BuiltInTextModel { Text = "AA", Format = { FontFamily = family, Bold = true } });
            var section = new BuiltInSectionModel(); section.Blocks.Add(paragraph);
            var model = new BuiltInDocumentModel(); model.Sections.Add(section);
            var package = new BuiltInOfdRenderer(options, new List<DocxConversionDiagnostic>(), default).Render(model);
            Assert.Equal(bold, Assert.Single(package.Fonts).Data);
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("Noto Sans CJK SC")]
    [InlineData("Noto Sans CJK SC Regular")]
    public void MissingNotoMustNotBecomeAViewerLocalSubstitute(string requested)
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-missing-noto-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var root = new DirectoryInfo(AppContext.BaseDirectory);
            while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
            // A filename alias deliberately supplies a system-name metric face.
            // It must not manufacture a Noto identity through cross-family aliases.
            File.Copy(Path.Combine(root.FullName, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"),
                Path.Combine(directory, "SimHei.ttf"));
            var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false };
            options.FontDirectories.Add(directory);
            var catalog = new DocxFontCatalog(options, new List<DocxConversionDiagnostic>(), default);
            Assert.False(catalog.ShouldEmbed("SimSun"));
            Assert.NotEqual("SimSun", catalog.Resolve("SimSun", false, false));
            Assert.Equal("SimSun", catalog.ResolveFallbackFamily(new[] { requested, "SimSun" }));
            Assert.False(catalog.ShouldEmbed(requested));
            Assert.Equal(requested, catalog.Resolve(requested, false, false));
        }
        finally { Directory.Delete(directory, true); }
    }
}

// Reuse the repository's pinned OFL bundle without installing a host font.
internal sealed class PinnedCjkFontDirectory : IDisposable
{
    internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ofdrw-pinned-cjk-" + Guid.NewGuid().ToString("N"));
    internal PinnedCjkFontDirectory()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(System.IO.Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        using var archive = System.IO.Compression.ZipFile.OpenRead(System.IO.Path.Combine(root.FullName,
            "tests/fixtures/trusted-fonts/sha256-40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59.zip"));
        var entry = archive.GetEntry("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf")!;
        Assert.Equal(21659544, entry.Length);
        var bytes = new byte[21659544];
        using var input = entry.Open(); input.ReadExactly(bytes);
        Assert.Equal(-1, input.ReadByte());
        Assert.Equal("3012A9B63F5ECA3E3B38F23A1BE5ED504675E394ABF8E7A4FA981506582C04AA",
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)));
        Directory.CreateDirectory(Path);
        File.WriteAllBytes(System.IO.Path.Combine(Path, "noto.ttf"), bytes);
    }
    public void Dispose() => Directory.Delete(Path, true);
}
