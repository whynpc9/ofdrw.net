using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Converter.Docx.Tests;

// RequireEmbedded must embed the declared family's own bytes or fail. System Chinese names are
// exercised with an open fixture whose name table is rewritten at test time; no proprietary font is used.
public sealed class NativeFontSelfContainmentTests : IDisposable
{
    private const string LatinFixture = "Ofdrw Test Face";
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-self-contained-" + Guid.NewGuid().ToString("N"));

    public NativeFontSelfContainmentTests() => Directory.CreateDirectory(_directory);

    public void Dispose() => Directory.Delete(_directory, true);

    [Theory]
    [InlineData(Ofdrw.Net.Core.Models.OfdFontEmbeddingMode.SubsetWhenSafe)]
    [InlineData(Ofdrw.Net.Core.Models.OfdFontEmbeddingMode.Full)]
    public async Task RequireEmbeddedWritesAPayloadForEveryUsedFontIncludingSystemChineseNames(Ofdrw.Net.Core.Models.OfdFontEmbeddingMode mode)
    {
        var simSun = Renamed(Fixture("narrow.ttf"), "SimSun", "宋体");
        var testFace = Fixture("style-metrics.ttf");
        var fonts = FontDirectory("fonts", ("simsun.ttf", simSun), ("test-face.ttf", testFace));
        var options = Options(fonts);
        options.FontEmbedding.Mode = mode;
        var (package, diagnostics) = await ConvertAsync(
            Docx(("宋体", "中文 Abc", false), (LatinFixture, "Latin Abc", false), ("SimSun", "字体 Bold", true)), options);

        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var declared = FontDeclarations(archive).ToList();
        Assert.Contains(declared, font => font.Family == "SimSun");
        Assert.Contains(declared, font => font.Family == LatinFixture);
        foreach (var font in declared)
        {
            Assert.NotNull(font.File);
            var entry = archive.GetEntry(font.File!);
            Assert.NotNull(entry);
            using var stream = entry!.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            var face = new OpenTypeFace(copy.ToArray());
            Assert.True(face.Tables.ContainsKey("glyf"));
            // Subsets carry package-unique names by design; full payloads are the declared family's own bytes.
            if (mode == Ofdrw.Net.Core.Models.OfdFontEmbeddingMode.Full)
                Assert.Equal(font.Family == "SimSun" ? simSun : testFace, copy.ToArray());
        }
        var notice = Assert.Single(diagnostics, value => value.Code == "DOCX_FONT_LICENSE_NOT_VERIFIED");
        Assert.Contains("not a licence grant", notice.Message);
        Assert.Contains("caller's responsibility", notice.Message);
        Assert.Contains(diagnostics, value => value.Code == "DOCX_EMBEDDED_BOLD_SIMULATED");
    }

    [Fact]
    public async Task CompatibilityKeepsTheSameSystemChineseFamilyNameOnly()
    {
        var fonts = FontDirectory("fonts", ("simsun.ttf", Renamed(Fixture("narrow.ttf"), "SimSun", "宋体")));
        var options = Options(fonts);
        options.NativeFontSelfContainment = DocxNativeFontSelfContainment.Compatibility;
        var (package, diagnostics) = await ConvertAsync(Docx(("宋体", "中文 Abc", false)), options);
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        Assert.All(FontDeclarations(archive), font => Assert.Null(font.File));
        Assert.DoesNotContain(diagnostics, value => value.Code == "DOCX_FONT_LICENSE_NOT_VERIFIED");
    }

    [Fact]
    public void SystemChineseFamilyWithAuthenticBytesIsEmbeddedAndMeasuredWithThoseBytes()
    {
        var simSun = Renamed(Fixture("narrow.ttf"), "SimSun", "宋体");
        var fonts = FontDirectory("fonts", ("simsun.ttf", simSun));
        var package = Render(Options(fonts), ("SimSun", "AB AB"));
        var font = Assert.Single(package.Fonts);
        Assert.Equal(simSun, font.Data);
        Assert.Equal("SimSun", font.FamilyName);
        Assert.Equal("SimSun|regular", font.FontName);

        // Latin advances come from the embedded face, not the name-only Arial compatibility metrics.
        var wide = Renamed(Fixture("wide.ttf"), "SimSun", "宋体");
        var widePackage = Render(Options(FontDirectory("wide", ("simsun.ttf", wide))), ("SimSun", "AB AB"));
        Assert.NotEqual(Text(package).Runs.Single().DeltaX, Text(widePackage).Runs.Single().DeltaX);
    }

    [Theory]
    [InlineData("SimSun")]
    [InlineData("宋体")]
    [InlineData("Microsoft YaHei")]
    [InlineData("KaiTi")]
    public void SystemChineseFamilyWithoutBytesFailsNamingTheFamily(string family)
    {
        var fonts = FontDirectory("fonts", ("test-face.ttf", Fixture("narrow.ttf")));
        var exception = Assert.Throws<InvalidDataException>(() => Render(Options(fonts), (family, "AB")));
        Assert.Contains("DOCX_FONT_NOT_SELF_CONTAINED", exception.Message);
        Assert.Contains("'" + NormalizedName(family) + "'", exception.Message);
        Assert.Contains("FontDirectories", exception.Message);
    }

    [Fact]
    public void CrossFamilyStandInIsNeverEmbeddedUnderTheRequestedName()
    {
        // Compatibility measures SimSun with the SimHei face; strict mode must not embed it as SimSun.
        var fonts = FontDirectory("fonts", ("simhei.ttf", Renamed(Fixture("narrow.ttf"), "SimHei", "黑体")));
        var compatibility = Options(fonts);
        compatibility.NativeFontSelfContainment = DocxNativeFontSelfContainment.Compatibility;
        Assert.All(Render(compatibility, ("SimSun", "AB")).Fonts, font => Assert.Empty(font.Data));

        var exception = Assert.Throws<InvalidDataException>(() => Render(Options(fonts), ("SimSun", "AB")));
        Assert.Contains("'SimSun'", exception.Message);
        Assert.Contains("FontDirectories", exception.Message);
    }

    [Fact]
    public async Task FileNameStandInIsNeverEmbeddedUnderTheRequestedName()
    {
        // The file is named simsun.ttf, but its own name table records a different family.
        var fonts = FontDirectory("fonts", ("simsun.ttf", Fixture("narrow.ttf")));
        var output = new MemoryStream();
        var exception = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new DocxToOfdConverter(Options(fonts)).ConvertWithResultAsync(new MemoryStream(Docx(("SimSun", "AB", false))), output));
        Assert.Contains("'SimSun'", exception.Message);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(0x0002, "restricted-licence")]
    [InlineData(0x0200, "bitmap embedding only")]
    [InlineData(0x020A, "restricted-licence")]
    public void ForbiddenEmbeddingPermissionFailsNamingTheFont(int fsType, string reason)
    {
        var fonts = FontDirectory("fonts", ("simsun.ttf", Renamed(Fixture("narrow.ttf"), "SimSun", "宋体", fsType)));
        var exception = Assert.Throws<InvalidDataException>(() => Render(Options(fonts), ("SimSun", "AB")));
        Assert.Contains("'SimSun'", exception.Message);
        Assert.Contains("OS/2.fsType", exception.Message);
        Assert.Contains(reason, exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task NoSubsettingPermissionEmbedsTheFullFace()
    {
        var face = Renamed(Fixture("narrow.ttf"), "SimSun", "宋体", 0x0100);
        var fonts = FontDirectory("fonts", ("simsun.ttf", face));
        var (package, _) = await ConvertAsync(Docx(("SimSun", "AB", false)), Options(fonts));
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var file = Assert.Single(FontDeclarations(archive).Select(font => font.File).Distinct());
        Assert.Equal(face.Length, archive.GetEntry(file!)!.Length);
    }

    [Fact]
    public void NoSubsettingPermissionFailsWithTheGlyphSubsettingLayoutEngine()
    {
        var fonts = FontDirectory("fonts", ("simsun.ttf", Renamed(Fixture("narrow.ttf"), "SimSun", "宋体", 0x0100)));
        var options = Options(fonts);
        options.NativeParagraphLayoutEngine = HarfBuzzTextPrelayout.CreateParagraphLayout();
        var exception = Assert.Throws<InvalidDataException>(() => Render(options, ("SimSun", "AB")));
        Assert.Contains("'SimSun'", exception.Message);
        Assert.Contains("forbids subsetting", exception.Message);
    }

    [Fact]
    public void ParagraphLayoutEngineEmbedsTheDeclaredSystemChineseFamily()
    {
        var simSun = Renamed(Fixture("narrow.ttf"), "SimSun", "宋体");
        var options = Options(FontDirectory("fonts", ("simsun.ttf", simSun)));
        options.NativeParagraphLayoutEngine = HarfBuzzTextPrelayout.CreateParagraphLayout();
        var font = Assert.Single(Render(options, ("SimSun", "AB")).Fonts);
        Assert.Equal("SimSun", font.FamilyName);
        Assert.NotEmpty(font.Data);
    }

    [Theory]
    [InlineData("Noto Sans CJK SC")]
    [InlineData("Noto Sans CJK JP")]
    public void NotoCjkWithoutDirectoryBytesFails(string family)
    {
        var fonts = FontDirectory("empty");
        var exception = Assert.Throws<InvalidDataException>(() => Render(Options(fonts), (family, "AB")));
        Assert.Contains("'" + family + "'", exception.Message);
    }

    [Fact]
    public async Task DualLayerAndPdfRejectRequireEmbedded()
    {
        var options = Options(FontDirectory("empty"));
        options.OfdMode = DocxToOfdMode.DualLayer;
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new DocxToOfdConverter(options).ConvertWithResultAsync(new MemoryStream(Docx(("SimSun", "AB", false))), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new DocxToPdfConverter(Options(FontDirectory("empty"))).ConvertWithResultAsync(new MemoryStream(Docx(("SimSun", "AB", false))), new MemoryStream()));
    }

    [Fact]
    public async Task UnknownPolicyIsRejected()
    {
        var options = new DocxConversionOptions { NativeFontSelfContainment = (DocxNativeFontSelfContainment)42 };
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            new DocxToOfdConverter(options).ConvertWithResultAsync(new MemoryStream(Docx(("SimSun", "AB", false))), new MemoryStream()));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new Ofdrw.Net.Converter.Docx.Internal.BuiltIn.BuiltInOfdRenderer(options, new List<DocxConversionDiagnostic>(), default));
    }

    private static string NormalizedName(string family) => family == "宋体" ? "SimSun" : family;

    private static DocxConversionOptions Options(string fonts)
    {
        var options = new DocxConversionOptions
        {
            UseInstalledMicrosoftOfficeFonts = false,
            NativeFontSelfContainment = DocxNativeFontSelfContainment.RequireEmbedded
        };
        options.FontDirectories.Add(fonts);
        return options;
    }

    private static Ofdrw.Net.Core.Models.OfdDocumentPackage Render(DocxConversionOptions options, params (string Family, string Text)[] runs)
    {
        var section = new Ofdrw.Net.Converter.Docx.Internal.BuiltIn.BuiltInSectionModel();
        foreach (var (family, text) in runs)
        {
            var paragraph = new Ofdrw.Net.Converter.Docx.Internal.BuiltIn.BuiltInParagraphModel();
            paragraph.Inlines.Add(new Ofdrw.Net.Converter.Docx.Internal.BuiltIn.BuiltInTextModel { Text = text, Format = { FontFamily = family } });
            section.Blocks.Add(paragraph);
        }
        var model = new Ofdrw.Net.Converter.Docx.Internal.BuiltIn.BuiltInDocumentModel();
        model.Sections.Add(section);
        return new Ofdrw.Net.Converter.Docx.Internal.BuiltIn.BuiltInOfdRenderer(options, new List<DocxConversionDiagnostic>(), default).Render(model);
    }

    private static Ofdrw.Net.Core.Models.OfdTextElement Text(Ofdrw.Net.Core.Models.OfdDocumentPackage package) =>
        package.Pages.Single().Elements.OfType<Ofdrw.Net.Core.Models.OfdTextElement>().Single();

    private static async Task<(byte[] Package, IReadOnlyList<DocxConversionDiagnostic> Diagnostics)> ConvertAsync(byte[] docx, DocxConversionOptions options)
    {
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(new MemoryStream(docx), output);
        return (output.ToArray(), result.Diagnostics);
    }

    // Font declarations from every resource file, with FontFile resolved against that file's BaseLoc.
    private static IEnumerable<(string? Family, string? File)> FontDeclarations(ZipArchive archive)
    {
        foreach (var entry in archive.Entries.Where(entry => entry.FullName.EndsWith("Res.xml", StringComparison.Ordinal)))
        {
            XDocument document;
            using (var stream = entry.Open()) document = XDocument.Load(stream);
            var folder = Path.GetDirectoryName(entry.FullName)!.Replace('\\', '/');
            var baseLoc = document.Root!.Attribute("BaseLoc")?.Value;
            foreach (var font in document.Descendants().Where(element => element.Name.LocalName == "Font"))
            {
                var file = font.Elements().FirstOrDefault(element => element.Name.LocalName == "FontFile")?.Value;
                yield return (font.Attribute("FamilyName")?.Value ?? font.Attribute("FontName")?.Value,
                    file is null ? null : string.Join("/", new[] { folder, baseLoc, file }.Where(part => !string.IsNullOrEmpty(part))));
            }
        }
    }

    private string FontDirectory(string name, params (string File, byte[] Data)[] files)
    {
        var path = Path.Combine(_directory, name);
        Directory.CreateDirectory(path);
        foreach (var (file, data) in files) File.WriteAllBytes(Path.Combine(path, file), data);
        return path;
    }

    private static byte[] Fixture(string name)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return File.ReadAllBytes(Path.Combine(root.FullName, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts", name));
    }

    // Rewrites the naming table (and optionally OS/2.fsType) of an open fixture face, keeping its outlines.
    private static byte[] Renamed(byte[] source, string family, string? localizedFamily = null, int? fsType = null)
    {
        var face = new OpenTypeFace(source);
        var records = new List<(int Language, int NameId, string Value)>
        {
            (0x0409, 1, family), (0x0409, 2, "Regular"), (0x0409, 4, family),
            (0x0409, 6, family.Replace(" ", string.Empty) + "-Regular")
        };
        if (localizedFamily is not null)
        {
            records.Add((0x0804, 1, localizedFamily));
            records.Add((0x0804, 4, localizedFamily));
        }
        records.Sort((left, right) => left.Language != right.Language
            ? left.Language.CompareTo(right.Language) : left.NameId.CompareTo(right.NameId));
        var strings = new MemoryStream();
        var table = new byte[6 + records.Count * 12];
        OpenTypeFace.Put16(table, 2, records.Count);
        OpenTypeFace.Put16(table, 4, table.Length);
        for (var index = 0; index < records.Count; index++)
        {
            var bytes = Encoding.BigEndianUnicode.GetBytes(records[index].Value);
            var record = 6 + index * 12;
            OpenTypeFace.Put16(table, record, 3);
            OpenTypeFace.Put16(table, record + 2, 1);
            OpenTypeFace.Put16(table, record + 4, records[index].Language);
            OpenTypeFace.Put16(table, record + 6, records[index].NameId);
            OpenTypeFace.Put16(table, record + 8, bytes.Length);
            OpenTypeFace.Put16(table, record + 10, (int)strings.Length);
            strings.Write(bytes);
        }
        face.Tables["name"] = table.Concat(strings.ToArray()).ToArray();
        if (fsType is { } permissions) OpenTypeFace.Put16(face.Table("OS/2", 10), 8, permissions);
        return face.Build();
    }

    private static byte[] Docx(params (string Family, string Text, bool Bold)[] runs)
    {
        var body = new StringBuilder();
        foreach (var (family, text, bold) in runs)
        {
            var escaped = System.Security.SecurityElement.Escape(family);
            body.Append("<w:p><w:r><w:rPr><w:rFonts w:ascii=\"").Append(escaped).Append("\" w:hAnsi=\"").Append(escaped)
                .Append("\" w:eastAsia=\"").Append(escaped).Append("\"/>").Append(bold ? "<w:b/>" : string.Empty)
                .Append("</w:rPr><w:t xml:space=\"preserve\">").Append(System.Security.SecurityElement.Escape(text)).Append("</w:t></w:r></w:p>");
        }
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
                """);
            Write(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """);
            Write(archive, "word/document.xml",
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" + body + "</w:body></w:document>");
        }
        return output.ToArray();
    }

    private static void Write(ZipArchive archive, string name, string contents)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(contents);
    }
}
