using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;
using UglyToad.PdfPig;

internal static class CnEnReleaseSamples
{
    internal static async Task Run(string directory, string baseline, string? fontDirectory)
    {
        var sourceSha = Environment.GetEnvironmentVariable("OFDRW_SOURCE_SHA");
        if (string.IsNullOrWhiteSpace(sourceSha) || sourceSha.Length != 40 || !sourceSha.All(Uri.IsHexDigit))
            throw new InvalidOperationException("Set OFDRW_SOURCE_SHA to the exact 40-character source commit before generating release evidence.");
        if (fontDirectory is not null && !Directory.Exists(fontDirectory))
            throw new DirectoryNotFoundException(fontDirectory);
        using (var sections = DocxReleaseFixtures.CreateMultiSectionDocx())
            File.WriteAllBytes(Path.Combine(directory, "headers-source.docx"), sections.ToArray());
        var image = DocxReleaseFixtures.CreateInlineJpegDocx("内嵌图片 / Native JPEG image fixture");
        using (image.Docx) File.WriteAllBytes(Path.Combine(directory, "image-source.docx"), image.Docx.ToArray());
        var regularDirectory = PrepareRegularFonts(directory, baseline);
        var records = new List<object>();
        foreach (var profile in fontDirectory is null ? new[] { "original" } : new[] { "original", "physical-bold" })
        foreach (var scenario in new[] { "layout", "headers", "image", "selected" })
        {
            if (profile == "physical-bold" && scenario != "layout") continue;
            var original = scenario == "layout" ? baseline : Path.Combine(directory, (scenario == "selected" ? "headers" : scenario) + "-source.docx");
            var source = original; // Both profiles use the exact same, unmodified DOCX input.
            foreach (var mode in new[] { "native", "default" })
            {
                var name = $"{scenario}-{profile}-{mode}";
                var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false };
                if (mode == "native") options.OfdMode = DocxToOfdMode.Native;
                options.FontDirectories.Add(profile == "physical-bold" ? fontDirectory! : regularDirectory);
                if (profile == "physical-bold")
                {
                    options.NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.Throw;
                }
                var ofd = Path.Combine(directory, name + ".ofd");
                DocxToOfdConversionResult result;
                using (var input = File.OpenRead(source))
                using (var output = File.Create(ofd))
                    result = await new DocxToOfdConverter(options).ConvertWithResultAsync(input, output, scenario == "selected" ? new[] { 2, 0, 2, -1, 99 } : null);
                using var read = File.OpenRead(ofd);
                var package = await new OfdReader().ReadAsync(read);
                Check(package, scenario, original, image.Jpeg);
                if (scenario == "selected" && !result.SourcePages.SequenceEqual(new[] { 2, 0, 2 }))
                    throw new InvalidDataException(name + ": selected source pages differ.");
                if (result.Mode != DocxToOfdMode.Native || !result.OriginalTextPreserved)
                    throw new InvalidDataException(name + ": expected native OFD with preserved source text.");
                if (profile == "physical-bold")
                {
                    if (result.Diagnostics.Any(diagnostic => diagnostic.Code == "DOCX_EMBEDDED_BOLD_SIMULATED") ||
                        package.Fonts.Any(font => font.Data.Length == 0))
                        throw new InvalidDataException(name + ": physical profile must embed every selected face without bold simulation.");
                    foreach (var text in package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Where(text => text.Weight == 700))
                        if (!package.Fonts.Single(font => font.Id == text.FontResourceId).Bold ||
                            XElement.Parse(text.SourceXml!).Attribute("Stroke")?.Value == "true")
                            throw new InvalidDataException(name + ": bold text did not select a physical bold face.");
                }
                var pdf = Path.Combine(directory, name + ".pdf");
                using (var input = File.OpenRead(ofd))
                using (var output = File.Create(pdf))
                    await new OfdToPdfConverter().ConvertAsync(input, output);
                using var document = PdfDocument.Open(pdf);
                if (document.NumberOfPages != package.Pages.Count) throw new InvalidDataException(name + ": PDF page count differs.");
                for (var index = 0; index < package.Pages.Count; index++)
                {
                    var expected = Compact(string.Concat(package.Pages[index].Elements.OfType<OfdTextElement>().Select(text => text.Text)));
                    if (Compact(document.GetPage(index + 1).Text) != expected)
                        throw new InvalidDataException(name + ": PDF lost, reordered or repeated source text on page " + (index + 1));
                }
                records.Add(new
                {
                    name, profile, scenario, requestedMode = mode, actualMode = result.Mode.ToString(),
                    source = Identity(source), originalSource = Identity(original), ofd = Identity(ofd), pdf = Identity(pdf),
                    pages = package.Pages.Count, sourcePages = result.SourcePages, sourceTextPreserved = result.OriginalTextPreserved,
                    compactTextByPage = package.Pages.Select(page => Compact(string.Concat(page.Elements.OfType<OfdTextElement>().Select(text => text.Text)))).ToArray(),
                    fonts = package.Fonts.Select(font => new { font.FontName, font.FamilyName, font.Bold, font.Italic, bytes = font.Data.Length, sha256 = Hash(font.Data) }).ToArray(),
                    diagnostics = result.Diagnostics.Select(diagnostic => new { diagnostic.Code, diagnostic.Message }).ToArray(),
                    functionalAssertions = "passed", pngVisualReview = "not-performed", macosPreview = "not-performed", copyFind = "not-performed"
                });
                Console.WriteLine($"{name}: {package.Pages.Count} pages; native OFD and OFD→PDF text assertions passed.");
            }
        }
        var manifest = new
        {
            schemaVersion = 1, sourceSha, generatedAtUtc = DateTimeOffset.UtcNow, runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            os = System.Runtime.InteropServices.RuntimeInformation.OSDescription, chain = "DOCX → native/default OFD → PDF",
            regularFontFiles = Directory.GetFiles(regularDirectory).OrderBy(path => path).Select(Identity).ToArray(),
            physicalFontDirectory = fontDirectory,
            physicalFontFiles = fontDirectory is null ? Array.Empty<object>() : Directory.GetFiles(fontDirectory).OrderBy(path => path).Select(Identity).ToArray(),
            originalQuantityVisualFailure = "Historical FAIL remains; a physical-bold profile is a separate configuration, never a replacement acceptance.",
            records
        };
        File.WriteAllText(Path.Combine(directory, "cn-en-release-manifest.json"), JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static void Check(OfdDocumentPackage package, string scenario, string source, byte[] jpeg)
    {
        var pages = package.Pages;
        var texts = pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().ToArray();
        var expectedPages = scenario == "layout" ? 2 : scenario == "headers" ? 5 : scenario == "selected" ? 3 : 1;
        if (pages.Count != expectedPages) throw new InvalidDataException(scenario + ": unexpected page count.");
        if (scenario is "headers" or "selected")
        {
            var headers = new[] { "First Header", "Even Header", "Odd Header", "Even Header", "Odd Header" };
            var numbers = new[] { 1, 2, 3, 10, 11 };
            var bodies = new[] { "one", "two", "three", "four", "five" };
            for (var index = 0; index < pages.Count; index++)
            {
                var sourceIndex = scenario == "selected" ? new[] { 2, 0, 2 }[index] : index;
                var pageText = Compact(string.Concat(pages[index].Elements.OfType<OfdTextElement>().Select(text => text.Text)));
                if (pageText != $"Body{bodies[sourceIndex]}{Compact(headers[sourceIndex])}Page{numbers[sourceIndex]}/5")
                    throw new InvalidDataException("Header/footer, section numbering or body text differs on page " + (index + 1));
                var header = pages[index].Elements.OfType<OfdTextElement>().Single(text => text.Text.Contains("Header"));
                var body = pages[index].Elements.OfType<OfdTextElement>().Single(text => text.Text.Contains("Body"));
                var footer = pages[index].Elements.OfType<OfdTextElement>().First(text => text.Text.Contains("Page"));
                if (header.YMillimeters + header.HeightMillimeters >= body.YMillimeters || footer.YMillimeters <= 260 ||
                    footer.YMillimeters + footer.HeightMillimeters >= pages[index].HeightMillimeters)
                    throw new InvalidDataException("Header/footer lies outside its expected page region.");
            }
        }
        else
        {
            using var archive = ZipFile.OpenRead(source);
            using var xml = archive.GetEntry("word/document.xml")!.Open();
            var ns = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var expected = Compact(string.Concat(XDocument.Load(xml).Descendants(ns + "t").Select(text => text.Value)));
            if (Compact(string.Concat(texts.Select(text => text.Text))) != expected)
                throw new InvalidDataException(scenario + ": raw OFD source text differs.");
        }
        if (scenario == "layout")
        {
            var normal = texts.Single(text => text.Text == "第二页用于确认分页保持稳定。");
            var emphasis = texts.Single(text => text.Text.Contains("这段文字应为红色粗体。"));
            var subtitle = texts.Single(text => text.Text.Contains("Generated DOCX"));
            if (texts.Count(text => text.Weight == 700) != 7 || normal.Weight != 400 || normal.FillColor != OfdColor.Black ||
                emphasis.Weight != 700 || emphasis.FillColor != new OfdColor(192, 0, 0) || !subtitle.Italic)
                throw new InvalidDataException("Core local bold/italic/color intent differs or leaked into ordinary text.");
            if (pages[0].Elements.OfType<OfdPathElement>().Count(path => path.Fill && path.FillColor == new OfdColor(217, 234, 247)) != 3)
                throw new InvalidDataException("Core table header fills differ.");
        }
        var images = pages.SelectMany(page => page.Elements).OfType<OfdImageElement>().ToArray();
        if (scenario == "image" && (images.Length != 1 || images[0].MediaType != "image/jpeg" || !images[0].Data.SequenceEqual(jpeg) ||
            Math.Abs(images[0].WidthMillimeters - 40) > .001 || Math.Abs(images[0].HeightMillimeters - 20) > .001))
            throw new InvalidDataException("Inline JPEG payload, media type or 40×20 mm size differs.");
    }

    private static string PrepareRegularFonts(string directory, string baseline)
    {
        const string archiveHash = "40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59";
        var archivePath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(baseline)!,
            "../../../tests/fixtures/trusted-fonts/sha256-" + archiveHash + ".zip"));
        if (Hash(File.ReadAllBytes(archivePath)) != archiveHash) throw new InvalidDataException("Trusted font bundle hash differs.");
        var target = Path.Combine(directory, "regular-fonts");
        Directory.CreateDirectory(target);
        using var archive = ZipFile.OpenRead(archivePath);
        foreach (var (member, hash) in new[]
        {
            ("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf", "3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa"),
            ("fonts/NotoSans-Regular.ttf", "b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5"),
            ("licenses/Ofdrw-CI-Noto-OFL.txt", "6a73f9541c2de74158c0e7cf6b0a58ef774f5a780bf191f2d7ec9cc53efe2bf2"),
            ("licenses/NotoSans-LICENSE", "0dab92d0544f7b233403f14b84a663bdbfa746982eda629e7f4f9ffe1b036feb")
        })
        {
            using var input = archive.GetEntry(member)!.Open(); using var bytes = new MemoryStream(); input.CopyTo(bytes);
            if (Hash(bytes.ToArray()) != hash) throw new InvalidDataException("Trusted font member hash differs: " + member);
            File.WriteAllBytes(Path.Combine(target, Path.GetFileName(member)), bytes.ToArray());
        }
        return target;
    }
    private static string Compact(string value) => new(value.Where(character => !char.IsWhiteSpace(character)).ToArray());
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static object Identity(string path) => new { path = Path.GetFullPath(path), bytes = new FileInfo(path).Length, sha256 = Hash(File.ReadAllBytes(path)) };
}
