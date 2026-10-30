using System.Xml.Linq;
using System.Globalization;
using System.IO.Compression;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class EmbeddedBoldCompatibilityTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MixedNativeSpanKeepsItsPlannedStrokeForEveryPositionedGlyph(bool explicitNative)
    {
        using var fonts = new PinnedCjkFontDirectory();
        using var source = BoldSample("Noto Sans CJK SC", "量A", 48);
        using var ofd = new MemoryStream();
        await new DocxToOfdConverter(Options(fonts, explicitNative)).ConvertAsync(source, ofd);
        ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
        var text = Assert.Single(package.Pages.Single().Elements.OfType<OfdTextElement>());
        Assert.Equal("量A", text.Text);
        Assert.NotEmpty(Assert.Single(package.Fonts).Data);
        var xml = XElement.Parse(text.SourceXml!);
        var width = double.Parse(xml.Attribute("LineWidth")!.Value, CultureInfo.InvariantCulture);
        Assert.InRange(width / text.FontSizeMillimeters, .019, .022);
        Assert.Equal(700, text.Weight);
        var originalRuns = text.Runs.Select(run => (run.Text, run.XMillimeters, run.YMillimeters,
            run.DeltaX, run.DeltaY)).ToArray();
        using var actual = new MemoryStream(); ofd.Position = 0;
        await new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(ofd, actual);
        // Same original face, positions and declared native stroke, with only
        // the requested weight cleared: the contour already supplies emphasis.
        text.Weight = 400; xml.SetAttributeValue("Weight", "400");
        text.SourceXml = xml.ToString(SaveOptions.DisableFormatting);
        using var referenceOfd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, referenceOfd);
        referenceOfd.Position = 0; var referencePackage = await new OfdReader().ReadAsync(referenceOfd);
        var referenceText = Assert.Single(referencePackage.Pages.Single().Elements.OfType<OfdTextElement>());
        Assert.Equal(Assert.Single(package.Fonts).Data, Assert.Single(referencePackage.Fonts).Data);
        Assert.Equal(width, double.Parse(XElement.Parse(referenceText.SourceXml!).Attribute("LineWidth")!.Value, CultureInfo.InvariantCulture));
        Assert.Equal(originalRuns, referenceText.Runs.Select(run => (run.Text, run.XMillimeters,
            run.YMillimeters, run.DeltaX, run.DeltaY)));
        using var expected = new MemoryStream(); referenceOfd.Position = 0;
        await new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(referenceOfd, expected);
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } artifacts)
        {
            Directory.CreateDirectory(artifacts);
            var prefix = Path.Combine(artifacts, $"mixed-planned-{explicitNative}");
            await File.WriteAllBytesAsync(prefix + ".docx", source.ToArray());
            await File.WriteAllBytesAsync(prefix + ".ofd", ofd.ToArray());
            await File.WriteAllBytesAsync(prefix + ".pdf", actual.ToArray());
            await File.WriteAllBytesAsync(prefix + "-reference.ofd", referenceOfd.ToArray());
            await File.WriteAllBytesAsync(prefix + "-reference.pdf", expected.ToArray());
        }
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(actual.ToArray());
        Assert.Equal("量A", semantic.GetPage(1).Text);
        using var actualReader = Docnet.Core.DocLib.Instance.GetDocReader(actual.ToArray(), new Docnet.Core.Models.PageDimensions(4d));
        using var expectedReader = Docnet.Core.DocLib.Instance.GetDocReader(expected.ToArray(), new Docnet.Core.Models.PageDimensions(4d));
        using var actualPage = actualReader.GetPageReader(0); using var expectedPage = expectedReader.GetPageReader(0);
        Assert.Equal(expectedPage.GetImage(), actualPage.GetImage());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RegularEmbeddedFaceHasTruthfulFlagsAndExplicitSemanticBoldStroke(bool explicitNative)
    {
        using var fonts = new PinnedCjkFontDirectory();
        var options = Options(fonts, explicitNative);
        using var source = File.OpenRead(Sample()); using var output = new MemoryStream();
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(source, output);
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } artifacts)
        {
            Directory.CreateDirectory(artifacts);
            await File.WriteAllBytesAsync(Path.Combine(artifacts, $"native-bold-{explicitNative}.ofd"), output.ToArray());
        }
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_EMBEDDED_BOLD_SIMULATED");
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, package.Pages.Count);
        Assert.Equal(3, package.Fonts.Count);
        Assert.All(package.Fonts, font => { Assert.NotEmpty(font.Data); Assert.False(font.Bold); Assert.False(font.Italic); });
        var texts = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().ToArray();
        var bold = texts.Where(text => text.Weight == 700).ToArray();
        Assert.Equal(7, bold.Length); // headings, three table cells, second heading and red emphasis
        foreach (var text in texts)
        {
            var xml = XElement.Parse(text.SourceXml!);
            Assert.Equal(text.Text, string.Concat(xml.Descendants().Where(e => e.Name.LocalName == "TextCode").Select(e => e.Value)));
            if (text.Weight == 700)
            {
                Assert.Equal("true", xml.Attribute("Stroke")?.Value);
                Assert.Equal("true", xml.Attribute("Fill")?.Value);
                var width = double.Parse(xml.Attribute("LineWidth")!.Value, CultureInfo.InvariantCulture);
                if (text.Text == "数量") Assert.InRange(width / text.FontSizeMillimeters, .0095, .012);
                else Assert.InRange(width / text.FontSizeMillimeters, .034, .0351);
                var color = Assert.Single(xml.Elements(), e => e.Name.LocalName == "StrokeColor");
                Assert.Equal($"{text.FillColor.Red} {text.FillColor.Green} {text.FillColor.Blue}", color.Attribute("Value")?.Value);
            }
            else Assert.Null(xml.Attribute("Stroke"));
        }
        output.Position = 0;
        using var png = new MemoryStream();
        await new Ofdrw.Net.Converter.Pdf.Converters.OfdToImageConverter().ConvertAsync(output, png);
        AssertDenseCounterOpen(png.ToArray());
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } imageArtifacts)
            await File.WriteAllBytesAsync(Path.Combine(imageArtifacts, $"native-bold-{explicitNative}-public144.png"), png.ToArray());
        // Rewriting imported XML keeps its single text object/stroke; no second application.
        using var rewritten = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, rewritten);
        rewritten.Position = 0; var reread = await new OfdReader().ReadAsync(rewritten);
        Assert.Equal(texts.Select(text => text.SourceXml), reread.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(text => text.SourceXml));
    }

    private static void AssertDenseCounterOpen(byte[] png)
    {
        using var image = SixLabors.ImageSharp.Image.Load<SixLabors.ImageSharp.PixelFormats.Rgb24>(png);
        const double x = 79.117, y = 71.091, size = 3.704, advance = 3.704167, ppm = 144d / 25.4;
        var left = (int)Math.Ceiling((x + advance + size * .3) * ppm);
        var right = (int)Math.Floor((x + advance + size * .7) * ppm);
        double Brightest(double bottom, double top)
        {
            var first = (int)Math.Floor((y + size - size * top / 1000) * ppm);
            var last = (int)Math.Ceiling((y + size - size * bottom / 1000) * ppm);
            return Enumerable.Range(first, last - first).Max(row =>
                Enumerable.Range(left, right - left + 1).Average(column => image[column, row].R));
        }
        // Original pinned-font 量 apertures, same public144-DPI threshold as
        // the renderer regression. This exercises the actual DOCX producer.
        Assert.True(Brightest(610, 665) > 100 && Brightest(709, 763) > 100);
        Assert.Contains(new[] { 780d, 790d, 800d }, designY =>
        {
            var row = (int)Math.Round((y + size - size * designY / 1000) * ppm);
            return Enumerable.Range(left, right - left + 1).Count(column => image[column, row].R < 150) >= 4;
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StrictBoldPolicyRefusesRegularFaceBeforeWritingOutput(bool explicitNative)
    {
        using var fonts = new PinnedCjkFontDirectory();
        var options = Options(fonts, explicitNative); options.NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.Throw;
        using var source = File.OpenRead(Sample()); using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new DocxToOfdConverter(options).ConvertAsync(source, output));
        Assert.Contains("actual bold face", error.Message);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task EmbeddedSourceAndNonDocxTextDoNotAcquireCompatibilityPaint()
    {
        using var fonts = new PinnedCjkFontDirectory();
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto", Data = File.ReadAllBytes(Path.Combine(fonts.Path, "noto.ttf")) });
        var text = new OfdTextElement { FontResourceId = "10", FontName = "Noto", Text = "A", Weight = 700 };
        package.Pages.Add(new OfdPage { Elements = { text } });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0; var reread = await new OfdReader().ReadAsync(output);
        var xml = XElement.Parse(Assert.Single(reread.Pages[0].Elements.OfType<OfdTextElement>()).SourceXml!);
        Assert.Null(xml.Attribute("Stroke"));
        var plainClone = Assert.IsType<OfdTextElement>(OfdModelCloner.CloneElement(text));
        Assert.False(plainClone.EmitSyntheticBoldStroke);
        text.EmitSyntheticBoldStroke = true;
        text.SyntheticBoldStrokeWidthMillimeters = .061;
        text.UseNameOnlyBoldSelection = true;
        var emphasizedClone = Assert.IsType<OfdTextElement>(OfdModelCloner.CloneElement(text));
        Assert.True(emphasizedClone.EmitSyntheticBoldStroke);
        Assert.Equal(.061, emphasizedClone.SyntheticBoldStrokeWidthMillimeters);
        Assert.True(emphasizedClone.UseNameOnlyBoldSelection);
        package.Pages[0].Elements[0] = emphasizedClone;
        using var clonedOutput = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, clonedOutput);
        clonedOutput.Position = 0; var clonedRead = await new OfdReader().ReadAsync(clonedOutput);
        var clonedXml = XElement.Parse(Assert.Single(clonedRead.Pages[0].Elements.OfType<OfdTextElement>()).SourceXml!);
        Assert.Equal("true", clonedXml.Attribute("Stroke")?.Value);
        Assert.Equal("0.061", clonedXml.Attribute("LineWidth")?.Value);
        xml.SetAttributeValue("Stroke", "false"); xml.SetAttributeValue("LineWidth", "0.123");
        text.SourceXml = xml.ToString(SaveOptions.DisableFormatting);
        text.EmitSyntheticBoldStroke = true; // Even a stale native hint cannot override a source envelope.
        package.Pages[0].Elements[0] = OfdModelCloner.CloneElement(text);
        using var preserved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, preserved);
        preserved.Position = 0; var read = await new OfdReader().ReadAsync(preserved);
        var actual = XElement.Parse(Assert.Single(read.Pages[0].Elements.OfType<OfdTextElement>()).SourceXml!);
        Assert.Equal("false", actual.Attribute("Stroke")?.Value);
        Assert.Equal("0.123", actual.Attribute("LineWidth")?.Value);
    }

    [Theory]
    [InlineData("Noto Sans CJK SC", false, false)]
    [InlineData("Noto Sans CJK SC", true, false)]
    [InlineData("SimSun", false, true)]
    [InlineData("SimSun", true, true)]
    [InlineData("SimHei", false, true)]
    [InlineData("SimHei", true, true)]
    public async Task FreshMissingNotoUsesSemanticBoldWhileViewerLocalFamiliesKeepCompatibilityStroke(
        string family, bool explicitNative, bool compatibilityStroke)
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-name-only-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false };
            options.FontDirectories.Add(directory);
            if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
            using var source = BoldSample(family, "回A", 36); using var output = new MemoryStream();
            await new DocxToOfdConverter(options).ConvertAsync(source, output);
            output.Position = 0; var package = await new OfdReader().ReadAsync(output);
            var font = Assert.Single(package.Fonts); Assert.Empty(font.Data); Assert.Equal(family, font.FamilyName);
            var text = Assert.Single(package.Pages.Single().Elements.OfType<OfdTextElement>());
            Assert.Equal("回A", text.Text); Assert.Equal(700, text.Weight);
            var original = XElement.Parse(text.SourceXml!);
            Assert.Equal(compatibilityStroke ? "true" : null, original.Attribute("Stroke")?.Value);
            if (compatibilityStroke)
                Assert.Equal(Math.Max(.08, text.FontSizeMillimeters * .045),
                    double.Parse(original.Attribute("LineWidth")!.Value, CultureInfo.InvariantCulture), 3);
            // A memory-only producer choice must survive reader/writer cycles
            // through the generated standard OFD paint, not a guessed width.
            for (var pass = 0; pass < 2; pass++)
            {
                using var rewritten = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, rewritten);
                rewritten.Position = 0; package = await new OfdReader().ReadAsync(rewritten);
                Assert.Equal(family, Assert.Single(package.Fonts).FamilyName);
                var again = Assert.Single(package.Pages.Single().Elements.OfType<OfdTextElement>());
                Assert.Equal(700, again.Weight); Assert.Equal("回A", again.Text);
                Assert.True(XNode.DeepEquals(original, XElement.Parse(again.SourceXml!)));
            }
            if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } artifacts)
            {
                Directory.CreateDirectory(artifacts);
                await File.WriteAllBytesAsync(Path.Combine(artifacts, $"name-only-{family.Replace(' ', '-')}-{explicitNative}.ofd"), output.ToArray());
            }
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("thin")]
    [InlineData("different-color")]
    [InlineData("alpha")]
    [InlineData("outline-only")]
    [InlineData("no-stroke")]
    [InlineData("italic-ctm")]
    [InlineData("sim-sun-explicit")]
    [InlineData("omitted-paint")]
    [InlineData("italic-no-ctm")]
    public async Task ImportedNameOnlyPaintIsPreservedThroughRepeatedWrites(string kind)
    {
        var family = kind == "sim-sun-explicit" ? "SimSun" : "Noto Sans CJK SC";
        var alpha = kind == "alpha" ? 128 : 255; var italic = kind.StartsWith("italic", StringComparison.Ordinal);
        var defaultPaint = kind is "omitted-paint" or "italic-no-ctm";
        var xml = XElement.Parse("<TextObject xmlns='http://www.ofdspec.org' ID='3' Font='10' Boundary='10 10 40 15' " +
            "Size='4' Weight='700' Stroke='" + (kind == "no-stroke" ? "false" : "true") + "' Fill='" +
            (kind == "outline-only" ? "false" : "true") + "' LineWidth='0.023' CTM='1 0.1 0.2 1 2 3' " +
            (italic ? "Italic='true' " : "") + ">" +
            $"<FillColor Value='10 20 30' Alpha='{alpha}'/><StrokeColor Value='" +
            (kind == "different-color" ? "190 20 30" : "10 20 30") + $"' Alpha='{alpha}'/>" +
            "<TextCode X='0' Y='4' DeltaX='4'>回A</TextCode></TextObject>");
        if (defaultPaint)
        {
            foreach (var attribute in new[] { "Stroke", "Fill", "LineWidth", "CTM" }) xml.Attribute(attribute)?.Remove();
            xml.Elements().Where(element => element.Name.LocalName is "FillColor" or "StrokeColor").Remove();
        }
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = family, FamilyName = family, Bold = true });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100, Elements = { new OfdTextElement
        {
            FontName = family, FontResourceId = "10", Text = "回A", Weight = 700, Italic = italic,
            FontSizeMillimeters = 4, FillColor = defaultPaint ? OfdColor.Black : new OfdColor(10, 20, 30, (byte)alpha), SourceXml = xml.ToString()
        } } });
        for (var pass = 0; pass < 2; pass++)
        {
            using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
            output.Position = 0; package = await new OfdReader().ReadAsync(output);
            var actual = XElement.Parse(Assert.Single(package.Pages.Single().Elements.OfType<OfdTextElement>()).SourceXml!);
            actual.SetAttributeValue("ID", "3"); actual.SetAttributeValue("Font", "10");
            foreach (var node in xml.DescendantsAndSelf().Concat(actual.DescendantsAndSelf()))
                node.Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            Assert.True(XNode.DeepEquals(xml, actual), actual.ToString());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FreshNameOnlyBoldSelectionSurvivesCloneWithoutChangingGenericWriterCompatibility(bool semanticSelection)
    {
        var family = semanticSelection ? "Noto Sans CJK SC" : "Unbound Fixture Family";
        var original = new OfdTextElement
        {
            FontName = family, FontResourceId = "10", Text = "回A", Weight = 700,
            FontSizeMillimeters = 4, UseNameOnlyBoldSelection = semanticSelection
        };
        var clone = Assert.IsType<OfdTextElement>(OfdModelCloner.CloneElement(original));
        Assert.Equal(semanticSelection, clone.UseNameOnlyBoldSelection);
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = family, FamilyName = family, Bold = true });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100, Elements = { clone } });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0; var read = await new OfdReader().ReadAsync(output);
        var text = Assert.Single(read.Pages.Single().Elements.OfType<OfdTextElement>());
        Assert.Equal(700, text.Weight); Assert.Equal("回A", text.Text);
        var xml = XElement.Parse(text.SourceXml!);
        Assert.Equal(semanticSelection ? null : "true", xml.Attribute("Stroke")?.Value);
        if (!semanticSelection) Assert.Equal("0.18", xml.Attribute("LineWidth")?.Value);
    }

    private static MemoryStream BoldSample(string family, string value, int halfPoints)
    {
        var source = new MemoryStream(); source.Write(File.ReadAllBytes(Sample()));
        using (var zip = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/document.xml")!;
            XDocument document;
            using (var input = entry.Open()) document = XDocument.Load(input);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var body = document.Root!.Element(w + "body")!;
            var section = body.Element(w + "sectPr"); section?.Remove();
            body.ReplaceNodes(new XElement(w + "p", new XElement(w + "r",
                new XElement(w + "rPr", new XElement(w + "rFonts",
                    new XAttribute(w + "ascii", family), new XAttribute(w + "eastAsia", family), new XAttribute(w + "hAnsi", family)),
                    new XElement(w + "b"), new XElement(w + "sz", new XAttribute(w + "val", halfPoints))),
                new XElement(w + "t", value))));
            if (section is not null) body.Add(section);
            entry.Delete();
            using var output = zip.CreateEntry("word/document.xml").Open(); document.Save(output);
        }
        source.Position = 0; return source;
    }

    private static DocxConversionOptions Options(PinnedCjkFontDirectory fonts, bool explicitNative)
    {
        var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false };
        options.FontDirectories.Add(fonts.Path);
        if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        return options;
    }
    private static string Sample()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return Path.Combine(root.FullName, "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx");
    }
}
