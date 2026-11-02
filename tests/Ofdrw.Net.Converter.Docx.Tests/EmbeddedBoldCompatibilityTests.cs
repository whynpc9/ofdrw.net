using System.Xml.Linq;
using System.Globalization;
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
    public async Task RegularEmbeddedFaceHasTruthfulFlagsAndExplicitSemanticBoldStroke(bool explicitNative)
    {
        using var fonts = new PinnedCjkFontDirectory();
        var options = Options(fonts, explicitNative);
        using var source = File.OpenRead(Sample()); using var output = new MemoryStream();
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(source, output);
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
                Assert.Equal(text.FontSizeMillimeters * .01, double.Parse(xml.Attribute("LineWidth")!.Value, CultureInfo.InvariantCulture), precision: 3);
                var color = Assert.Single(xml.Elements(), e => e.Name.LocalName == "StrokeColor");
                Assert.Equal($"{text.FillColor.Red} {text.FillColor.Green} {text.FillColor.Blue}", color.Attribute("Value")?.Value);
            }
            else Assert.Null(xml.Attribute("Stroke"));
        }
        // Rewriting imported XML keeps its single text object/stroke; no second application.
        using var rewritten = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, rewritten);
        rewritten.Position = 0; var reread = await new OfdReader().ReadAsync(rewritten);
        Assert.Equal(texts.Select(text => text.SourceXml), reread.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(text => text.SourceXml));
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
        var emphasizedClone = Assert.IsType<OfdTextElement>(OfdModelCloner.CloneElement(text));
        Assert.True(emphasizedClone.EmitSyntheticBoldStroke);
        package.Pages[0].Elements[0] = emphasizedClone;
        using var clonedOutput = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, clonedOutput);
        clonedOutput.Position = 0; var clonedRead = await new OfdReader().ReadAsync(clonedOutput);
        var clonedXml = XElement.Parse(Assert.Single(clonedRead.Pages[0].Elements.OfType<OfdTextElement>()).SourceXml!);
        Assert.Equal("true", clonedXml.Attribute("Stroke")?.Value);
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
    public async Task FreshNameOnlyTextKeepsDefaultEmphasisAndIsStableOnImport(bool italic)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "SimSun" });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100, Elements = { new OfdTextElement
        {
            FontName = "SimSun", FontResourceId = "10", Text = "回A", Weight = 700,
            Italic = italic, FontSizeMillimeters = 4, FillColor = new OfdColor(10, 20, 30, 128)
        } } });
        string? first = null;
        for (var pass = 0; pass < 2; pass++)
        {
            using var output = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(package, output);
            output.Position = 0;
            package = await new OfdReader().ReadAsync(output);
            var text = Assert.Single(package.Pages.Single().Elements.OfType<OfdTextElement>());
            Assert.Equal("回A", text.Text);
            var xml = XElement.Parse(text.SourceXml!);
            Assert.Equal("true", xml.Attribute("Stroke")?.Value);
            Assert.Equal("true", xml.Attribute("Fill")?.Value);
            Assert.Equal("0.18", xml.Attribute("LineWidth")?.Value);
            var stroke = Assert.Single(xml.Elements(), node => node.Name.LocalName == "StrokeColor");
            Assert.Equal("10 20 30", stroke.Attribute("Value")?.Value);
            Assert.Equal("128", stroke.Attribute("Alpha")?.Value);
            Assert.Equal(italic, xml.Attribute("CTM") is not null);
            if (pass == 0) first = text.SourceXml;
            else Assert.Equal(first, text.SourceXml);
        }
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
