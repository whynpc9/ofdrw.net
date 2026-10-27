using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[Collection("Font registry")]
public sealed class ExplicitGlyphTransformTests
{
    private static byte[] Font(string fileName = "NotoSans-Regular.ttf")
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return File.ReadAllBytes(Path.Combine(root!.FullName,
            "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts", fileName));
    }

    private static OfdTextElement Text(int firstGid, int secondGid, string? deltas = "8", string logicalText = "AB")
    {
        var deltaAttribute = deltas is null ? "" : $" DeltaX='{deltas}'";
        return new OfdTextElement
        {
            FontResourceId = "10", FontName = "Noto Sans", Text = logicalText, FontSizeMillimeters = 6,
            XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 50, HeightMillimeters = 12,
            SourceXml = $"<TextObject ID='2' Boundary='10 10 50 12' Font='10' Size='6'>" +
                $"<CGTransform CodePosition='0' CodeCount='2' GlyphCount='2'><Glyphs>{firstGid} {secondGid}</Glyphs></CGTransform>" +
                $"<TextCode X='0' Y='6'{deltaAttribute}>{logicalText}</TextCode></TextObject>",
            Runs = { new OfdTextRun { Text = logicalText, XMillimeters = 0, YMillimeters = 6, DeltaX = deltas } }
        };
    }

    private static OfdTextElement SourceSpanText(string sourceSpan, int gid, string? deltaX = null)
    {
        var text = Text(gid, gid, deltas: deltaX, logicalText: sourceSpan);
        var delta = deltaX is null ? "" : $" DeltaX='{deltaX}'";
        text.SourceXml = $"<TextObject ID='2' Boundary='10 10 50 12' Font='10' Size='6'>" +
            $"<CGTransform CodePosition='0' CodeCount='{sourceSpan.Length}' GlyphCount='1'><Glyphs>{gid}</Glyphs></CGTransform>" +
            $"<TextCode X='0' Y='6'{delta}>{sourceSpan}</TextCode></TextObject>";
        return text;
    }

    private static OfdTextElement ThaiSequenceText(string glyphs = "59 86", string? deltaX = null)
    {
        const string source = "\u0E33";
        var text = Text(59, 86, deltas: deltaX, logicalText: source);
        text.FontName = "Noto Sans Thai";
        var delta = deltaX is null ? "" : $" DeltaX='{deltaX}'";
        text.SourceXml = "<TextObject ID='2' Boundary='10 10 50 12' Font='10' Size='6'>" +
            $"<CGTransform CodePosition='0' CodeCount='1' GlyphCount='2'><Glyphs>{glyphs}</Glyphs></CGTransform>" +
            $"<TextCode X='0' Y='6'{delta}>{source}</TextCode></TextObject>";
        return text;
    }

    private static OfdTextElement PositionedSequenceText(string source, string glyphs,
        string? deltaX, string? deltaY = null)
    {
        var text = Text(1, 2, deltaX, source);
        text.FontName = source == "போ" ? "Noto Sans Tamil" : "Noto Sans";
        text.Runs[0].DeltaY = deltaY;
        var count = glyphs.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
        var x = deltaX is null ? "" : $" DeltaX='{deltaX}'";
        var y = deltaY is null ? "" : $" DeltaY='{deltaY}'";
        text.SourceXml = "<TextObject ID='2' Boundary='10 10 50 12' Font='10' Size='6'>" +
            $"<CGTransform CodePosition='0' CodeCount='{source.Length}' GlyphCount='{count}'><Glyphs>{glyphs}</Glyphs></CGTransform>" +
            $"<TextCode X='0' Y='6'{x}{y}>{source}</TextCode></TextObject>";
        return text;
    }

    [Theory]
    [InlineData("போ", "47 28 41", "4.32 4.668", null, "NotoSansTamil-Regular.ttf", "Noto Sans Tamil")]
    [InlineData("A", "37 38 39", "2 3", "1 -2", "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("AB", "37 38 39 40", "2 3 4", null, "NotoSans-Regular.ttf", "Noto Sans")]
    public async Task CompletePositionedSourceSpanKeepsOriginalGidsAndOneLogicalText(
        string sourceText, string glyphs, string deltas, string? deltaY, string fontFile, string fontName)
    {
        var ids = glyphs.Split(' ').Select(int.Parse).ToArray();
        var bytes = Font(fontFile);
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = fontName, Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { PositionedSequenceText(sourceText, glyphs, deltas, deltaY) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();
        var readback = await new OfdReader().ReadAsync(new MemoryStream(input));
        var original = Assert.Single(readback.Pages[0].Elements.OfType<OfdTextElement>());
        Assert.Equal(sourceText, original.Text);
        Assert.Contains($"GlyphCount=\"{ids.Length}\"", original.SourceXml);

        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg);
        var svgDocument = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var group = Assert.Single(svgDocument.Descendants(), node =>
            node.Attribute("data-ofd-text")?.Value == sourceText);
        var paths = group.Elements().Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(ids, paths.Select(path => (int)path.Attribute("data-original-gid")!).ToArray());
        var outlines = new TrueTypeGlyphOutlines(bytes);
        var positions = new double[ids.Length];
        var increments = deltas.Split(' ').Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        var heights = new double[ids.Length];
        var vertical = deltaY?.Split(' ').Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        for (var i = 1; i < ids.Length; i++) positions[i] = positions[i - 1] + increments[i - 1];
        if (vertical is not null)
            for (var i = 1; i < ids.Length; i++) heights[i] = heights[i - 1] + vertical[i - 1];
        for (var i = 0; i < ids.Length; i++)
        {
            var command = outlines.Read(ids[i]).ToPathCommands().First();
            Assert.StartsWith("M " + (positions[i] + command.X * 6d / outlines.UnitsPerEm)
                    .ToString("0.###", CultureInfo.InvariantCulture) + " " +
                (6 + heights[i] - command.Y * 6d / outlines.UnitsPerEm)
                    .ToString("0.###", CultureInfo.InvariantCulture),
                paths[i].Attribute("d")!.Value);
        }
        using var html = new MemoryStream();
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(input), html);
        Assert.Equal(ids.Length, XDocument.Parse(Encoding.UTF8.GetString(html.ToArray())).Descendants()
            .Count(node => node.Attribute("data-original-gid") is not null));
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal(sourceText, parsed.GetPage(1).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var content = Encoding.ASCII.GetString(inspection.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
        Assert.Equal(1, content.Split(" Tj", StringSplitOptions.None).Length - 1);
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } output)
        {
            Directory.CreateDirectory(output);
            var stem = sourceText == "போ" ? "tamil-two-to-three" : $"synthetic-{sourceText.Length}-to-{ids.Length}";
            await File.WriteAllBytesAsync(Path.Combine(output, stem + ".ofd"), input);
            await File.WriteAllBytesAsync(Path.Combine(output, stem + ".svg"), svg.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(output, stem + ".pdf"), pdf.ToArray());
        }
    }

    [Fact]
    public async Task PositionedTamilStrokePaintsOnlyOriginalOutlinesAndRetainsOneSemanticSpan()
    {
        var sourceText = "போ";
        var text = PositionedSequenceText(sourceText, "47 28 41", "4.32 4.668");
        text.SourceXml = text.SourceXml!.Replace("<TextObject ID='2'",
            "<TextObject Fill='false' Stroke='true' LineWidth='0.2' ID='2'")
            .Replace("<CGTransform", "<StrokeColor Value='220 20 30'/><CGTransform");
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans Tamil",
            Data = Font("NotoSansTamil-Regular.ttf") });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { text } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg);
        var paths = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants()
            .Where(node => node.Attribute("data-original-gid") is not null).ToArray();
        Assert.Equal(new[] { 47, 28, 41 }, paths
            .Select(node => (int)node.Attribute("data-original-gid")!).ToArray());
        Assert.All(paths, node => Assert.Equal("none", node.Attribute("fill")!.Value));
        Assert.All(paths, node => Assert.NotEqual("none", node.Attribute("stroke")!.Value));
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal(sourceText, parsed.GetPage(1).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var content = Encoding.ASCII.GetString(inspection.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
        var semantic = content.LastIndexOf(" Tj", StringComparison.Ordinal);
        Assert.True(semantic > 0);
        Assert.DoesNotContain(" c\n", content.Substring(semantic));
        Assert.DoesNotContain(" m\n", content.Substring(semantic));
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } output)
        {
            Directory.CreateDirectory(output);
            await File.WriteAllBytesAsync(Path.Combine(output, "tamil-two-to-three-stroke.ofd"), input);
            await File.WriteAllBytesAsync(Path.Combine(output, "tamil-two-to-three-stroke.svg"), svg.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(output, "tamil-two-to-three-stroke.pdf"), pdf.ToArray());
        }
    }

    [Fact]
    public async Task ThaiOneCharacterTwoOriginalGlyphOutlinesKeepOneLogicalCharacter()
    {
        var bytes = Font("NotoSansThai-Regular.ttf");
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans Thai", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { ThaiSequenceText() } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var source = ofd.ToArray();
        var reread = await new OfdReader().ReadAsync(new MemoryStream(source));
        Assert.Equal("\u0E33", Assert.Single(reread.Pages[0].Elements.OfType<OfdTextElement>()).Text);

        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(source), svg);
        var svgDocument = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var group = Assert.Single(svgDocument.Descendants(), node => node.Attribute("data-ofd-text")?.Value == "\u0E33");
        var paths = group.Elements().Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(new[] { 59, 86 }, paths.Select(path => (int)path.Attribute("data-original-gid")!).ToArray());
        var outlines = new TrueTypeGlyphOutlines(bytes);
        for (var index = 0; index < paths.Length; index++)
        {
            var first = outlines.Read(new[] { 59, 86 }[index]).ToPathCommands().First();
            var expectedStart = "M " + (first.X * 6d / outlines.UnitsPerEm).ToString("0.###", CultureInfo.InvariantCulture) +
                " " + (6 - first.Y * 6d / outlines.UnitsPerEm).ToString("0.###", CultureInfo.InvariantCulture);
            Assert.StartsWith(expectedStart, paths[index].Attribute("d")!.Value);
        }

        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(source), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("\u0E33", parsed.GetPage(1).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var content = Encoding.ASCII.GetString(inspection.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
        Assert.DoesNotContain("003B0056", content);
        Assert.Contains(" Tj", content);
        Assert.Contains(" c\n", content);
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } output)
        {
            Directory.CreateDirectory(output);
            await File.WriteAllBytesAsync(Path.Combine(output, "thai-sequence.ofd"), source);
            await File.WriteAllBytesAsync(Path.Combine(output, "thai-sequence.svg"), svg.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(output, "thai-sequence.pdf"), pdf.ToArray());
        }
    }

    [Theory]
    [InlineData("59", null)]
    [InlineData("59 59", null)]
    [InlineData("59 99999", null)]
    public async Task InvalidOrPositionedThaiSequenceRefusesBothExportsWithoutOutput(string glyphs, string? deltaX)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans Thai",
            Data = Font("NotoSansThai-Regular.ttf") });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { ThaiSequenceText(glyphs, deltaX) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf));
        Assert.Equal(0, pdf.Length);
        using var svg = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), svg));
        Assert.Equal(0, svg.Length);
    }

    [Theory]
    [InlineData(null, null)]
    [InlineData("4.32", null)]
    [InlineData("4.32 4.668 1", null)]
    [InlineData("4.32 4.668", "1")]
    [InlineData("NaN 4.668", null)]
    [InlineData("1e308 1e308", null)]
    public async Task IncompletePositionedSequenceRefusesBothExportsWithoutOutput(
        string? deltaX, string? deltaY)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans Tamil",
            Data = Font("NotoSansTamil-Regular.ttf") });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { PositionedSequenceText("போ", "47 28 41", deltaX, deltaY) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();
        using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf));
        Assert.Equal(0, pdf.Length);
        using var svg = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg));
        Assert.Equal(0, svg.Length);
    }

    [Fact]
    public async Task PositionedSpanWithDistinctSourceCharactersSharingSemanticGidRefusesPdf()
    {
        var face = new OpenTypeFace(Font());
        var shared = new OpenTypeCmap(face).Glyph('A');
        Assert.True(shared > 0);
        face.Tables["cmap"] = OpenTypeCmap.Build(new System.Collections.Generic.Dictionary<int, int>
        {
            ['A'] = shared, ['B'] = shared
        });
        var text = PositionedSequenceText("AB", "37 38 39", "8 8");
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = face.Build() });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { text } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();
        var saved = await new OfdReader().ReadAsync(new MemoryStream(input));
        Assert.Equal("AB", saved.Pages[0].Elements.OfType<OfdTextElement>().Single().Text);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg);
        Assert.Equal(new[] { 37, 38, 39 }, XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()))
            .Descendants().Where(node => node.Attribute("data-original-gid") is not null)
            .Select(node => (int)node.Attribute("data-original-gid")!).ToArray());
        using var pdf = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf));
        Assert.Contains("share a semantic-font GID", error.Message);
        Assert.Equal(0, pdf.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PositionedSpanRejectsCmapAliasUsedByAnotherTextObject(bool secondExplicit)
    {
        var face = new OpenTypeFace(Font());
        var shared = new OpenTypeCmap(face).Glyph('A');
        face.Tables["cmap"] = OpenTypeCmap.Build(new System.Collections.Generic.Dictionary<int, int>
        {
            ['A'] = shared, ['B'] = shared
        });
        var second = secondExplicit ? PositionedSequenceText("B", "37 38 39", "8 8") :
            new OfdTextElement
            {
                FontResourceId = "10", FontName = "Noto Sans", Text = "B", FontSizeMillimeters = 6,
                XMillimeters = 10, YMillimeters = 25, WidthMillimeters = 50, HeightMillimeters = 12,
                Runs = { new OfdTextRun { Text = "B", XMillimeters = 0, YMillimeters = 6 } }
            };
        if (secondExplicit) second.SourceXml = second.SourceXml!.Replace("ID='2'", "ID='3'");
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = face.Build() });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { PositionedSequenceText("A", "37 38 39", "8 8"), second } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();
        using var pdf = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf));
        Assert.Contains("share a semantic-font GID", error.Message);
        Assert.Equal(0, pdf.Length);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg);
        var sourceGroup = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants()
            .Single(node => node.Attribute("data-ofd-text")?.Value == "A");
        Assert.Equal(new[] { 37, 38, 39 }, sourceGroup.Elements()
            .Where(node => node.Attribute("data-original-gid") is not null)
            .Select(node => (int)node.Attribute("data-original-gid")!).ToArray());
    }

    [Fact]
    public async Task TwoGlyphProfileUsesSourceGidsEvenWhenFirstGlyphHasNaturalAdvance()
    {
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        var first = cmap.Glyph('B'); var second = cmap.Glyph('C');
        var outlines = new TrueTypeGlyphOutlines(bytes);
        Assert.True(outlines.Read(first).AdvanceWidth > 0);
        var text = Text(first, second, deltas: null, logicalText: "A");
        text.SourceXml = "<TextObject ID='2' Boundary='10 10 50 12' Font='10' Size='6'>" +
            $"<CGTransform CodePosition='0' CodeCount='1' GlyphCount='2'><Glyphs>{first} {second}</Glyphs></CGTransform>" +
            "<TextCode X='0' Y='6'>A</TextCode></TextObject>";
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { text } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), svg);
        var document = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var group = Assert.Single(document.Descendants(), node => node.Attribute("data-ofd-text")?.Value == "A");
        var paths = group.Elements().Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(new[] { first, second }, paths
            .Select(node => (int)node.Attribute("data-original-gid")!).ToArray());
        for (var index = 0; index < paths.Length; index++)
        {
            var point = outlines.Read(new[] { first, second }[index]).ToPathCommands().First();
            Assert.StartsWith("M " + (point.X * 6d / outlines.UnitsPerEm).ToString("0.###", CultureInfo.InvariantCulture) +
                " " + (6 - point.Y * 6d / outlines.UnitsPerEm).ToString("0.###", CultureInfo.InvariantCulture),
                paths[index].Attribute("d")!.Value);
        }
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("A", parsed.GetPage(1).Text);
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } output)
        {
            Directory.CreateDirectory(output);
            await File.WriteAllBytesAsync(Path.Combine(output, "nonzero-first-sequence.ofd"), ofd.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(output, "nonzero-first-sequence.svg"), svg.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(output, "nonzero-first-sequence.pdf"), pdf.ToArray());
        }
    }

    [Fact]
    public async Task TwoGlyphPdfWithoutSourceUnicodeCoverageRefusesWhileSvgKeepsOriginalOutlines()
    {
        var text = ThaiSequenceText();
        text.Text = "A";
        text.Runs[0].Text = "A";
        text.SourceXml = text.SourceXml!.Replace("\u0E33", "A");
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans Thai",
            Data = Font("NotoSansThai-Regular.ttf") });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { text } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), svg);
        var paths = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()))
            .Descendants().Where(node => node.Attribute("data-original-gid") is not null).ToArray();
        Assert.Equal(new[] { 59, 86 }, paths.Select(path => (int)path.Attribute("data-original-gid")!).ToArray());
        using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf));
        Assert.Equal(0, pdf.Length);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task TwoGlyphStrokeDrawsOnlySourceGidOutlinesAndOneInvisibleSemanticGlyph(bool fill)
    {
        var text = ThaiSequenceText();
        text.SourceXml = text.SourceXml!.Replace("<TextObject ID='2'",
            $"<TextObject Fill='{fill.ToString().ToLowerInvariant()}' Stroke='true' LineWidth='0.2' ID='2'")
            .Replace("<CGTransform", "<StrokeColor Value='220 20 30'/><CGTransform");
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans Thai",
            Data = Font("NotoSansThai-Regular.ttf") });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { text } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), svg);
        var paths = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()))
            .Descendants().Where(node => node.Attribute("data-original-gid") is not null).ToArray();
        Assert.Equal(new[] { 59, 86 }, paths.Select(node => (int)node.Attribute("data-original-gid")!).ToArray());
        Assert.All(paths, node => Assert.NotEqual("none", node.Attribute("stroke")!.Value));
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("\u0E33", parsed.GetPage(1).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var content = Encoding.ASCII.GetString(inspection.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
        var semantic = content.LastIndexOf(" Tj", StringComparison.Ordinal);
        Assert.True(semantic > 0);
        Assert.DoesNotContain(" c\n", content.Substring(semantic));
        Assert.DoesNotContain(" m\n", content.Substring(semantic));
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } output)
        {
            Directory.CreateDirectory(output);
            var stem = fill ? "thai-fill-stroke" : "thai-stroke-only";
            await File.WriteAllBytesAsync(Path.Combine(output, stem + ".ofd"), ofd.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(output, stem + ".svg"), svg.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(output, stem + ".pdf"), pdf.ToArray());
        }
    }

    [Fact]
    public async Task TwoTextCodesKeepGlobalGlyphPositionsAndIndependentStarts()
    {
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        var gids = new[] { 'B', 'A', 'D', 'C' }.Select(value => cmap.Glyph(value)).ToArray();
        var text = Text(gids[0], gids[1]);
        text.Text = "ABCD";
        text.HeightMillimeters = 30;
        text.SourceXml = $"<TextObject ID='2' Boundary='10 10 50 30' Font='10' Size='6'>" +
            $"<CGTransform CodePosition='0' CodeCount='2' GlyphCount='2'><Glyphs>{gids[0]} {gids[1]}</Glyphs></CGTransform>" +
            $"<CGTransform CodePosition='2' CodeCount='2' GlyphCount='2'><Glyphs>{gids[2]} {gids[3]}</Glyphs></CGTransform>" +
            "<TextCode X='0' Y='6' DeltaX='8'>AB</TextCode>" +
            "<TextCode X='0' Y='14' DeltaX='8'>CD</TextCode></TextObject>";
        text.Runs.Add(new OfdTextRun { Text = "CD", XMillimeters = 0, YMillimeters = 14, DeltaX = "8" });
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { text } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var source = ofd.ToArray();
        var reread = await new OfdReader().ReadAsync(new MemoryStream(source));
        var actual = Assert.Single(reread.Pages[0].Elements.OfType<OfdTextElement>());
        Assert.Equal("ABCD", actual.Text);
        Assert.Equal(2, actual.Runs.Count);
        Assert.Contains("CodePosition=\"2\"", actual.SourceXml);

        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(source), svg);
        var document = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var group = Assert.Single(document.Descendants(), node => node.Attribute("data-ofd-text")?.Value == "ABCD");
        var paths = group.Elements().Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(gids, paths.Select(path => (int)path.Attribute("data-original-gid")!).ToArray());
        Assert.NotEqual(paths[0].Attribute("d")!.Value, paths[2].Attribute("d")!.Value);

        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(source), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        var letters = parsed.GetPage(1).Letters.ToArray();
        Assert.Equal("ABCD", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(8 * 72d / 25.4, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 1);
        Assert.True(letters[2].StartBaseLine.Y < letters[0].StartBaseLine.Y);
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } artifactDirectory)
        {
            Directory.CreateDirectory(artifactDirectory);
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, "multicode-global.ofd"), source);
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, "multicode-global.svg"), svg.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, "multicode-global.pdf"), pdf.ToArray());
        }
    }

    [Theory]
    [InlineData("fl", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("ffi", 1654, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("f\u115Fl", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("f\u1160l", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("f\u3164l", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("f\uFFA0l", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("\u115F\u115F", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("\u1160\u1160", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("\u3164\u3164", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("\uFFA0\uFFA0", 1653, "NotoSans-Regular.ttf", "Noto Sans")]
    [InlineData("fl", 2135, "NotoSerif-Regular.ttf", "Noto Serif")]
    public async Task SourceSpanUsesOneOriginalGidAndExtractsAllSourceCharacters(
        string sourceSpan, int ligatureGid, string fontFile, string fontName)
    {
        var bytes = Font(fontFile);
        var text = SourceSpanText(sourceSpan, ligatureGid);
        text.FontName = fontName;
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = fontName, Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { text } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var source = ofd.ToArray();
        var reread = await new OfdReader().ReadAsync(new MemoryStream(source));
        var actualText = Assert.Single(reread.Pages[0].Elements.OfType<OfdTextElement>());
        Assert.Equal(sourceSpan, actualText.Text);
        Assert.Contains("GlyphCount=\"1\"", actualText.SourceXml);

        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(source), svg);
        var svgDocument = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var group = Assert.Single(svgDocument.Descendants(), node => node.Attribute("data-ofd-text")?.Value == sourceSpan);
        var path = Assert.Single(group.Elements().Where(node => node.Name.LocalName == "path"));
        Assert.Equal(ligatureGid, (int)path.Attribute("data-original-gid")!);
        var firstCommand = new TrueTypeGlyphOutlines(bytes).Read(ligatureGid).ToPathCommands().First();
        var firstScale = 6d / new TrueTypeGlyphOutlines(bytes).UnitsPerEm;
        var expectedStart = "M " + (firstCommand.X * firstScale).ToString("0.###", CultureInfo.InvariantCulture) +
            " " + (6 - firstCommand.Y * firstScale).ToString("0.###", CultureInfo.InvariantCulture);
        Assert.StartsWith(expectedStart, path.Attribute("d")!.Value);
        Assert.DoesNotContain(group.Elements(), node => node.Name.LocalName == "text");

        using var html = new MemoryStream();
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(source), html);
        var htmlDocument = XDocument.Parse(Encoding.UTF8.GetString(html.ToArray()));
        var htmlGroup = Assert.Single(htmlDocument.Descendants(),
            node => node.Attribute("data-ofd-text")?.Value == sourceSpan);
        Assert.Single(htmlGroup.Elements().Where(node => node.Name.LocalName == "path"));

        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(source), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal(sourceSpan, parsed.GetPage(1).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var pageContent = Encoding.ASCII.GetString(inspection.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
        Assert.Contains($"<{ligatureGid:X4}> Tj", pageContent);
        var unicode = inspection.Internals.GetAllObjects().OfType<PdfSharpCore.Pdf.PdfDictionary>()
            .Where(item => item.Elements.GetName("/Subtype") == "/Type0")
            .Select(item => item.Elements.GetDictionary("/ToUnicode"))
            .Where(item => item?.Stream is not null)
            .Select(item => Encoding.ASCII.GetString(item!.Stream!.Value)).ToArray();
        var sourceHex = BitConverter.ToString(Encoding.BigEndianUnicode.GetBytes(sourceSpan)).Replace("-", string.Empty);
        Assert.Contains(unicode, map => map.Contains($"<{ligatureGid:X4}><{sourceHex}>"));
        var embedded = Assert.Single(inspection.Internals.GetAllObjects().OfType<PdfSharpCore.Pdf.PdfDictionary>()
            .Where(item => item.Elements.ContainsKey("/FontFile2")))
            .Elements.GetDictionary("/FontFile2").Stream.UnfilteredValue;
        var expectedOutline = new TrueTypeGlyphOutlines(bytes).Read(ligatureGid).ToPathCommands()
            .Select(command => (command.Verb, command.X, command.Y, command.ControlX, command.ControlY)).ToArray();
        var embeddedOutline = new TrueTypeGlyphOutlines(embedded).Read(ligatureGid).ToPathCommands()
            .Select(command => (command.Verb, command.X, command.Y, command.ControlX, command.ControlY)).ToArray();
        Assert.Equal(expectedOutline, embeddedOutline);
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } artifactDirectory)
        {
            Directory.CreateDirectory(artifactDirectory);
            var stem = $"single-glyph-{Path.GetFileNameWithoutExtension(fontFile)}-{sourceSpan}";
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, stem + ".ofd"), source);
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, stem + ".svg"), svg.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, stem + ".html"), html.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, stem + ".pdf"), pdf.ToArray());
        }
    }

    [Fact]
    public async Task PositionedMultiCodeClusterRefusesBothVisualExports()
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = Font() });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { SourceSpanText("fl", 1653, deltaX: "8") } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf));
        Assert.Equal(0, pdf.Length);
        using var svg = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), svg));
        Assert.Equal(0, svg.Length);
    }

    [Fact]
    public async Task OnePdfAliasCannotExtractTwoDifferentSourceSpansFromTheSameGid()
    {
        var second = SourceSpanText("fi", 1653);
        second.SourceXml = second.SourceXml!.Replace("ID='2'", "ID='3'");
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = Font() });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { SourceSpanText("fl", 1653) } });
        package.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { second } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf));
        Assert.Equal(0, pdf.Length);
    }

    [Fact]
    public async Task ExplicitJamoGlyphSlotsKeepTheirSpecifiedPdfAdvance()
    {
        const string logical = "\u1100\u1161";
        Assert.Single(OfdTextGeometry.Glyphs(logical));
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        // The logical Jamo map to independent original outlines, even when
        // the source font has no Unicode cmap entries for the Jamo themselves.
        var firstGid = cmap.Glyph('B');
        var secondGid = cmap.Glyph('A');
        Assert.True(firstGid > 0 && secondGid > 0 && firstGid != secondGid);
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(firstGid, secondGid, logicalText: logical) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        var letters = parsed.GetPage(1).Letters;
        Assert.Equal(2, letters.Count);
        Assert.Equal(8 * 72d / 25.4, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 1);
    }

    [Fact]
    public async Task CompositeOriginalGidKeepsItsCompletePdfOutline()
    {
        const int compositeGid = 171;
        var bytes = Font();
        var originalOutline = new TrueTypeGlyphOutlines(bytes).Read(compositeGid).ToPathCommands()
            .Select(command => (command.Verb, command.X, command.Y, command.ControlX, command.ControlY)).ToArray();
        Assert.NotEmpty(originalOutline);
        var a = new OpenTypeCmap(new OpenTypeFace(bytes)).Glyph('A');
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(compositeGid, a) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("AB", parsed.GetPage(1).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var embedded = Assert.Single(inspection.Internals.GetAllObjects().OfType<PdfSharpCore.Pdf.PdfDictionary>()
            .Where(item => item.Elements.ContainsKey("/FontFile2")))
            .Elements.GetDictionary("/FontFile2").Stream.UnfilteredValue;
        var embeddedOutline = new TrueTypeGlyphOutlines(embedded).Read(compositeGid).ToPathCommands()
            .Select(command => (command.Verb, command.X, command.Y, command.ControlX, command.ControlY)).ToArray();
        Assert.Equal(originalOutline, embeddedOutline);
    }

    [Fact]
    public async Task CompatibleExplicitMappingsShareOnePdfFontAcrossPages()
    {
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        var second = Text(cmap.Glyph('D'), cmap.Glyph('C'), logicalText: "CD");
        second.SourceXml = second.SourceXml!.Replace("ID='2'", "ID='3'");
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(cmap.Glyph('B'), cmap.Glyph('A')) } });
        package.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { second } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("AB", parsed.GetPage(1).Text);
        Assert.Equal("CD", parsed.GetPage(2).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        Assert.Single(inspection.Internals.GetAllObjects().OfType<PdfSharpCore.Pdf.PdfDictionary>()
            .Where(item => item.Elements.ContainsKey("/FontFile2")));
    }

    [Fact]
    public async Task ConflictingExplicitMappingsUseSeparatePdfAliases()
    {
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        var a = cmap.Glyph('A'); var b = cmap.Glyph('B');
        var second = Text(a, b);
        second.SourceXml = second.SourceXml!.Replace("ID='2'", "ID='3'");
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(b, a) } });
        package.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { second } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd.ToArray()), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("AB", parsed.GetPage(1).Text);
        Assert.Equal("AB", parsed.GetPage(2).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var mappings = inspection.Internals.GetAllObjects().OfType<PdfSharpCore.Pdf.PdfDictionary>()
            .Where(item => item.Elements.ContainsKey("/FontFile2"))
            .Select(item => new OpenTypeCmap(new OpenTypeFace(
                item.Elements.GetDictionary("/FontFile2").Stream.UnfilteredValue)))
            .Select(face => (face.Glyph('A'), face.Glyph('B'))).ToArray();
        Assert.Equal(2, mappings.Length);
        Assert.Contains((b, a), mappings);
        Assert.Contains((a, b), mappings);
    }

    [Fact]
    public async Task OneToOneOriginalGidsPaintDifferentOutlinesWhilePdfKeepsLogicalText()
    {
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        var a = cmap.Glyph('A'); var b = cmap.Glyph('B');
        Assert.True(a > 0 && b > 0 && a != b);
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(b, a) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var saved = ofd.ToArray();
        var reread = await new OfdReader().ReadAsync(new MemoryStream(saved));
        Assert.Equal("AB", reread.Pages[0].Elements.OfType<OfdTextElement>().Single().Text);
        Assert.Contains("CGTransform", reread.Pages[0].Elements.OfType<OfdTextElement>().Single().SourceXml);

        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(saved), svg);
        var doc = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var group = Assert.Single(doc.Descendants(), node => node.Attribute("data-ofd-text")?.Value == "AB");
        var paths = group.Elements().Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(new[] { b, a }, paths.Select(path => (int)path.Attribute("data-original-gid")!).ToArray());
        var firstCommand = new TrueTypeGlyphOutlines(bytes).Read(b).ToPathCommands().First();
        var firstScale = 6d / new TrueTypeGlyphOutlines(bytes).UnitsPerEm;
        var expectedStart = "M " + (firstCommand.X * firstScale).ToString("0.###", CultureInfo.InvariantCulture) +
            " " + (6 - firstCommand.Y * firstScale).ToString("0.###", CultureInfo.InvariantCulture);
        Assert.StartsWith(expectedStart, paths[0].Attribute("d")!.Value);
        Assert.All(paths, path => Assert.NotEmpty(path.Attribute("d")!.Value));
        Assert.DoesNotContain(group.Elements(), node => node.Name.LocalName == "text");

        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(saved), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("AB", parsed.GetPage(1).Text);
        using var inspection = PdfSharpCore.Pdf.IO.PdfReader.Open(new MemoryStream(pdf.ToArray()),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var embedded = inspection.Internals.GetAllObjects().OfType<PdfSharpCore.Pdf.PdfDictionary>()
            .Where(item => item.Elements.ContainsKey("/FontFile2"))
            .Select(item => item.Elements.GetDictionary("/FontFile2").Stream.UnfilteredValue)
            .ToArray();
        var pdfOutlines = new TrueTypeGlyphOutlines(Assert.Single(embedded));
        var sourceOutlines = new TrueTypeGlyphOutlines(bytes);
        foreach (var gid in new[] { b, a })
        {
            var expected = sourceOutlines.Read(gid).ToPathCommands()
                .Select(command => (command.Verb, command.X, command.Y, command.ControlX, command.ControlY)).ToArray();
            var actual = pdfOutlines.Read(gid).ToPathCommands()
                .Select(command => (command.Verb, command.X, command.Y, command.ControlX, command.ControlY)).ToArray();
            Assert.Equal(expected, actual);
        }
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } artifactDirectory)
        {
            Directory.CreateDirectory(artifactDirectory);
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, "explicit-gid-synthetic.ofd"), saved);
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, "explicit-gid-synthetic.svg"), svg.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, "explicit-gid-synthetic.pdf"), pdf.ToArray());
        }
    }

    [Theory]
    [InlineData(@"\0041B", "AB")]
    [InlineData(@"A\005C", "A\\")]
    public async Task EscapedSourceTextKeepsOriginalGidsAndDecodedLogicalText(string sourceText, string logicalText)
    {
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        var firstGid = cmap.Glyph('B');
        var secondGid = cmap.Glyph('A');
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(firstGid, secondGid) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        using (var archive = new ZipArchive(ofd, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.Entries.Where(item => item.FullName.EndsWith(".xml", StringComparison.Ordinal))
                .Single(item => { using var stream = item.Open(); return XDocument.Load(stream).Descendants()
                    .Any(node => node.Name.LocalName == "TextObject"); });
            XDocument xml;
            using (var stream = entry.Open()) xml = XDocument.Load(stream);
            xml.Descendants().Single(node => node.Name.LocalName == "TextCode").Value = sourceText;
            var name = entry.FullName;
            entry.Delete();
            using var replacement = archive.CreateEntry(name).Open();
            xml.Save(replacement);
        }
        var input = ofd.ToArray();
        var readback = await new OfdReader().ReadAsync(new MemoryStream(input));
        var original = Assert.IsType<OfdTextElement>(Assert.Single(readback.Pages[0].Elements));
        Assert.Equal(sourceText, original.Text);
        Assert.Contains(sourceText, original.SourceXml);

        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg);
        var group = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants()
            .Single(node => node.Attribute("data-ofd-text")?.Value == logicalText);
        Assert.Equal(new[] { firstGid, secondGid }, group.Elements()
            .Where(node => node.Name.LocalName == "path")
            .Select(node => (int)node.Attribute("data-original-gid")!).ToArray());
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal(logicalText, parsed.GetPage(1).Text);
    }

    [Fact]
    public async Task HtmlPreviewReservesOriginalGidOutlinesBeforeBuildingPaths()
    {
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(cmap.Glyph('B'), cmap.Glyph('A')) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();

        using var accepted = new MemoryStream();
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(input), accepted);
        var paths = XDocument.Parse(Encoding.UTF8.GetString(accepted.ToArray())).Descendants()
            .Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(2, paths.Length);

        using var rejected = new MemoryStream();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdToHtmlConverter(new OfdHtmlOptions { MaxGeneratedCharacters = 200_000 })
                .ConvertAsync(new MemoryStream(input), rejected));
        Assert.Contains("characters", error.Message);
        Assert.Equal(0, rejected.Length);
    }

    [Fact]
    public void IncompleteAmbiguousAndInvalidPositioningMappingsFailClosed()
    {
        Assert.Throws<NotSupportedException>(() => OfdExplicitGlyphRun.Read(Text(1, 2, "3 4")));
        Assert.Throws<NotSupportedException>(() => OfdExplicitGlyphRun.Read(Text(1, 1)));
        var missing = Text(1, 2);
        missing.SourceXml = missing.SourceXml!.Replace("CodeCount='2' GlyphCount='2'", "CodeCount='1' GlyphCount='1'")
            .Replace("<Glyphs>1 2</Glyphs>", "<Glyphs>1</Glyphs>");
        Assert.Throws<NotSupportedException>(() => OfdExplicitGlyphRun.Read(missing));
        var unsupported = Text(1, 2);
        unsupported.SourceXml = unsupported.SourceXml!.Replace("CodeCount='2' GlyphCount='2'", "CodeCount='2' GlyphCount='1'")
            .Replace("<Glyphs>1 2</Glyphs>", "<Glyphs>1</Glyphs>");
        Assert.Throws<NotSupportedException>(() => OfdExplicitGlyphRun.Read(unsupported));
    }

    [Fact]
    public async Task MissingOriginalGidLeavesBothOutputsEmpty()
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = Font() });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(ushort.MaxValue, 2) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();
        using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf));
        Assert.Equal(0, pdf.Length);
        using var svg = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg));
        Assert.Equal(0, svg.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HiddenUnsupportedOriginalGidDoesNotBlockVisibleText(bool hidden)
    {
        var invalid = Text(ushort.MaxValue, 36);
        if (hidden)
            invalid.SourceXml = invalid.SourceXml!.Replace("ID='2'", "ID='2' Visible=' false '");
        var visible = new OfdTextElement
        {
            FontResourceId = "10", FontName = "Noto Sans", Text = "OK", FontSizeMillimeters = 6,
            XMillimeters = 10, YMillimeters = 25, WidthMillimeters = 50, HeightMillimeters = 12,
            Runs = { new OfdTextRun { Text = "OK", XMillimeters = 0, YMillimeters = 6 } }
        };
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = Font() });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { invalid, visible } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();
        var readback = await new OfdReader().ReadAsync(new MemoryStream(input));
        var original = readback.Pages[0].Elements.OfType<OfdTextElement>().Single(text => text.Text == "AB");
        Assert.Contains("<Glyphs>65535 36</Glyphs>", original.SourceXml);
        Assert.Equal(hidden, XElement.Parse(original.SourceXml!).Attribute("Visible")?.Value.Trim() == "false");

        using var pdf = new MemoryStream();
        using var svg = new MemoryStream();
        if (hidden)
        {
            await new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf);
            await new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg);
            using var parsed = PdfPigDocument.Open(pdf.ToArray());
            Assert.Equal("OK", parsed.GetPage(1).Text);
            var markup = Encoding.UTF8.GetString(svg.ToArray());
            Assert.Contains("OK", markup);
            Assert.DoesNotContain("data-ofd-text=\"AB\"", markup);
        }
        else
        {
            await Assert.ThrowsAsync<NotSupportedException>(() =>
                new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf));
            await Assert.ThrowsAsync<NotSupportedException>(() =>
                new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg));
            Assert.Equal(0, pdf.Length);
            Assert.Equal(0, svg.Length);
        }
    }

    [Fact]
    public async Task GidSourceDoesNotNeedAnOriginalUnicodeCmap()
    {
        var face = new OpenTypeFace(Font());
        var originalCmap = new OpenTypeCmap(face);
        var a = originalCmap.Glyph('A'); var b = originalCmap.Glyph('B');
        Assert.True(face.Tables.Remove("cmap"));
        var resource = new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = face.Build() };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource]);
        var family = context.Resolve(Text(b, a), out var selected);
        Assert.Same(resource, selected);
        var resolverFace = PdfSharpCore.Fonts.GlobalFontSettings.FontResolver.ResolveTypeface(family, false, false);
        var aliased = PdfSharpCore.Fonts.GlobalFontSettings.FontResolver.GetFont(resolverFace.FaceName);
        var mapped = new OpenTypeCmap(new OpenTypeFace(aliased));
        Assert.Equal(b, mapped.Glyph('A'));
        Assert.Equal(a, mapped.Glyph('B'));
        Assert.Throws<NotSupportedException>(() => context.Resolve(new OfdTextElement
            { FontResourceId = "10", Text = "AB" }, out _));

        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = Font() });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(b, a) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        // The writer's ordinary embedding plan requires a cmap. Replace the
        // persisted font bytes to model an external cmap-free GID source.
        using var foreign = new MemoryStream();
        using (var sourceZip = new ZipArchive(new MemoryStream(ofd.ToArray()), ZipArchiveMode.Read))
        using (var targetZip = new ZipArchive(foreign, ZipArchiveMode.Create, leaveOpen: true))
        {
            var replaced = 0;
            foreach (var sourceEntry in sourceZip.Entries)
            {
                var targetEntry = targetZip.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
                using var output = targetEntry.Open();
                if (sourceEntry.FullName.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase))
                { output.Write(resource.Data); replaced++; }
                else
                { using var input = sourceEntry.Open(); input.CopyTo(output); }
            }
            Assert.Equal(1, replaced);
        }
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(foreign.ToArray()), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("AB", parsed.GetPage(1).Text);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(foreign.ToArray()), svg);
        var paths = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()))
            .Descendants().Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(new[] { b, a }, paths.Select(path => (int)path.Attribute("data-original-gid")!).ToArray());
    }

    [Fact]
    public async Task MissingDeltaKeepsBothGlyphOriginsAtTheSamePoint()
    {
        var bytes = Font();
        var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
        var b = cmap.Glyph('B'); var a = cmap.Glyph('A');
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { Text(b, a, deltas: null) } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        var input = ofd.ToArray();
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        var letters = parsed.GetPage(1).Letters;
        Assert.Equal("AB", parsed.GetPage(1).Text);
        Assert.Equal(2, letters.Count);
        Assert.Equal(letters[0].StartBaseLine.X, letters[1].StartBaseLine.X, 3);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg);
        var document = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var paths = document.Descendants().Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(2, paths.Length);
        for (var index = 0; index < 2; index++)
        {
            var glyph = new TrueTypeGlyphOutlines(bytes).Read(index == 0 ? b : a);
            var move = glyph.ToPathCommands().First();
            var scale = 6d / glyph.UnitsPerEm;
            var expected = "M " + (move.X * scale).ToString("0.###", CultureInfo.InvariantCulture) +
                " " + (6 - move.Y * scale).ToString("0.###", CultureInfo.InvariantCulture);
            Assert.StartsWith(expected, paths[index].Attribute("d")!.Value);
        }
        if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            await File.WriteAllBytesAsync(Path.Combine(directory, "explicit-gid-overlap.ofd"), input);
            await File.WriteAllBytesAsync(Path.Combine(directory, "explicit-gid-overlap.pdf"), pdf.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(directory, "explicit-gid-overlap.svg"), svg.ToArray());
        }
    }
}
