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
