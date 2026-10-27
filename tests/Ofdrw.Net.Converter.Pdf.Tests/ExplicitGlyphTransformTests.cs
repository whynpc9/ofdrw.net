using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf.Converters;
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
    private static byte[] Font()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return File.ReadAllBytes(Path.Combine(root!.FullName,
            "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));
    }

    private static OfdTextElement Text(int firstGid, int secondGid, string? deltas = "8")
    {
        var deltaAttribute = deltas is null ? "" : $" DeltaX='{deltas}'";
        return new OfdTextElement
        {
            FontResourceId = "10", FontName = "Noto Sans", Text = "AB", FontSizeMillimeters = 6,
            XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 50, HeightMillimeters = 12,
            SourceXml = $"<TextObject ID='2' Boundary='10 10 50 12' Font='10' Size='6'>" +
                $"<CGTransform CodePosition='0' CodeCount='2' GlyphCount='2'><Glyphs>{firstGid} {secondGid}</Glyphs></CGTransform>" +
                $"<TextCode X='0' Y='6'{deltaAttribute}>AB</TextCode></TextObject>",
            Runs = { new OfdTextRun { Text = "AB", XMillimeters = 0, YMillimeters = 6, DeltaX = deltas } }
        };
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
