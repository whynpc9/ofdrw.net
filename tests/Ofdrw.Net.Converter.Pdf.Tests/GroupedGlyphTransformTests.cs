using System;
using System.Globalization;
using System.IO;
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

public sealed partial class ExplicitGlyphTransformTests
{
    private const OfdGlyphMappingStrategy Grouped = OfdGlyphMappingStrategy.GroupedLocal;

    private static OfdTextElement GroupedText(string first, string firstMaps, string? firstDelta,
        string second, string secondMaps, string? secondDelta, bool inheritedX = false)
    {
        var text = new OfdTextElement
        {
            FontResourceId = "10", FontName = "Noto Sans", Text = first + second,
            FontSizeMillimeters = 6, XMillimeters = 10, YMillimeters = 10,
            WidthMillimeters = 90, HeightMillimeters = 40
        };
        var root = new XElement("TextObject", new XAttribute("ID", "2"),
            new XAttribute("Boundary", "10 10 90 40"), new XAttribute("Font", "10"), new XAttribute("Size", "6"));
        foreach (var item in new[] { (first, firstMaps, firstDelta, 6d), (second, secondMaps, secondDelta, 22d) })
        {
            root.Add(XElement.Parse("<maps>" + item.Item2 + "</maps>").Elements());
            var code = new XElement("TextCode", new XAttribute("Y", item.Item4), item.Item1);
            if (!inheritedX || text.Runs.Count == 0) code.SetAttributeValue("X", "3");
            if (item.Item3 is not null) code.SetAttributeValue("DeltaX", item.Item3);
            root.Add(code);
            text.Runs.Add(new OfdTextRun { Text = item.Item1, DeltaX = item.Item3,
                XMillimeters = 3, YMillimeters = item.Item4 });
        }
        text.SourceXml = root.ToString(SaveOptions.DisableFormatting);
        return text;
    }

    private static async Task<byte[]> GroupedPackage(OfdTextElement text)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = Font() });
        package.Pages.Add(new OfdPage { WidthMillimeters = 120, HeightMillimeters = 65, Elements = { text } });
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        return output.ToArray();
    }

    private static Task ConvertGrouped(string outlet, byte[] source, Stream output) => outlet switch
    {
        "PDF" => new OfdToPdfConverter(new OfdToPdfOptions { GlyphMappingStrategy = Grouped })
            .ConvertAsync(new MemoryStream(source), output),
        "SVG" => new OfdToSvgConverter(Grouped).ConvertAsync(new MemoryStream(source), output),
        _ => new OfdToHtmlConverter(new OfdHtmlOptions { GlyphMappingStrategy = Grouped })
            .ConvertAsync(new MemoryStream(source), output)
    };

    [Fact]
    public void GroupedHebrewTextCodesCannotShareOneWholeWordSemanticCid()
    {
        var face = Font("DejaVuSans-R22.ttf");
        var cmap = new OpenTypeCmap(new OpenTypeFace(face));
        string Mapping(string source) => MapSpan(0, source.Length,
            source.Reverse().Select(value => cmap.Glyph(value)).ToArray());
        var text = GroupedText("של", Mapping("של"), "5", "ום", Mapping("ום"), "5");
        text.FontName = "DejaVu Sans";
        var resource = new OfdFontResource { Id = "10", FontName = text.FontName, Data = face };
        var projected = Assert.IsType<OfdExplicitGlyphRun>(OfdExplicitGlyphRun.Read(
            text, resource, strategy: Grouped));
        Assert.Equal("שלום", projected.LogicalText);
        Assert.Null(projected.VisualHebrewSourceSpan);
        Assert.Equal(new[] { "של", "ום" }, projected.Groups!.Select(group =>
            group.Projection.VisualHebrewSourceSpan));
    }

    [Theory]
    [InlineData("second-only", "AB", "CD", "36 37 40 39", "3 11 3 11", "6 6 22 22")]
    [InlineData("both-local-zero", "AB", "CD", "40 37 41 39", "3 11 3 11", "6 6 22 22")]
    [InlineData("unequal-both", "fl", "A", "1653 36 37", "3 3 12", "6 22 22")]
    [InlineData("mixed-second", "AB", "CflD", "36 37 38 1653 39", "3 11 3 12 22", "6 6 22 22 22")]
    [InlineData("combining-second", "AB", "e\u0301", "36 37 171", "3 11 3", "6 6 22")]
    public async Task GroupedLocalPreservesOwnershipSlotsAndTextInAllOutlets(
        string stem, string first, string second, string gids, string xs, string ys)
    {
        var firstMaps = stem == "both-local-zero" ? MapAt(0, 40) : stem == "unequal-both" ? MapSpan(0, 2, 1653) : "";
        var secondMaps = stem switch { "both-local-zero" => MapAt(0, 41), "unequal-both" => MapSpan(0, 1, 36, 37),
            "mixed-second" => MapSpan(1, 2, 1653), "combining-second" => MapSpan(0, 2, 171), _ => MapAt(0, 40) };
        var text = GroupedText(first, firstMaps, stem == "unequal-both" ? null : "8",
            second, secondMaps, stem == "combining-second" ? null : stem == "unequal-both" ? "9" : stem == "mixed-second" ? "9 10" : "8", inheritedX: true);
        var input = await GroupedPackage(text);
        var loaded = await new OfdReader().ReadAsync(new MemoryStream(input));
        var original = Assert.Single(loaded.Pages[0].Elements.OfType<OfdTextElement>());
        var sourceXml = original.SourceXml;
        var projection = OfdExplicitGlyphRun.Read(OfdTextRenderProjection.Create(original,
            new OfdModelMaterialization(new OfdModelLimits()), default), loaded.Fonts.Single(), strategy: Grouped)!;
        Assert.Equal(2, projection.Groups!.Count);
        Assert.Equal(first.Length, projection.Groups[1].SourceOffset);
        Assert.Equal(stem == "unequal-both" ? 1 : 2, projection.Groups[1].GlyphOffset);
        Assert.Equal(stem == "mixed-second" ? "fl" : stem == "combining-second" ? second : second[..1], projection.Groups[1].Projection.Spans.Single().LogicalText);
        Assert.Equal(sourceXml, original.SourceXml);
        if (stem == "second-only")
        {
            var legacy = OfdExplicitGlyphRun.Read(OfdTextRenderProjection.Create(original,
                new OfdModelMaterialization(new OfdModelLimits()), default), loaded.Fonts.Single())!;
            Assert.Equal(new[] { 40, 37, 38, 39 }, legacy.GlyphIds);
            Assert.Equal("A", Assert.Single(legacy.Spans).LogicalText);
            Assert.Equal("C", Assert.Single(projection.Groups[1].Projection.Spans).LogicalText);
        }
        var expectedGids = gids.Split(' ').Select(int.Parse).ToArray();
        var expectedX = xs.Split(' ').Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        var expectedY = ys.Split(' ').Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(expectedGids, projection.GlyphIds);
        Assert.Equal(expectedX.Zip(expectedY, (x, y) => (x, y)), projection.PathGlyphOrigins(default));
        foreach (var outlet in new[] { "PDF", "SVG", "HTML" })
        {
            using var output = new MemoryStream();
            await ConvertGrouped(outlet, input, output);
            if (outlet == "PDF")
            {
                using var parsed = PdfPigDocument.Open(output.ToArray());
                Assert.Equal(first + second, parsed.GetPage(1).Text);
            }
            else
            {
                var document = XDocument.Parse(Encoding.UTF8.GetString(output.ToArray()));
                Assert.Equal(new[] { first, second }, document.Descendants().Where(node => node.Attribute("data-ofd-text") is not null)
                    .Select(node => node.Attribute("data-ofd-text")!.Value));
                var paths = document.Descendants().Where(node => node.Attribute("data-original-gid") is not null).ToArray();
                Assert.Equal(expectedGids, paths.Select(node => (int)node.Attribute("data-original-gid")!));
                var outlines = new TrueTypeGlyphOutlines(Font());
                for (var index = 0; index < paths.Length; index++)
                {
                    var command = outlines.Read(expectedGids[index]).ToPathCommands().First();
                    var x = expectedX[index] + command.X * 6d / outlines.UnitsPerEm;
                    var y = expectedY[index] - command.Y * 6d / outlines.UnitsPerEm;
                    Assert.StartsWith("M " + x.ToString("0.###", CultureInfo.InvariantCulture) + " " +
                        y.ToString("0.###", CultureInfo.InvariantCulture), paths[index].Attribute("d")!.Value);
                }
            }
            if (Environment.GetEnvironmentVariable("R02_ARTIFACT_DIR") is { Length: > 0 } directory)
            {
                Directory.CreateDirectory(directory);
                await File.WriteAllBytesAsync(Path.Combine(directory, "grouped-" + stem + "." + outlet.ToLowerInvariant()), output.ToArray());
                await File.WriteAllBytesAsync(Path.Combine(directory, "grouped-" + stem + ".ofd"), input);
            }
        }
        // Reader/writer preservation is independent of the transient render plan.
        using var rewritten = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(loaded, rewritten);
        var reread = await new OfdReader().ReadAsync(new MemoryStream(rewritten.ToArray()));
        Assert.Equal(sourceXml, Assert.Single(reread.Pages[0].Elements.OfType<OfdTextElement>()).SourceXml);
    }

    [Theory]
    [InlineData("crosses-group")]
    [InlineData("overlap")]
    [InlineData("glyph-count")]
    [InlineData("missing-outline")]
    [InlineData("incomplete-delta")]
    [InlineData("absent-unequal-delta")]
    [InlineData("dangling")]
    [InlineData("non-bmp")]
    public async Task GroupedLocalRejectsInvalidPlansBeforePublicOutput(string scenario)
    {
        var maps = scenario switch
        {
            "crosses-group" => MapSpan(1, 2, 40, 41),
            "overlap" => MapAt(0, 40) + MapAt(0, 41),
            "glyph-count" => MapSpan(0, 1, 40).Replace("GlyphCount='1'", "GlyphCount='2'"),
            "missing-outline" => MapAt(0, 65535),
            "absent-unequal-delta" => MapSpan(0, 1, 36, 37),
            _ => MapAt(0, 40)
        };
        var text = GroupedText(scenario == "non-bmp" ? "A𠀀" : scenario == "absent-unequal-delta" ? "A" : "AB",
            maps, scenario == "absent-unequal-delta" ? null : scenario == "incomplete-delta" ? "8 8" : "8",
            "CD", "", "8");
        if (scenario == "dangling") text.SourceXml = text.SourceXml!.Replace("</TextObject>", MapAt(0, 41) + "</TextObject>");
        var input = await GroupedPackage(text);
        foreach (var outlet in new[] { "PDF", "SVG", "HTML" })
        {
            using var output = new MemoryStream();
            await Assert.ThrowsAnyAsync<NotSupportedException>(() => ConvertGrouped(outlet, input, output));
            Assert.Equal(0, output.Length);
        }
    }

    [Fact]
    public void MappingStrategyMustBeExplicitlyKnown()
    {
        var invalid = (OfdGlyphMappingStrategy)42;
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdToPdfConverter(new OfdToPdfOptions { GlyphMappingStrategy = invalid }));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdToSvgConverter(invalid));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdToHtmlConverter(new OfdHtmlOptions { GlyphMappingStrategy = invalid }));
    }
}
