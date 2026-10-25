using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class TypedGraphicRenderingTests
{
    [Theory]
    [InlineData("valid", true)]
    [InlineData("missingBoundary", false)]
    [InlineData("zeroWidth", false)]
    [InlineData("malformedCtm", false)]
    public async Task PageBlockCompositePlacementAdmitsOriginalScalarsBeforeProjection(string kind, bool accepted)
    {
        var package = new OfdDocumentPackage { CompositeResources = new() };
        var resource = new OfdCompositeResource { WidthMillimeters = 20, HeightMillimeters = 20 };
        resource.Elements.Add(new OfdPathElement { WidthMillimeters = 20, HeightMillimeters = 20,
            AbbreviatedData = "M 0 0 L 20 0 L 20 20 L 0 20 C", Fill = true, Stroke = false });
        if (!accepted)
        {
            var inside = new OfdTextElement { Text = "inside", FontName = "Arial", FontSizeMillimeters = 4,
                XMillimeters = 2, YMillimeters = 2, WidthMillimeters = 15, HeightMillimeters = 5 };
            inside.Runs.Add(new OfdTextRun { Text = "inside", YMillimeters = 4 });
            resource.Elements.Add(inside);
        }
        package.CompositeResources.Items.Add(resource);
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdCompositeElement { Resource = resource, XMillimeters = 10,
            YMillimeters = 10, WidthMillimeters = 20, HeightMillimeters = 20 });
        package.Pages.Add(page);
        if (!accepted)
        {
            page.Elements.Add(new OfdTextElement { Text = "sibling", FontName = "Arial", FontSizeMillimeters = 4,
                XMillimeters = 5, YMillimeters = 60, WidthMillimeters = 30, HeightMillimeters = 5 });
            var later = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
            later.Elements.Add(new OfdTextElement { Text = "later", FontName = "Arial", FontSizeMillimeters = 4,
                XMillimeters = 5, YMillimeters = 5, WidthMillimeters = 30, HeightMillimeters = 5 });
            package.Pages.Add(later);
        }
        using var fresh = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, fresh);
        using var patched = new MemoryStream();
        using (var input = new System.IO.Compression.ZipArchive(new MemoryStream(fresh.ToArray()), System.IO.Compression.ZipArchiveMode.Read))
        using (var output = new System.IO.Compression.ZipArchive(patched, System.IO.Compression.ZipArchiveMode.Create, true))
        {
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var target = output.CreateEntry(entry.FullName).Open();
                if (!entry.FullName.EndsWith("/Pages/Page_0/Content.xml", System.StringComparison.Ordinal))
                { source.CopyTo(target); continue; }
                var xml = System.Xml.Linq.XDocument.Load(source);
                var placement = xml.Descendants().Single(n => n.Name.LocalName == "CompositeObject");
                if (kind == "missingBoundary") placement.Attribute("Boundary")!.Remove();
                if (kind == "zeroWidth") placement.SetAttributeValue("Boundary", "10 10 0 20");
                if (kind == "malformedCtm") placement.SetAttributeValue("CTM", "1 0 bogus 1 0 0");
                placement.ReplaceWith(new System.Xml.Linq.XElement(placement.Name.Namespace + "PageBlock", new System.Xml.Linq.XElement(placement)));
                xml.Save(target);
            }
        }
        var ofd = patched.ToArray();
        var loaded = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.IsType<OfdCompositeElement>(Assert.Single(Assert.IsType<OfdRawElement>(loaded.Pages[0].Elements[0]).PageBlockChildren!));
        if (accepted)
        {
            using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
            using var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
            Assert.Contains(parsed.GetPage(1).Operations, op => op.Operator == "f");
            using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
            Assert.Contains("<path", Encoding.UTF8.GetString(svg.ToArray()));
            if (System.Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } dir)
            {
                Directory.CreateDirectory(dir);
                File.WriteAllBytes(Path.Combine(dir, "pageblock-composite-valid.ofd"), ofd);
                File.WriteAllBytes(Path.Combine(dir, "pageblock-composite-valid.pdf"), pdf.ToArray());
            }
        }
        else
        {
            var extracted = new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().ExtractPages(loaded);
            Assert.Equal(2, extracted.Count);
            Assert.Contains("inside", extracted[0]);
            Assert.Contains("sibling", extracted[0]);
            Assert.Contains("later", extracted[1]);
            await Assert.ThrowsAsync<System.NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
            await Assert.ThrowsAsync<System.NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
            await Assert.ThrowsAsync<System.NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        }
    }

    [Fact]
    public async Task SharedNestedVectorPlacementsDrawTwiceWithIntersectingClips()
    {
        var package = new OfdDocumentPackage { CompositeResources = new() };
        var leaf = new OfdCompositeResource { WidthMillimeters = 20, HeightMillimeters = 20 };
        leaf.Elements.Add(new OfdPathElement { AbbreviatedData = "M 0 0 L 20 0 L 20 20 L 0 20 C M 5 5 L 15 5 L 15 15 L 5 15 C", Fill = true, Stroke = false, FillRule = OfdFillRule.EvenOdd, WidthMillimeters = 20, HeightMillimeters = 20 });
        var parent = new OfdCompositeResource { WidthMillimeters = 30, HeightMillimeters = 30 };
        parent.Elements.Add(new OfdCompositeElement { Resource = leaf, XMillimeters = 3, YMillimeters = 4, WidthMillimeters = 20, HeightMillimeters = 20 });
        package.CompositeResources.Items.Add(leaf); package.CompositeResources.Items.Add(parent);
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdCompositeElement { Resource = parent, XMillimeters = 10, YMillimeters = 15, WidthMillimeters = 30, HeightMillimeters = 30 });
        page.Elements.Add(new OfdCompositeElement { Resource = parent, XMillimeters = 50, YMillimeters = 50, WidthMillimeters = 30, HeightMillimeters = 30, Transform = new double[] { 1, 0.2, 0, 1, 0, 0 } });
        package.Pages.Add(page);
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, input); input.Position = 0;
        using var output = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, output);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(output.ToArray());
        Assert.Equal(2, pdf.GetPage(1).Operations.Count(op => op.Operator == "f*"));
        Assert.True(pdf.GetPage(1).Operations.Count(op => op.Operator is "W" or "W*") >= 8);
        input.Position = 0; using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(input, svg);
        var xml = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Equal(2, xml.Descendants().Count(n => n.Name.LocalName == "path" && n.Attribute("fill-rule")?.Value == "evenodd"));
    }

    [Fact]
    public async Task MultipleAreasRoundtripButExportExplicitlyRefusesUnion()
    {
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        var path = new OfdPathElement { AbbreviatedData = "M 0 0 L 20 20", Clips = new() }; var clip = new OfdClip();
        for (var i = 0; i < 2; i++) clip.Areas.Add(new OfdClipArea { Path = new OfdPathElement { AbbreviatedData = "M 0 0 L 10 0 L 10 10 C", Stroke = false, Fill = true } });
        path.Clips.Items.Add(clip); page.Elements.Add(path); package.Pages.Add(page);
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, input); input.Position = 0;
        using var output = new MemoryStream(); await Assert.ThrowsAsync<System.NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(input, output));
        input.Position = 0; await Assert.ThrowsAsync<System.NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(input, output));
    }

    [Theory]
    [InlineData("qualifiedPlacement")]
    [InlineData("qualifiedDefinition")]
    [InlineData("crossNamespace")]
    [InlineData("pi")]
    public async Task OpaqueCompositeSemanticsRoundtripButCannotBeExported(string kind)
    {
        var package = new OfdDocumentPackage { CompositeResources = new() };
        var resource = new OfdCompositeResource { WidthMillimeters = 20, HeightMillimeters = 20 };
        resource.Elements.Add(new OfdPathElement { WidthMillimeters = 20, HeightMillimeters = 20, AbbreviatedData = "M 0 0 L 20 20" });
        package.CompositeResources.Items.Add(resource);
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdCompositeElement { Resource = resource, WidthMillimeters = 20, HeightMillimeters = 20 }); package.Pages.Add(page);
        using var fresh = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, fresh); fresh.Position = 0;
        using var changed = new MemoryStream();
        using (var source = new System.IO.Compression.ZipArchive(fresh, System.IO.Compression.ZipArchiveMode.Read, true))
        using (var target = new System.IO.Compression.ZipArchive(changed, System.IO.Compression.ZipArchiveMode.Create, true))
            foreach (var entry in source.Entries)
            {
                using var data = entry.Open(); using var destination = target.CreateEntry(entry.FullName).Open();
                if (!entry.FullName.EndsWith(".xml")) { data.CopyTo(destination); continue; }
                var xml = System.Xml.Linq.XDocument.Load(data);
                foreach (var node in xml.Descendants().Where(n => n.Name.LocalName == "CompositeObject").ToArray())
                {
                    if (kind == "qualifiedPlacement") node.SetAttributeValue(System.Xml.Linq.XName.Get("CTM", "urn:vendor"), "1 0 0 1 5 5");
                    if (kind == "crossNamespace") node.Name = System.Xml.Linq.XName.Get("CompositeObject", node.Name.NamespaceName == "http://www.ofdspec.org" ? "http://www.ofdspec.org/2016" : "http://www.ofdspec.org");
                    if (kind == "pi") node.AddFirst(new System.Xml.Linq.XProcessingInstruction("vendor", "drawing"));
                }
                if (kind == "qualifiedDefinition") foreach (var node in xml.Descendants().Where(n => n.Name.LocalName == "CompositeGraphicUnit")) node.SetAttributeValue(System.Xml.Linq.XName.Get("Width", "urn:vendor"), "99");
                xml.Save(destination);
            }
        changed.Position = 0; var loaded = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(changed);
        if (kind == "crossNamespace") Assert.IsType<OfdRawElement>(loaded.Pages[0].Elements[0]);
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(loaded, saved); saved.Position = 0;
        using var output = new MemoryStream(); await Assert.ThrowsAsync<System.NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(saved, output));
        saved.Position = 0; await Assert.ThrowsAsync<System.NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(saved, output));
    }
}
