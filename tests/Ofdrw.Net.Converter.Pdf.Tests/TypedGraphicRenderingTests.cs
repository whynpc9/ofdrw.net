using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class TypedGraphicRenderingTests
{
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
}
