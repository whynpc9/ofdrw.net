using System;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Xunit;

namespace Ofdrw.Net.Core.Tests;

/// <summary>
/// PDF, SVG/HTML and text extraction read a page through one draw-order module. Before, each assembled the layers
/// itself and PDF projected fresh annotation appearances without the document namespace.
/// </summary>
public sealed class OfdPageDrawOrderTests
{
    private const string Ns2016 = "http://www.ofdspec.org/2016";

    private static OfdTextElement Text(string value) => new OfdTextElement { Text = value, WidthMillimeters = 10, HeightMillimeters = 5 };

    private static OfdPage Page()
    {
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        var foreground = new OfdTemplateContent { TemplateId = "2", ZOrder = "Foreground" };
        foreground.Elements.Add(Text("foreground"));
        var background = new OfdTemplateContent { TemplateId = "1", ZOrder = "background" };
        background.Elements.Add(Text("background"));
        page.Templates.Add(foreground);
        page.Templates.Add(background);
        page.Elements.Add(Text("body"));
        var annotation = new OfdAnnotation { Type = OfdAnnotationType.Stamp, Creator = "test", LastModifiedDate = new DateTime(2026, 11, 23),
            Appearance = new OfdAnnotationAppearance { WidthMillimeters = 30, HeightMillimeters = 30 } };
        var stamped = Text("annotation");
        var clip = new OfdClip();
        clip.Areas.Add(new OfdClipArea { Path = new OfdPathElement { AbbreviatedData = "M 0 0 L 10 0 L 10 10 C", Fill = true, Stroke = false } });
        stamped.Clips = new OfdClipCollection();
        stamped.Clips.Items.Add(clip);
        annotation.Appearance.Elements.Add(stamped);
        page.Annotations = new OfdAnnotationCollection();
        page.Annotations.Items.Add(annotation);
        return page;
    }

    [Fact]
    public void BackgroundTemplatesBodyForegroundTemplatesThenAnnotations()
    {
        var placed = OfdPageDrawOrder.Placed(Page(), Ns2016, new OfdModelMaterialization(new OfdModelLimits()), CancellationToken.None);
        Assert.Equal(new[] { "background", "body", "foreground", "annotation" }, placed.OfType<OfdTextElement>().Select(x => x.Text));
    }

    [Fact]
    public void TemplatesCanBeLeftOutForBodyOnlyExtraction()
    {
        var placed = OfdPageDrawOrder.Placed(Page(), Ns2016, new OfdModelMaterialization(new OfdModelLimits()), CancellationToken.None, includeTemplates: false);
        Assert.Equal(new[] { "body", "annotation" }, placed.OfType<OfdTextElement>().Select(x => x.Text));
    }

    [Fact]
    public void FreshAnnotationAppearancesUseTheDocumentNamespace()
    {
        var annotation = OfdPageDrawOrder.Placed(Page(), Ns2016, new OfdModelMaterialization(new OfdModelLimits()), CancellationToken.None)
            .OfType<OfdTextElement>().Single(x => x.Text == "annotation");
        Assert.Contains(Ns2016, annotation.ClippingXml, StringComparison.Ordinal);
    }
}
