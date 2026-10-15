using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Reader.Extraction;

/// <summary>
/// Extracts logical text from the strongly typed page model.
/// </summary>
public sealed class OfdTextExtractor
{
    public IReadOnlyList<string> ExtractPages(
        OfdDocumentPackage package,
        bool includeTemplates = false)
    {
        if (package is null)
        {
            throw new ArgumentNullException(nameof(package));
        }

        var material = new OfdModelMaterialization(new OfdModelLimits());
        OfdAnnotationContract.Validate(package, material, default);
        OfdCompositeGraph.Validate(package, material, default);
        return package.Pages
            .OrderBy(page => page.Index)
            .Select(page => ExtractPage(page, includeTemplates, material, package.Options.Namespace))
            .ToList();
    }

    public string Extract(
        OfdDocumentPackage package,
        bool includeTemplates = false,
        string pageSeparator = "\f")
    {
        return string.Join(pageSeparator, ExtractPages(package, includeTemplates));
    }

    private static string ExtractPage(OfdPage page, bool includeTemplates, OfdModelMaterialization material, string documentNamespace)
    {
        IEnumerable<OfdTextElement> textElements =
            OfdCompositeProjection.Expand(page.Elements, material, default).OfType<OfdTextElement>();
        if (includeTemplates)
        {
            textElements = page.Templates
                .Where(template => string.Equals(
                    template.ZOrder,
                    "Background",
                    StringComparison.OrdinalIgnoreCase))
                .SelectMany(template => OfdCompositeProjection.Expand(template.Elements, material, default))
                .OfType<OfdTextElement>()
                .Concat(textElements)
                .Concat(page.Templates
                    .Where(template => !string.Equals(
                        template.ZOrder,
                        "Background",
                        StringComparison.OrdinalIgnoreCase))
                    .SelectMany(template => OfdCompositeProjection.Expand(template.Elements, material, default))
                    .OfType<OfdTextElement>());
        }

        textElements = textElements.Concat(OfdAnnotationProjection.Resolve(page, material, documentNamespace: documentNamespace).Select(occurrence => occurrence.RenderElement).OfType<OfdTextElement>());
        var lines = textElements
            .Where(OfdGraphicXmlContract.IsVisible)
            .OrderBy(element => element.YMillimeters)
            .ThenBy(element => element.XMillimeters)
            .GroupBy(element => Math.Round(element.YMillimeters, 1))
            .Select(line => string.Concat(line
                .OrderBy(element => element.XMillimeters)
                .Select(element => element.Text)));
        return string.Join(Environment.NewLine, lines);
    }
}
