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

        return package.Pages
            .OrderBy(page => page.Index)
            .Select(page => ExtractPage(page, includeTemplates))
            .ToList();
    }

    public string Extract(
        OfdDocumentPackage package,
        bool includeTemplates = false,
        string pageSeparator = "\f")
    {
        return string.Join(pageSeparator, ExtractPages(package, includeTemplates));
    }

    private static string ExtractPage(OfdPage page, bool includeTemplates)
    {
        IEnumerable<OfdTextElement> textElements =
            page.Elements.OfType<OfdTextElement>();
        if (includeTemplates)
        {
            textElements = page.Templates
                .Where(template => string.Equals(
                    template.ZOrder,
                    "Background",
                    StringComparison.OrdinalIgnoreCase))
                .SelectMany(template => template.Elements)
                .OfType<OfdTextElement>()
                .Concat(textElements)
                .Concat(page.Templates
                    .Where(template => !string.Equals(
                        template.ZOrder,
                        "Background",
                        StringComparison.OrdinalIgnoreCase))
                    .SelectMany(template => template.Elements)
                    .OfType<OfdTextElement>());
        }

        var lines = textElements
            .OrderBy(element => element.YMillimeters)
            .ThenBy(element => element.XMillimeters)
            .GroupBy(element => Math.Round(element.YMillimeters, 1))
            .Select(line => string.Concat(line
                .OrderBy(element => element.XMillimeters)
                .Select(element => element.Text)));
        return string.Join(Environment.NewLine, lines);
    }
}
