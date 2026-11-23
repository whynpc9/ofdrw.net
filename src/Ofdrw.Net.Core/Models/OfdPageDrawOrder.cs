using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

/// <summary>
/// The paint order of one page: background templates, the page body, the remaining templates, then annotation
/// appearances, with composite objects and PageBlocks expanded into placed elements (accumulated CTM and inherited
/// clips). PDF, SVG/HTML and text extraction all read pages through this module, so layer order and annotation
/// projection cannot drift between them.
/// </summary>
internal static class OfdPageDrawOrder
{
    /// <summary>Placed elements in paint order; composites and PageBlocks are expanded.</summary>
    internal static IReadOnlyList<OfdElement> Placed(OfdPage page, string? documentNamespace,
        OfdModelMaterialization material, CancellationToken token, bool includeTemplates = true) =>
        OfdCompositeProjection.Expand(Layers(page, documentNamespace, material, token, includeTemplates), material, token);

    /// <summary>
    /// Unexpanded page-level elements in paint order. <paramref name="documentNamespace"/> qualifies typed (fresh)
    /// annotation appearances; read annotations keep their source namespace.
    /// </summary>
    internal static IEnumerable<OfdElement> Layers(OfdPage page, string? documentNamespace,
        OfdModelMaterialization material, CancellationToken token, bool includeTemplates = true)
    {
        if (includeTemplates)
            foreach (var template in page.Templates.Where(IsBackground))
                foreach (var element in template.Elements) yield return element;
        foreach (var element in page.Elements) yield return element;
        if (includeTemplates)
            foreach (var template in page.Templates.Where(x => !IsBackground(x)))
                foreach (var element in template.Elements) yield return element;
        foreach (var occurrence in OfdAnnotationProjection.Resolve(page, material, token, documentNamespace: documentNamespace))
            yield return occurrence.RenderElement;
    }

    private static bool IsBackground(OfdTemplateContent template) =>
        string.Equals(template.ZOrder, "Background", StringComparison.OrdinalIgnoreCase);
}
