using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

internal static class OfdPageSelection
{
    // Preserve the established converter selection policy while bounding every
    // requested position, including invalid/duplicate positions.
    internal static IReadOnlyList<int> Normalize(int totalPages, IReadOnlyList<int>? requestedPages,
        CancellationToken cancellationToken = default, OfdModelMaterialization? material = null)
    {
        if (totalPages <= 0) return [];
        material ??= new OfdModelMaterialization(new OfdModelLimits());
        var valid = new List<int>();
        if (requestedPages is not null)
            foreach (var index in requestedPages)
            {
                cancellationToken.ThrowIfCancellationRequested(); material.Structure.Reference();
                if (index >= 0 && index < totalPages) valid.Add(index);
            }
        if (valid.Count != 0) return valid;
        for (var index = 0; index < totalPages; index++) { cancellationToken.ThrowIfCancellationRequested(); material.Structure.Reference(); valid.Add(index); }
        return valid;
    }

    internal static IReadOnlyList<int> Apply(OfdDocumentPackage package, IReadOnlyList<int>? requestedPages,
        int maximumPages = int.MaxValue, CancellationToken cancellationToken = default)
    {
        var material = new OfdModelMaterialization(new OfdModelLimits());
        material.InspectStorage(package, cancellationToken); material.GuardPreservedXml(package, cancellationToken);
        // This mutable selection seam is used by the newly rendered native DOCX
        // model. Imported source ownership uses the explicit owned Split path.
        if (package.SourceIdentity is not null || package.SourceEnvelope is not null || OfdNavigationXml.HasTypedNavigation(package, material, cancellationToken) || package.PreservedEntries.Count != 0)
            throw new NotSupportedException("Imported or interactive page selection requires the owned document splitter.");
        var selected = Normalize(package.Pages.Count, requestedPages, cancellationToken, material);
        if (selected.Count > maximumPages) throw new InvalidDataException("Page selection exceeds the configured output page limit.");
        var seen = new HashSet<int>(); var pages = new List<OfdPage>();
        foreach (var index in selected)
        {
            cancellationToken.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (!seen.Add(index))
            {
                foreach (var element in package.Pages[index].Elements)
                {
                    if (element is OfdRawElement || element is OfdTextElement { SourceXml: not null } || element is OfdImageElement { SourceXml: not null } || element is OfdPathElement { SourceXml: not null } || element.ClippingXml is not null)
                        throw new NotSupportedException("Repeated page selection cannot safely duplicate opaque source IDs or references.");
                }
            }
            pages.Add(OfdModelCloner.ClonePage(package.Pages[index], clonePayloads: false, material: material, cancellationToken: cancellationToken));
        }
        cancellationToken.ThrowIfCancellationRequested(); package.Pages.Clear();
        for (var index = 0; index < pages.Count; index++) { pages[index].Index = index; package.Pages.Add(pages[index]); }
        return selected;
    }
}
