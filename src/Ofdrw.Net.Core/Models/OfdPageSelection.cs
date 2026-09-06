using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdPageSelection
{
    // Keep the established PDF contract across every multi-page converter:
    // preserve requested order/duplicates, ignore individual invalid indexes,
    // and use all pages for an omitted, empty, or entirely invalid selection.
    internal static IReadOnlyList<int> Normalize(int totalPages, IReadOnlyList<int>? requestedPages)
    {
        if (totalPages <= 0) return [];
        var valid = requestedPages?.Where(index => index >= 0 && index < totalPages).ToList();
        return valid is { Count: > 0 } ? valid : Enumerable.Range(0, totalPages).ToList();
    }

    internal static IReadOnlyList<int> Apply(OfdDocumentPackage package, IReadOnlyList<int>? requestedPages, int maximumPages = int.MaxValue)
    {
        var selected = Normalize(package.Pages.Count, requestedPages);
        if (selected.Count > maximumPages) throw new InvalidDataException("Page selection exceeds the configured output page limit.");
        var seen = new HashSet<int>();
        var pages = selected.Select(index => OfdModelCloner.ClonePage(package.Pages[index], seen.Add(index), clonePayloads: false)).ToList();
        package.Pages.Clear();
        for (var index = 0; index < pages.Count; index++)
        {
            pages[index].Index = index;
            package.Pages.Add(pages[index]);
        }
        return selected;
    }
}
