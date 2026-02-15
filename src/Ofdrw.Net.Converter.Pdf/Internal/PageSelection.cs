using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal static class PageSelection
{
    public static IReadOnlyList<int> Normalize(int totalPages, IReadOnlyList<int>? requestedPages)
    {
        if (totalPages <= 0)
        {
            return [];
        }

        if (requestedPages is null || requestedPages.Count == 0)
        {
            return Enumerable.Range(0, totalPages).ToList();
        }

        var selected = requestedPages.Where(x => x >= 0 && x < totalPages).ToList();
        return selected.Count == 0 ? Enumerable.Range(0, totalPages).ToList() : selected;
    }
}
