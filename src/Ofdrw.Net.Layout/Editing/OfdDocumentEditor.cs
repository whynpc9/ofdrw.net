using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>
/// In-memory page operations for an OFD document package.
/// </summary>
public static class OfdDocumentEditor
{
    /// <summary>
    /// Reorders all pages using zero-based positions. Every existing page must
    /// occur exactly once.
    /// </summary>
    public static void ReorderPages(
        OfdDocumentPackage package,
        IReadOnlyList<int> pageOrder)
    {
        if (package is null)
        {
            throw new ArgumentNullException(nameof(package));
        }

        if (pageOrder is null)
        {
            throw new ArgumentNullException(nameof(pageOrder));
        }

        if (pageOrder.Count != package.Pages.Count ||
            pageOrder.Distinct().Count() != package.Pages.Count ||
            pageOrder.Any(index => index < 0 || index >= package.Pages.Count))
        {
            throw new ArgumentException(
                "Page order must contain every zero-based page position exactly once.",
                nameof(pageOrder));
        }

        var reordered = pageOrder.Select(index => package.Pages[index]).ToList();
        package.Pages.Clear();
        for (var index = 0; index < reordered.Count; index++)
        {
            reordered[index].Index = index;
            package.Pages.Add(reordered[index]);
        }
    }

    /// <summary>
    /// Removes the specified zero-based page positions and normalizes indexes.
    /// </summary>
    public static void RemovePages(
        OfdDocumentPackage package,
        IEnumerable<int> pageIndexes)
    {
        if (package is null)
        {
            throw new ArgumentNullException(nameof(package));
        }

        if (pageIndexes is null)
        {
            throw new ArgumentNullException(nameof(pageIndexes));
        }

        var indexes = new HashSet<int>(pageIndexes);
        if (indexes.Any(index => index < 0 || index >= package.Pages.Count))
        {
            throw new ArgumentOutOfRangeException(
                nameof(pageIndexes),
                "A page position is outside the document.");
        }

        var remaining = package.Pages
            .Where((_, index) => !indexes.Contains(index))
            .ToList();
        package.Pages.Clear();
        for (var index = 0; index < remaining.Count; index++)
        {
            remaining[index].Index = index;
            package.Pages.Add(remaining[index]);
        }
    }

    /// <summary>
    /// Sets a page physical crop box in the document coordinate system.
    /// Content coordinates remain unchanged.
    /// </summary>
    public static void CropPage(
        OfdPage page,
        double xMillimeters,
        double yMillimeters,
        double widthMillimeters,
        double heightMillimeters)
    {
        if (page is null)
        {
            throw new ArgumentNullException(nameof(page));
        }

        if (widthMillimeters <= 0 || heightMillimeters <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(widthMillimeters),
                "Crop width and height must both be positive.");
        }

        page.XMillimeters = xMillimeters;
        page.YMillimeters = yMillimeters;
        page.WidthMillimeters = widthMillimeters;
        page.HeightMillimeters = heightMillimeters;
    }
}
