using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>
/// In-memory page operations for an OFD document package. Operations that take the package reject
/// packages declaring <c>DocType="OFD-A"</c> with <see cref="OfdArchivalOriginalException"/>.
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

        OfdArchivalEditGuard.ThrowIfArchival(package, "OfdDocumentEditor.ReorderPages");

        var material = new OfdModelMaterialization(new OfdModelLimits());
        material.InspectStorage(package, default); material.GuardPreservedXml(package, default);
        if (pageOrder.Count != package.Pages.Count)
            throw new ArgumentException("Page order must contain every zero-based page position exactly once.", nameof(pageOrder));
        var seen = new HashSet<int>(); var reordered = new List<OfdPage>();
        foreach (var index in pageOrder)
        {
            material.Structure.Reference();
            if (index < 0 || index >= package.Pages.Count || !seen.Add(index))
                throw new ArgumentException("Page order must contain every zero-based page position exactly once.", nameof(pageOrder));
            reordered.Add(package.Pages[index]);
        }
        if (seen.Count != package.Pages.Count) throw new ArgumentException("Page order count is inconsistent.", nameof(pageOrder));
        OfdAnnotationContract.ValidatePageSelection(package, reordered, material, default);
        OfdNavigationXml.Validate(package, reordered, default, material);
        package.SourceEnvelope?.ValidatePageOrder(package, reordered, material, default);
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

        OfdArchivalEditGuard.ThrowIfArchival(package, "OfdDocumentEditor.RemovePages");

        var material = new OfdModelMaterialization(new OfdModelLimits());
        material.InspectStorage(package, default); material.GuardPreservedXml(package, default);
        var indexes = new HashSet<int>();
        foreach (var index in pageIndexes) { material.Structure.Reference(); indexes.Add(index); }
        if (indexes.Any(index => index < 0 || index >= package.Pages.Count))
        {
            throw new ArgumentOutOfRangeException(
                nameof(pageIndexes),
                "A page position is outside the document.");
        }

        var remaining = package.Pages
            .Where((_, index) => !indexes.Contains(index))
            .ToList();
        OfdAnnotationContract.ValidatePageSelection(package, remaining, material, default);
        OfdNavigationXml.Validate(package, remaining, default, material);
        package.SourceEnvelope?.ValidatePageOrder(package, remaining, material, default);
        package.Pages.Clear();
        for (var index = 0; index < remaining.Count; index++)
        {
            remaining[index].Index = index;
            package.Pages.Add(remaining[index]);
        }
    }

    /// <summary>
    /// Sets the physical crop box of a page that belongs to <paramref name="package"/>.
    /// Content coordinates remain unchanged. Packages declaring <c>DocType="OFD-A"</c> are rejected
    /// before any change.
    /// </summary>
    /// <exception cref="OfdArchivalOriginalException">The package declares OFD-A.</exception>
    /// <exception cref="ArgumentException">The page is not one of the package's pages.</exception>
    public static void CropPage(
        OfdDocumentPackage package,
        OfdPage page,
        double xMillimeters,
        double yMillimeters,
        double widthMillimeters,
        double heightMillimeters)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        if (page is null) throw new ArgumentNullException(nameof(page));
        OfdArchivalEditGuard.ThrowIfArchival(package, "OfdDocumentEditor.CropPage");
        if (!package.Pages.Any(candidate => ReferenceEquals(candidate, page)))
            throw new ArgumentException("The page does not belong to the package.", nameof(page));
        CropPage(page, xMillimeters, yMillimeters, widthMillimeters, heightMillimeters);
    }

    /// <summary>
    /// Sets a page physical crop box in the document coordinate system.
    /// Content coordinates remain unchanged. This page-level overload has no document context and
    /// cannot check the document declaration; prefer the overload that takes the package.
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
