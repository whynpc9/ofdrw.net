using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Reader.Extraction;

/// <summary>Extracts individual verified Unicode scalars whose glyph box intersects a page rectangle.</summary>
/// <remarks>Coordinates are millimeters relative to the page PhysicalBox. Visible background templates, body,
/// then foreground templates are visited in paint order; annotation appearances are excluded. Exact embedded-font geometry is supplied by the caller.
/// Unsupported source or geometry fails the entire query.</remarks>
public sealed class OfdRegionTextExtractor
{
    private readonly IOfdGlyphGeometryProvider provider;
    private readonly OfdKeywordLocatorLimits limits;

    /// <summary>Creates an extractor using exact glyph geometry from the supplied provider.</summary>
    /// <param name="geometryProvider">Provider for the document's embedded fonts.</param>
    /// <param name="locatorLimits">Optional finite source admission limits.</param>
    public OfdRegionTextExtractor(IOfdGlyphGeometryProvider geometryProvider, OfdKeywordLocatorLimits? locatorLimits = null)
    { provider = geometryProvider ?? throw new ArgumentNullException(nameof(geometryProvider)); limits = locatorLimits ?? new(); }

    /// <summary>Returns visible text whose glyph boxes intersect the selected page rectangle.</summary>
    /// <param name="package">Source document package.</param>
    /// <param name="pageIndex">Zero-based page position.</param>
    /// <param name="region">Rectangle in page millimeters.</param>
    /// <param name="cancellationToken">Cancellation token for the query.</param>
    /// <returns>Text and fragments in page paint order.</returns>
    public OfdRegionTextResult Extract(OfdDocumentPackage package, int pageIndex, OfdGlyphBox region,
        CancellationToken cancellationToken = default)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        if (region is null) throw new ArgumentNullException(nameof(region));
        if (pageIndex < 0 || pageIndex >= package.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        if (region.Right <= region.Left || region.Bottom <= region.Top)
            throw new ArgumentException("Region must have positive area.", nameof(region));
        cancellationToken.ThrowIfCancellationRequested();
        var fragments = new List<OfdKeywordFragment>();
        var text = new StringBuilder();
        var analysis = new OfdKeywordAnalysis(provider, limits, capture: captured =>
        {
            foreach (var occurrence in captured.Traversal.Occurrences)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!captured.Sources.TryGetValue(occurrence, out var source)) continue;
                foreach (var cluster in captured.Positions[source])
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!Intersects(cluster, region)) continue;
                    captured.Budget.Fragment(1);
                    captured.Budget.Copy(1024L + 2L * cluster.Geometry.Utf16Length);
                    text.Append(cluster.Run.Text, cluster.Geometry.Utf16Start, cluster.Geometry.Utf16Length);
                    fragments.Add(new OfdKeywordFragment(cluster));
                }
            }
        });
        var scanned = analysis.CapturePositions(package, new[] { pageIndex }, null, cancellationToken);
        return new OfdRegionTextResult(pageIndex, region, text.ToString(), fragments, scanned.Diagnostics);
    }

    private static bool Intersects(OfdPositionedCluster cluster, OfdGlyphBox region)
    {
        if (cluster.Quad is null)
            return cluster.Baseline.X >= region.Left && cluster.Baseline.X < region.Right &&
                   cluster.Baseline.Y >= region.Top && cluster.Baseline.Y < region.Bottom;
        var bounds = cluster.Quad.BoundsMillimeters;
        if (bounds.Right <= region.Left || bounds.Left >= region.Right ||
            bounds.Bottom <= region.Top || bounds.Top >= region.Bottom) return false;
        var polygon = cluster.Quad.Corners.Select(point => (point.X, point.Y)).ToList();
        polygon = Clip(polygon, point => point.X - region.Left, (a, b, t) => (a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y)));
        polygon = Clip(polygon, point => region.Right - point.X, (a, b, t) => (a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y)));
        polygon = Clip(polygon, point => point.Y - region.Top, (a, b, t) => (a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y)));
        polygon = Clip(polygon, point => region.Bottom - point.Y, (a, b, t) => (a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y)));
        if (polygon.Count < 3) return false;
        double twiceArea = 0;
        for (var i = 0; i < polygon.Count; i++)
        {
            var a = polygon[i]; var b = polygon[(i + 1) % polygon.Count];
            twiceArea += a.X * b.Y - b.X * a.Y;
        }
        return Math.Abs(twiceArea) > 1e-12;
    }

    private static List<(double X, double Y)> Clip(List<(double X, double Y)> input,
        Func<(double X, double Y), double> side,
        Func<(double X, double Y), (double X, double Y), double, (double X, double Y)> interpolate)
    {
        var output = new List<(double X, double Y)>(input.Count + 1);
        if (input.Count == 0) return output;
        var previous = input[input.Count - 1]; var prior = side(previous);
        foreach (var current in input)
        {
            var next = side(current);
            if ((prior >= 0) != (next >= 0)) output.Add(interpolate(previous, current, prior / (prior - next)));
            if (next >= 0) output.Add(current);
            previous = current; prior = next;
        }
        return output;
    }
}

/// <summary>Text and source fragments selected from one page rectangle.</summary>
public sealed class OfdRegionTextResult
{
    internal OfdRegionTextResult(int pageIndex, OfdGlyphBox region, string text, List<OfdKeywordFragment> fragments,
        IReadOnlyList<OfdKeywordDiagnostic> diagnostics)
    { PageIndex = pageIndex; RegionMillimeters = region; Text = text; Fragments = new ReadOnlyCollection<OfdKeywordFragment>(fragments);
        Diagnostics = diagnostics; }
    /// <summary>Zero-based source page position.</summary>
    public int PageIndex { get; }
    /// <summary>Queried rectangle in page millimeters.</summary>
    public OfdGlyphBox RegionMillimeters { get; }
    /// <summary>Selected Unicode text in page paint order.</summary>
    public string Text { get; }
    /// <summary>Source-backed glyph fragments contributing to <see cref="Text"/>.</summary>
    public IReadOnlyList<OfdKeywordFragment> Fragments { get; }
    /// <summary>Named excluded scopes, including annotation appearances when present.</summary>
    public IReadOnlyList<OfdKeywordDiagnostic> Diagnostics { get; }
}
