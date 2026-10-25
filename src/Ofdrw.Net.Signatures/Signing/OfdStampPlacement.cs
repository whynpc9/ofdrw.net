using System.Collections.Generic;

namespace Ofdrw.Net.Signatures.Signing;

/// <summary>Placement of one or more appearances from the signature provider's seal.</summary>
public abstract class OfdStampPlacement
{
    internal abstract OfdStampPlacement Snapshot();
}

/// <summary>One complete seal appearance on a numbered page.</summary>
public sealed class OfdNormalStampPlacement : OfdStampPlacement
{
    /// <summary>Creates a complete seal at an OFD page position, in millimeters.</summary>
    public OfdNormalStampPlacement(int pageNumber, double xMillimeters, double yMillimeters,
        double widthMillimeters, double heightMillimeters)
    {
        PageNumber = pageNumber;
        XMillimeters = xMillimeters;
        YMillimeters = yMillimeters;
        WidthMillimeters = widthMillimeters;
        HeightMillimeters = heightMillimeters;
    }

    /// <summary>One-based page number in the signed document.</summary>
    public int PageNumber { get; }
    /// <summary>Left edge in the page's OFD coordinate system.</summary>
    public double XMillimeters { get; }
    /// <summary>Top edge in the page's OFD coordinate system.</summary>
    public double YMillimeters { get; }
    /// <summary>Displayed seal width in millimeters.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Displayed seal height in millimeters.</summary>
    public double HeightMillimeters { get; }

    internal override OfdStampPlacement Snapshot() => this;
}

/// <summary>The page edge on which each riding-seal slice is visible.</summary>
public enum OfdStampEdge
{
    /// <summary>Left page edge; the first page receives the rightmost image slice.</summary>
    Left,
    /// <summary>Right page edge; the first page receives the leftmost image slice.</summary>
    Right,
    /// <summary>Top page edge; the first page receives the bottom image slice.</summary>
    Top,
    /// <summary>Bottom page edge; the first page receives the top image slice.</summary>
    Bottom
}

/// <summary>A full seal divided across consecutive pages, optionally repeated by group.</summary>
public sealed class OfdRidingStampPlacement : OfdStampPlacement
{
    /// <summary>Creates a riding seal on a one-based page range, in millimeters.</summary>
    public OfdRidingStampPlacement(OfdStampEdge edge, int firstPageNumber, int pageCount,
        double widthMillimeters, double heightMillimeters)
    {
        Edge = edge;
        FirstPageNumber = firstPageNumber;
        PageCount = pageCount;
        WidthMillimeters = widthMillimeters;
        HeightMillimeters = heightMillimeters;
    }

    /// <summary>Edge where the slice appears.</summary>
    public OfdStampEdge Edge { get; }
    /// <summary>One-based first page number in the signed document.</summary>
    public int FirstPageNumber { get; }
    /// <summary>Number of consecutive pages receiving slices.</summary>
    public int PageCount { get; }
    /// <summary>Width of the complete seal before clipping.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Height of the complete seal before clipping.</summary>
    public double HeightMillimeters { get; }
    /// <summary>Pages per repeated seal; zero means one seal across the entire range.</summary>
    public int GroupSize { get; set; }
    /// <summary>
    /// Positive relative widths or heights for one full group. Empty means equal slices.
    /// A short final group uses and renormalizes the first weights.
    /// </summary>
    public IList<double> SliceWeights { get; } = new List<double>();
    /// <summary>Distance of the visible slice from its chosen page edge.</summary>
    public double MarginMillimeters { get; set; }
    /// <summary>
    /// Distance along the edge from the page's top or left; null centers the full seal.
    /// </summary>
    public double? OffsetMillimeters { get; set; }

    internal override OfdStampPlacement Snapshot()
    {
        var copy = new OfdRidingStampPlacement(Edge, FirstPageNumber, PageCount,
            WidthMillimeters, HeightMillimeters)
        {
            GroupSize = GroupSize,
            MarginMillimeters = MarginMillimeters,
            OffsetMillimeters = OffsetMillimeters
        };
        foreach (var weight in SliceWeights) copy.SliceWeights.Add(weight);
        return copy;
    }
}
