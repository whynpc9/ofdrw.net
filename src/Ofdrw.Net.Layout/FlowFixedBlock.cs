using System;
using Ofdrw.Net.Layout.Graphics;

namespace Ofdrw.Net.Layout;

/// <summary>A final page-coordinate rectangle in millimeters.</summary>
public sealed class FlowBoundary
{
    internal FlowBoundary(int pageIndex, double x, double y, double width, double height)
    { PageIndex = pageIndex; XMillimeters = x; YMillimeters = y; WidthMillimeters = width; HeightMillimeters = height; }

    /// <summary>Zero-based final page index.</summary>
    public int PageIndex { get; }
    /// <summary>Left edge on the physical page.</summary>
    public double XMillimeters { get; }
    /// <summary>Top edge on the physical page.</summary>
    public double YMillimeters { get; }
    /// <summary>Final width, including reserved blank space when applicable.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Final height, including reserved blank space when applicable.</summary>
    public double HeightMillimeters { get; }
}

/// <summary>A final fragment of one flow block. Paragraphs can have one fragment per line and tables one per row.</summary>
public sealed class FlowElementPlacement
{
    internal FlowElementPlacement(FlowBlock block, FlowBoundary boundary)
    { Block = block; Boundary = boundary; }
    /// <summary>Original flow block description.</summary>
    public FlowBlock Block { get; }
    /// <summary>Final fragment boundary.</summary>
    public FlowBoundary Boundary { get; }
}

/// <summary>A fixed-size, indivisible flowing block.</summary>
public abstract class FlowFixedBlock : FlowBlock
{
    internal FlowFixedBlock(double widthMillimeters, double heightMillimeters)
    { WidthMillimeters = widthMillimeters; HeightMillimeters = heightMillimeters; }

    /// <summary>Requested width in millimeters. It must fit the content width; no implicit scaling occurs.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Requested reserved height in millimeters. It must fit one fresh page.</summary>
    public double HeightMillimeters { get; }
    /// <summary>Horizontal placement within the page content width.</summary>
    public ParagraphAlignment Alignment { get; set; }
    /// <summary>Space before the block in millimeters.</summary>
    public double SpaceBeforeMillimeters { get; set; }
    /// <summary>Space after the block in millimeters.</summary>
    public double SpaceAfterMillimeters { get; set; }
    /// <summary>Starts a new page when preceding flow content exists.</summary>
    public bool PageBreakBefore { get; set; }
}

/// <summary>A PNG or baseline JPEG occupying an indivisible flowing rectangle.</summary>
public sealed class ImageBlock : FlowFixedBlock
{
    private readonly byte[] _encodedImage;
    /// <summary>Copies encoded bytes so later caller mutations cannot change admission or output.</summary>
    public ImageBlock(byte[] encodedImage, double widthMillimeters, double heightMillimeters)
        : base(widthMillimeters, heightMillimeters)
    { _encodedImage = (byte[])(encodedImage ?? throw new ArgumentNullException(nameof(encodedImage))).Clone(); }
    /// <summary>Object opacity from transparent (0) to opaque (255).</summary>
    public int Alpha { get; set; } = 255;
    internal byte[] SnapshotImage() => (byte[])_encodedImage.Clone();
}

/// <summary>A reserved flowing rectangle drawn after its final page position is known.</summary>
public sealed class CanvasBlock : FlowFixedBlock
{
    /// <summary>Creates a block with local-coordinate drawing. An empty callback is a valid blank reservation.</summary>
    public CanvasBlock(double widthMillimeters, double heightMillimeters,
        Action<OfdGraphics, FlowBoundary> draw)
        : base(widthMillimeters, heightMillimeters)
    { Draw = draw ?? throw new ArgumentNullException(nameof(draw)); }

    /// <summary>Called once after all flow measurements succeed. The graphics origin is the final boundary's top-left and drawing is clipped to it.</summary>
    public Action<OfdGraphics, FlowBoundary> Draw { get; }
}
