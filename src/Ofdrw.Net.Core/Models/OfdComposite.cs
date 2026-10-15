using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

/// <summary>Shared vector definitions. Placements reference these objects without copying children.</summary>
public sealed class OfdCompositeResourceCollection
{
    /// <summary>Ordered typed definitions; retained imported owners preserve their source order.</summary>
    public List<OfdCompositeResource> Items { get; } = [];
}

/// <summary>A shared local vector definition with positive millimeter dimensions.</summary>
public sealed class OfdCompositeResource
{
    /// <summary>Document-unit identifier; null lets the writer allocate a new identity.</summary>
    public string? Id { get; set; }
    /// <summary>Positive local width in millimeters.</summary>
    public double WidthMillimeters { get; set; }
    /// <summary>Positive local height in millimeters.</summary>
    public double HeightMillimeters { get; set; }
    /// <summary>Ordered local graphic definitions; repeated owners are rejected.</summary>
    public List<OfdElement> Elements { get; } = [];
    /// <summary>Original Reader-owned XML, retained for extension and ownership checks.</summary>
    public string? SourceXml { get; internal set; }
    internal string? SourceEntryPath { get; set; }
    internal int SourceSlot { get; set; } = -1;
}

/// <summary>A placement of a shared vector. Scaling is explicit in Transform.</summary>
public sealed class OfdCompositeElement : OfdElement
{
    /// <summary>Creates an unbound composite placement; assign its package-owned Resource before saving.</summary>
    public OfdCompositeElement() { Type = OfdElementType.Composite; }
    /// <summary>Shared definition referenced by this placement; must belong to the same package.</summary>
    public OfdCompositeResource Resource { get; set; } = null!;
    /// <summary>Optional six-component affine transform [a,b,c,d,e,f] in local millimeter coordinates.</summary>
    public double[]? Transform { get; set; }
    /// <summary>Original Reader-owned XML, retained for extension and ownership checks.</summary>
    public string? SourceXml { get; internal set; }
}
