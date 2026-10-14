using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

/// <summary>Shared vector definitions. Placements reference these objects without copying children.</summary>
public sealed class OfdCompositeResourceCollection
{
    public List<OfdCompositeResource> Items { get; } = [];
}

public sealed class OfdCompositeResource
{
    public string? Id { get; set; }
    public double WidthMillimeters { get; set; }
    public double HeightMillimeters { get; set; }
    public List<OfdElement> Elements { get; } = [];
    public string? SourceXml { get; internal set; }
    internal string? SourceEntryPath { get; set; }
    internal int SourceSlot { get; set; } = -1;
}

/// <summary>A placement of a shared vector. Scaling is explicit in Transform.</summary>
public sealed class OfdCompositeElement : OfdElement
{
    public OfdCompositeElement() { Type = OfdElementType.Composite; }
    public OfdCompositeResource Resource { get; set; } = null!;
    public double[]? Transform { get; set; }
    public string? SourceXml { get; internal set; }
}
