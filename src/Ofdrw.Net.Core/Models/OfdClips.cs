using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

/// <summary>Fill rule applied to typed path geometry.</summary>
public enum OfdFillRule
{
    /// <summary>Nonzero winding fill rule.</summary>
    NonZero,
    /// <summary>Even-odd parity fill rule.</summary>
    EvenOdd,
}

/// <summary>Successive clips intersect. Areas within one clip form a union.</summary>
public sealed class OfdClipCollection
{
    /// <summary>Ordered typed definitions; retained imported owners preserve their source order.</summary>
    public List<OfdClip> Items { get; } = [];
    /// <summary>Original Reader-owned XML, retained for extension and ownership checks.</summary>
    public string? SourceXml { get; internal set; }
}

/// <summary>A single clipping operation; successive clips intersect.</summary>
public sealed class OfdClip
{
    /// <summary>Areas forming this clip; export currently requires one filled Path Area per Clip.</summary>
    public List<OfdClipArea> Areas { get; } = [];
    /// <summary>Original Reader-owned XML, retained for extension and ownership checks.</summary>
    public string? SourceXml { get; internal set; }
    internal int SourceSlot { get; set; } = -1;
}

/// <summary>A transformed path area inside one clipping operation.</summary>
public sealed class OfdClipArea
{
    internal OfdPathElement? OriginalPath;
    internal string? OriginalPathXml;
    /// <summary>Optional six-component affine transform [a,b,c,d,e,f] in local millimeter coordinates.</summary>
    public double[]? Transform { get; set; }
    /// <summary>Filled local path defining this clipping area.</summary>
    public OfdPathElement Path { get; set; } = new() { Stroke = false, Fill = true };
    /// <summary>Original Reader-owned XML, retained for extension and ownership checks.</summary>
    public string? SourceXml { get; internal set; }
    internal int SourceSlot { get; set; } = -1;
}

// Keeps the raw alias and typed view from becoming competing editors.
internal sealed class OfdClipOwnership
{
    internal OfdClipCollection? Value, Original;
    internal OfdModelSnapshot? Baseline;
    internal string? Raw;
    internal void Set(OfdClipCollection? value)
    {
        if (value is null && Value is not null && Baseline is not null && !Baseline.Matches(Value))
            throw new InvalidOperationException("Cannot withdraw changed typed clips.");
        Value = value;
    }
    internal void Load(OfdClipCollection value, string raw, OfdModelMaterialization material, CancellationToken token)
    { Value = Original = value; Raw = raw; Baseline = OfdModelSnapshot.Capture(value, material, token); }
    internal void Validate(string? raw, OfdModelMaterialization material, CancellationToken token)
    {
        if (Value is not null && raw != Raw && !(Original is null && string.IsNullOrEmpty(raw)))
            throw new InvalidDataException("Typed clips and legacy raw clips have mixed authority.");
        if (Value is null && Original is not null && !Baseline!.Matches(Original, material, token))
            throw new InvalidDataException("Withdrawn clip view was modified.");
    }
}
