using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

public enum OfdFillRule { NonZero, EvenOdd }

/// <summary>Successive clips intersect. Areas within one clip form a union.</summary>
public sealed class OfdClipCollection
{
    public List<OfdClip> Items { get; } = [];
    public string? SourceXml { get; internal set; }
}

public sealed class OfdClip
{
    public List<OfdClipArea> Areas { get; } = [];
    public string? SourceXml { get; internal set; }
    internal int SourceSlot { get; set; } = -1;
}

public sealed class OfdClipArea
{
    public double[]? Transform { get; set; }
    public OfdPathElement Path { get; set; } = new() { Stroke = false, Fill = true };
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
