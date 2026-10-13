using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

/// <summary>The five annotation types defined by OFD.</summary>
public enum OfdAnnotationType { Link, Path, Highlight, Stamp, Watermark }

/// <summary>Ordered editable annotations. Unmodeled source slots are preserved separately.</summary>
public sealed class OfdAnnotationCollection
{
    public List<OfdAnnotation> Items { get; } = [];
}

/// <summary>An annotation whose Appearance is expressed in local coordinates.</summary>
public sealed class OfdAnnotation
{
    public string? Id { get; set; }
    public OfdAnnotationType Type { get; set; }
    public string Creator { get; set; } = string.Empty;
    /// <summary>A civil date: midnight with DateTimeKind.Unspecified.</summary>
    public DateTime? LastModifiedDate { get; set; }
    public string? Subtype { get; set; }
    public bool? Visible { get; set; }
    public bool? Print { get; set; }
    public bool? NoZoom { get; set; }
    public bool? NoRotate { get; set; }
    public bool? ReadOnly { get; set; }
    public string? Remark { get; set; }
    public OfdAnnotationAppearance Appearance { get; set; } = new();
    public string? SourceXml { get; internal set; }
    internal OfdAnnotationSource? Source { get; set; }
}

/// <summary>Local primitive definitions, placed once by Boundary and optional CTM.</summary>
public sealed class OfdAnnotationAppearance
{
    public double XMillimeters { get; set; }
    public double YMillimeters { get; set; }
    public double WidthMillimeters { get; set; }
    public double HeightMillimeters { get; set; }
    public double[]? Transform { get; set; }
    public List<OfdElement> Elements { get; } = [];
    public string? SourceXml { get; internal set; }
}

internal sealed class OfdAnnotationSource(string entry, int slot, int projectionStart, int projectionCount,
    OfdModelSnapshot primitives, string sourceNamespace)
{
    internal string Entry { get; } = entry;
    internal int Slot { get; } = slot;
    internal int ProjectionStart { get; } = projectionStart;
    internal int ProjectionCount { get; } = projectionCount;
    internal string Namespace { get; } = sourceNamespace;
    internal OfdModelSnapshot Primitives { get; } = primitives;
    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("sourceNamespace", Namespace); sink.Item("entry", Entry); sink.Path(Entry); sink.Item("slot", Slot);
        sink.Item("projectionStart", ProjectionStart); sink.Item("projectionCount", ProjectionCount);
        Primitives.VisitInventory(sink);
    }
}

internal sealed class OfdAnnotationOwnership
{
    internal OfdAnnotationCollection? Value { get; private set; }
    internal OfdAnnotationCollection? Original { get; private set; }
    internal OfdAnnotationCollection? BaselineCollection { get; private set; }
    internal OfdModelSnapshot? Baseline { get; private set; }
    internal OfdModelSnapshot? ProjectionBaseline { get; private set; }
    // Includes entries without admitted slots, so unsupported/shared sources cannot be claimed accidentally.
    internal readonly Dictionary<string, string> Entries = new(StringComparer.OrdinalIgnoreCase);
    internal bool HasUnclaimedSource;
    internal readonly HashSet<string> UnsafeDeletionEntries = new(StringComparer.OrdinalIgnoreCase);

    internal void Set(OfdAnnotationCollection? value)
    {
        if (value is null && Value is not null && Baseline is not null && !Baseline.Matches(Value))
            throw new InvalidOperationException("Cannot withdraw modified annotations; null preserves source.");
        if (Baseline is null && value is not null) { BaselineCollection = value; Baseline = OfdModelSnapshot.Capture(value); }
        Value = value;
    }
    internal void Bind(OfdPage page, OfdAnnotationCollection? value, OfdModelMaterialization material, CancellationToken token)
    {
        Value = value;
        if (value is not null)
        {
            material.ChargeCopy(128L + value.Items.Count * 16L, token);
            Original = new OfdAnnotationCollection(); Original.Items.AddRange(value.Items);
        }
        BaselineCollection = value;
        Baseline = value is null ? null : OfdModelSnapshot.Capture(value, material, token);
        ProjectionBaseline = OfdModelSnapshot.Capture(page.AnnotationAppearances, material, token);
    }
    internal void Validate(OfdPage page, OfdModelMaterialization material, CancellationToken token)
    {
        if (Value is null && BaselineCollection is not null && Baseline is not null && !Baseline.Matches(BaselineCollection, material, token))
            throw new InvalidDataException("Withdrawn annotation definitions changed through a retained reference.");
        if (ProjectionBaseline is not null && !ProjectionBaseline.Matches(page.AnnotationAppearances, material, token))
            throw new InvalidDataException("Reader-owned annotation projections cannot be edited. Edit local Annotations instead.");
        if (ProjectionBaseline is null && Value is not null && page.AnnotationAppearances.Count != 0)
            throw new InvalidDataException("Typed annotations and a manual legacy projection have ambiguous ownership.");
        var original = Original?.Items;
        if (Value is not null)
        {
            if (HasUnclaimedSource && original is null)
                throw new NotSupportedException("This annotation source has no safely editable slots.");
            material.ChargeCopy(256L + (original?.Count ?? 0) * 128L, token);
            var originals = new Dictionary<OfdAnnotation, int>();
            if (original is not null) for (var index = 0; index < original.Count; index++) { token.ThrowIfCancellationRequested(); originals.Add(original[index], index); }
            var seen = new HashSet<OfdAnnotation>(); var previous = -1; var added = false;
            foreach (var annotation in Value.Items)
            {
                token.ThrowIfCancellationRequested(); material.Structure.Node(1);
                if (annotation is null || !seen.Add(annotation)) throw new InvalidDataException("Repeated/null annotation ownership.");
                if (!Enum.IsDefined(typeof(OfdAnnotationType), annotation.Type) || string.IsNullOrWhiteSpace(annotation.Creator) ||
                    annotation.LastModifiedDate is not DateTime date || date.TimeOfDay != TimeSpan.Zero || date.Kind != DateTimeKind.Unspecified)
                    throw new InvalidDataException("Annotations require Type, Creator and a civil LastModifiedDate.");
                var appearance = annotation.Appearance ?? throw new InvalidDataException("Annotation requires Appearance.");
                var values = new[] { appearance.XMillimeters, appearance.YMillimeters, appearance.WidthMillimeters, appearance.HeightMillimeters };
                if (values.Any(v => double.IsNaN(v) || double.IsInfinity(v)) || appearance.WidthMillimeters <= 0 || appearance.HeightMillimeters <= 0 ||
                    appearance.Transform is { } matrix && (matrix.Length != 6 || matrix.Any(v => double.IsNaN(v) || double.IsInfinity(v))))
                    throw new InvalidDataException("Invalid local annotation geometry.");
                if (annotation.Source is { } source)
                {
                    if (!originals.TryGetValue(annotation, out var position)) throw new InvalidDataException("Foreign annotation source owner.");
                    if (added) throw new NotSupportedException("New annotations must follow retained source annotations.");
                    if (position <= previous) throw new NotSupportedException("Imported annotation reorder requires an opaque slot remap.");
                    previous = position;
                    if (!source.Primitives.Matches(appearance.Elements, material, token))
                        throw new NotSupportedException("Imported annotation primitive edits require source-slot reconciliation; Appearance placement and metadata remain editable.");
                }
                else
                {
                    added = true;
                    if (annotation.SourceXml is not null || appearance.SourceXml is not null) throw new InvalidDataException("Unauthenticated annotation source XML.");
                }
            }
        }
    }
    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("unclaimed", HasUnclaimedSource); sink.Item("annotationEntries", Entries, true); sink.Item("count", Entries.Count);
        foreach (var pair in Entries) { sink.Item("entry", pair.Key); sink.Path(pair.Key); sink.Item("entryXml", pair.Value); sink.Xml(pair.Value); }
        sink.Item("unsafeDeletionEntries", UnsafeDeletionEntries, true); sink.Item("count", UnsafeDeletionEntries.Count);
        foreach (var path in UnsafeDeletionEntries) { sink.Item("unsafeDeletionEntry", path); sink.Path(path); }
        sink.Item("annotationBaseline", Baseline, true); Baseline?.VisitInventory(sink);
        sink.Item("projectionBaseline", ProjectionBaseline, true); ProjectionBaseline?.VisitInventory(sink);
    }
    internal void Restore(OfdAnnotationCollection? value, OfdAnnotationCollection? original, OfdModelSnapshot? baseline,
        OfdModelSnapshot? projection, bool unclaimed, OfdAnnotationCollection? baselineCollection)
    { Value = value; Original = original; Baseline = baseline; ProjectionBaseline = projection; HasUnclaimedSource = unclaimed; BaselineCollection = baselineCollection; }
}
