using System;
using System.Collections.Generic;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

/// <summary>A document outline tree. Empty Items removes modeled entries while retaining source extensions.</summary>
public sealed class OfdOutlineCollection : IOfdNavigationNode
{
    /// <summary>Ordered root outline entries.</summary>
    public List<OfdOutline> Items { get; } = [];
    /// <summary>Reader-owned original XML snapshot, including namespace bindings.</summary>
    public string? SourceXml { get => Source.Xml; internal set => Source.Xml = value; }
    internal OfdNavigationSource Source { get; } = new();
    OfdNavigationSource IOfdNavigationNode.Source => Source;
}

/// <summary>An outline entry with optional actions and nested entries.</summary>
public sealed class OfdOutline : IOfdNavigationNode
{
    private readonly OfdNavigationOwnership<OfdActionCollection> actions = new();
    /// <summary>Displayed title.</summary>
    public string Title { get; set; } = string.Empty;
    /// <summary>Initial expanded state; null preserves omission.</summary>
    public bool? Expanded { get; set; }
    /// <summary>Optional source leaf-count hint; not automatically recomputed.</summary>
    public int? Count { get; set; }
    /// <summary>Actions. Null preserves original XML; withdrawing a modified typed value fails.</summary>
    public OfdActionCollection? Actions { get => actions.Value; set => actions.Set(value); }
    /// <summary>Ordered nested outline entries.</summary>
    public List<OfdOutline> Children { get; } = [];
    /// <summary>Reader-owned source snapshot.</summary>
    public string? SourceXml { get => Source.Xml; internal set => Source.Xml = value; }
    internal OfdNavigationSource Source { get; } = new();
    internal OfdNavigationOwnership<OfdActionCollection> ActionOwnership => actions;
    OfdNavigationSource IOfdNavigationNode.Source => Source;
}

/// <summary>Named positions available to GoTo actions.</summary>
public sealed class OfdBookmarkCollection : IOfdNavigationNode
{
    /// <summary>Ordered named positions; names are case sensitive.</summary>
    public List<OfdBookmark> Items { get; } = [];
    /// <summary>Reader-owned source snapshot.</summary>
    public string? SourceXml { get => Source.Xml; internal set => Source.Xml = value; }
    internal OfdNavigationSource Source { get; } = new();
    OfdNavigationSource IOfdNavigationNode.Source => Source;
}

/// <summary>A named page destination.</summary>
public sealed class OfdBookmark : IOfdNavigationNode
{
    /// <summary>Unique case-sensitive bookmark name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Destination bound to a page object rather than its current index.</summary>
    public OfdDestination Destination { get; set; } = new();
    /// <summary>Reader-owned source snapshot.</summary>
    public string? SourceXml { get => Source.Xml; internal set => Source.Xml = value; }
    internal OfdNavigationSource Source { get; } = new();
    OfdNavigationSource IOfdNavigationNode.Source => Source;
}

/// <summary>How a reader aligns the destination page.</summary>
public enum OfdDestinationType
{
    /// <summary>Position and optional zoom destination.</summary>
    XYZ,
    /// <summary>Fit the complete page.</summary>
    Fit,
    /// <summary>Fit page width at a vertical position.</summary>
    FitH,
    /// <summary>Fit page height at a horizontal position.</summary>
    FitV,
    /// <summary>Fit the specified page rectangle.</summary>
    FitR,
}

/// <summary>An OFD page destination. Coordinates are millimeters.</summary>
public sealed class OfdDestination : IOfdNavigationNode
{
    /// <summary>Target page belonging to the same document.</summary>
    public OfdPage Page { get; set; } = null!;
    /// <summary>Destination alignment, default Fit.</summary>
    public OfdDestinationType Type { get; set; } = OfdDestinationType.Fit;
    /// <summary>Optional left coordinate in millimeters.</summary>
    public double? Left { get; set; }
    /// <summary>Optional top coordinate in millimeters.</summary>
    public double? Top { get; set; }
    /// <summary>Optional right coordinate in millimeters.</summary>
    public double? Right { get; set; }
    /// <summary>Optional bottom coordinate in millimeters.</summary>
    public double? Bottom { get; set; }
    /// <summary>Optional zoom; zero preserves the reader's zoom, otherwise 0.1 through 64.</summary>
    public double? Zoom { get; set; }
    /// <summary>Reader-owned source snapshot.</summary>
    public string? SourceXml { get => Source.Xml; internal set => Source.Xml = value; }
    internal OfdNavigationSource Source { get; } = new();
    OfdNavigationSource IOfdNavigationNode.Source => Source;
}

/// <summary>Ordered actions; the library serializes but does not execute them.</summary>
public sealed class OfdActionCollection : IOfdNavigationNode
{
    /// <summary>Ordered typed or preserved actions.</summary>
    public List<OfdAction> Items { get; } = [];
    /// <summary>Reader-owned source snapshot.</summary>
    public string? SourceXml { get => Source.Xml; internal set => Source.Xml = value; }
    internal OfdNavigationSource Source { get; } = new();
    OfdNavigationSource IOfdNavigationNode.Source => Source;
}

/// <summary>Base action description; no action is executed by this library.</summary>
public abstract class OfdAction : IOfdNavigationNode
{
    /// <summary>Optional DO, PO or CLICK event; outline activation commonly omits it.</summary>
    public string? Event { get; set; }
    /// <summary>Reader-owned source snapshot.</summary>
    public string? SourceXml { get => Source.Xml; internal set => Source.Xml = value; }
    internal OfdNavigationSource Source { get; } = new();
    OfdNavigationSource IOfdNavigationNode.Source => Source;
}

/// <summary>A local GoTo with exactly one direct destination or named bookmark.</summary>
public sealed class OfdGoToAction : OfdAction
{
    /// <summary>Direct target; mutually exclusive with Bookmark.</summary>
    public OfdDestination? Destination { get; set; }
    /// <summary>Bookmark object in this document; mutually exclusive with Destination.</summary>
    public OfdBookmark? Bookmark { get; set; }
}

/// <summary>An unsupported action retained as a complete Action XML element.</summary>
public sealed class OfdRawAction : OfdAction
{
    /// <summary>Complete source Action XML, including its unknown children and namespace bindings.</summary>
    public string Xml { get; set; } = string.Empty;
}

internal interface IOfdNavigationNode { OfdNavigationSource Source { get; } }

internal sealed class OfdNavigationSource
{
    internal string? Xml;
    internal object? DocumentOwner;
    internal bool HasOpaqueLocalContent;
    internal string? OriginalName;
    // Element ordinal -> modeled child identity; opaque slots are absent.
    internal readonly Dictionary<int, object> Slots = new();
}

internal sealed class OfdNavigationOwnership<T> where T : class
{
    private OfdNavigationSnapshot? baseline;
    internal T? Value { get; private set; }
    internal string? OriginalXml { get; private set; }
    internal OfdNavigationSource? OriginalSource { get; private set; }
    internal OfdNavigationSnapshot? BaselineSnapshot => baseline;

    internal void Set(T? value)
    {
        if (value is null && Value is not null && baseline is not null && !baseline.Matches(Value))
            throw new InvalidOperationException("Cannot withdraw a modified typed navigation value; null preserves source XML.");
        if (Value is null && baseline is null && value is not null) baseline = OfdNavigationSnapshot.Capture(value);
        Value = value;
    }

    internal void RestoreClonedState(T? value, string? xml, OfdNavigationSource? source, OfdNavigationSnapshot? snapshot)
    { Value = value; OriginalXml = xml; OriginalSource = source; baseline = snapshot; }

    internal void Bind(T value, OfdModelMaterialization? material = null, CancellationToken token = default)
    {
        Value = value;
        OriginalXml = ((IOfdNavigationNode)value).Source.Xml;
        OriginalSource = ((IOfdNavigationNode)value).Source;
        baseline = OfdNavigationSnapshot.Capture(value, material, token);
    }
}
