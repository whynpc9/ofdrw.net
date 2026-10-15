using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdDocumentPackage
{
    internal OfdCompositeResourceCollection? CompositeValue, OriginalCompositeResources;
    internal OfdModelSnapshot? CompositeBaseline;
    internal OfdCompositeResource[] OriginalCompositeOrder = Array.Empty<OfdCompositeResource>();
    /// <summary>Shared vector definitions. Null preserves imported content; imported definitions require proven closure before removal or reorder.</summary>
    public OfdCompositeResourceCollection? CompositeResources
    {
        get => CompositeValue;
        set
        {
            if (value is null && CompositeValue is not null && CompositeBaseline is not null && !CompositeBaseline.Matches(CompositeValue))
                throw new InvalidOperationException("Cannot withdraw changed typed composite resources.");
            CompositeValue = value;
        }
    }
    internal void ValidateCompositeOwnership(OfdModelMaterialization material, System.Threading.CancellationToken token)
    {
        if (CompositeValue is null && OriginalCompositeResources is not null && !CompositeBaseline!.Matches(OriginalCompositeResources, material, token))
            throw new System.IO.InvalidDataException("Withdrawn composite resources were modified.");
        if (CompositeValue is not null && OriginalCompositeOrder.Length > 0 &&
            (CompositeValue.Items.Count < OriginalCompositeOrder.Length || !OriginalCompositeOrder.SequenceEqual(CompositeValue.Items.Take(OriginalCompositeOrder.Length))))
            throw new NotSupportedException("Imported composite definitions cannot be removed or reordered without proven closure; delete placements and let Writer prune unreachable resources.");
        if (OriginalCompositeResources is not null && CompositeValue is not null && CompositeValue != OriginalCompositeResources)
            throw new NotSupportedException("Imported composite collection replacement requires a proven resource closure.");
    }
    internal OfdDocumentEnvelope? SourceEnvelope { get; set; }
    internal OfdAnnotationIndexSource? AnnotationIndexSource { get; set; }
    internal OfdCustomTagSource? CustomTagSource { get; set; }
    internal OfdSourceIdentity? SourceIdentity { get; set; }
    private readonly OfdNavigationOwnership<OfdOutlineCollection> outlines = new();
    private readonly OfdNavigationOwnership<OfdBookmarkCollection> bookmarks = new();
    private readonly OfdNavigationOwnership<OfdActionCollection> actions = new();

    /// <summary>Outline tree. Null preserves source; an empty collection removes modeled entries while retaining extensions.</summary>
    public OfdOutlineCollection? Outlines { get => outlines.Value; set => outlines.Set(value); }

    /// <summary>Named positions. Null preserves source; withdrawing a changed typed collection fails.</summary>
    public OfdBookmarkCollection? Bookmarks { get => bookmarks.Value; set => bookmarks.Set(value); }

    /// <summary>Document actions, serialized but never executed.</summary>
    public OfdActionCollection? Actions { get => actions.Value; set => actions.Set(value); }

    internal OfdNavigationOwnership<OfdOutlineCollection> OutlineOwnership => outlines;
    internal OfdNavigationOwnership<OfdBookmarkCollection> BookmarkOwnership => bookmarks;
    internal OfdNavigationOwnership<OfdActionCollection> ActionOwnership => actions;

    /// <summary>Original document XML entry, or null for a newly created document.</summary>
    public string? DocumentEntryPath { get; set; }

    public OfdDocumentOptions Options { get; set; } = new OfdDocumentOptions();

    public List<OfdPage> Pages { get; } = [];

    public List<OfdAttachment> Attachments { get; } = [];

    public List<OfdFontResource> Fonts { get; } = [];

    public Dictionary<string, string> CustomTags { get; } = new Dictionary<string, string>();

    /// <summary>
    /// Original package entries retained while reading an existing OFD package.
    /// Writers copy these entries first and replace only entries they regenerate,
    /// which prevents unrelated resources and extension files from being discarded.
    /// </summary>
    public Dictionary<string, byte[]> PreservedEntries { get; } =
        new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// CommonData children not owned by the current strongly typed model, such as
    /// template page declarations and default color spaces.
    /// </summary>
    public List<string> PreservedCommonDataElements { get; } = [];

    /// <summary>
    /// Document children not owned by the current strongly typed model, such as
    /// annotations, outlines, permissions, actions, bookmarks and extensions.
    /// </summary>
    public List<string> PreservedDocumentElements { get; } = [];

    /// <summary>
    /// DocBody children other than DocInfo and DocRoot, including version and
    /// signature declarations.
    /// </summary>
    public List<string> PreservedDocBodyElements { get; } = [];

    public string? PublicResourceLocation { get; set; }

    public string? DocumentResourceLocation { get; set; }
}
