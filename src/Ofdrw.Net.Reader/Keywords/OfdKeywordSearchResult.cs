using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Reader.Keywords;

/// <summary>A finite millimeter point relative to the selected page's PhysicalBox top-left.</summary>
public sealed class OfdKeywordPoint
{
    internal OfdKeywordPoint(double x, double y)
    {
        if (!OfdGlyphBox.Finite(x) || !OfdGlyphBox.Finite(y)) throw new System.IO.InvalidDataException("Nonfinite transformed point.");
        X = x; Y = y;
    }
    /// <summary>Horizontal millimeters, increasing rightward.</summary>
    public double X { get; }
    /// <summary>Vertical millimeters, increasing downward.</summary>
    public double Y { get; }
}

/// <summary>A transformed original font glyph box, not a tight outline or painted-pixel quadrilateral.</summary>
public sealed class OfdKeywordQuadrilateral
{
    internal OfdKeywordQuadrilateral(OfdKeywordPoint a, OfdKeywordPoint b, OfdKeywordPoint c, OfdKeywordPoint d)
    {
        Corners = Array.AsReadOnly(new[] { a, b, c, d });
        BoundsMillimeters = new OfdGlyphBox(Math.Min(Math.Min(a.X, b.X), Math.Min(c.X, d.X)),
            Math.Min(Math.Min(a.Y, b.Y), Math.Min(c.Y, d.Y)), Math.Max(Math.Max(a.X, b.X), Math.Max(c.X, d.X)),
            Math.Max(Math.Max(a.Y, b.Y), Math.Max(c.Y, d.Y)));
    }
    /// <summary>Corners in original top-left, top-right, bottom-right, bottom-left order.</summary>
    public IReadOnlyList<OfdKeywordPoint> Corners { get; }
    /// <summary>The quadrilateral's axis-aligned envelope.</summary>
    public OfdGlyphBox BoundsMillimeters { get; }
}

/// <summary>One mapped scalar in an immutable keyword hit.</summary>
public sealed class OfdKeywordFragment
{
    internal OfdKeywordFragment(OfdPositionedCluster cluster)
    {
        var source = cluster.Source;
        OccurrencePath = Array.AsReadOnly((int[])source.OccurrencePath.Clone());
        DefinitionPath = Array.AsReadOnly((int[])source.DefinitionPath.Clone());
        SourceEntryPath = source.EntryPath; TextCodePath = cluster.Run.Path;
        LogicalUtf16Start = cluster.Geometry.Utf16Start; LogicalUtf16Length = cluster.Geometry.Utf16Length;
        SourceUtf16Start = cluster.Run.SourceOffsets[LogicalUtf16Start];
        SourceUtf16Length = cluster.Run.SourceOffsets[LogicalUtf16Start + LogicalUtf16Length] - SourceUtf16Start;
        FontResourceId = source.Font!.ResourceId; FontSha256 = source.Font.ContentSha256; FontByteCount = source.Font.ByteCount;
        GlyphId = cluster.Geometry.GlyphId; BaselineMillimeters = cluster.Baseline; InkQuadrilateralMillimeters = cluster.Quad;
        FontSizeMillimeters = source.Size; Weight = source.Weight; Italic = source.Italic;
        FontBold = source.Font.Bold; FontItalic = source.Font.Italic;
        AdvanceXMillimeters = cluster.Advance.X; AdvanceYMillimeters = cluster.Advance.Y;
        IsAuthoritativeXml = source.HasXml;
    }
    /// <summary>Operation-local occurrence identity, not a persistent document ID.</summary>
    public IReadOnlyList<int> OccurrencePath { get; }
    /// <summary>Operation-local model definition identity.</summary>
    public IReadOnlyList<int> DefinitionPath { get; }
    /// <summary>Original package entry, or null for source-free constructed models.</summary>
    public string? SourceEntryPath { get; }
    /// <summary>Namespace-qualified structural XML locator, or an explicitly model-relative run locator.</summary>
    /// <remarks>This is not XPath. Indices are zero-based operation-local ordinals: Page/Content use zero,
    /// Layer counts Content's Layers, a graphic counts all direct Layer graphic children, and TextCode counts
    /// direct TextCode children only. Use the accompanying occurrence identity and source version.</remarks>
    public string TextCodePath { get; }
    /// <summary>Start in this TextCode's decoded logical UTF16 text.</summary>
    public int LogicalUtf16Start { get; }
    /// <summary>Length in decoded logical UTF16 units.</summary>
    public int LogicalUtf16Length { get; }
    /// <summary>Start in the XML TextCode value before OFD escape decoding, or in logical model text. This is not a ZIP byte offset.</summary>
    public int SourceUtf16Start { get; }
    /// <summary>Corresponding source-value UTF16 length, including any OFD escape sequence.</summary>
    public int SourceUtf16Length { get; }
    /// <summary>Whether the locator selects authoritative object XML.</summary>
    public bool IsAuthoritativeXml { get; }
    /// <summary>Exact ordinal font resource identity.</summary>
    public string FontResourceId { get; }
    /// <summary>Internally computed original-font SHA256.</summary>
    public string FontSha256 { get; }
    /// <summary>Original embedded font byte length.</summary>
    public int FontByteCount { get; }
    /// <summary>Frozen source font size.</summary>
    public double FontSizeMillimeters { get; }
    /// <summary>Frozen source drawing weight.</summary>
    public int Weight { get; }
    /// <summary>Frozen source italic drawing flag.</summary>
    public bool Italic { get; }
    /// <summary>Actual font resource bold flag.</summary>
    public bool FontBold { get; }
    /// <summary>Actual font resource italic flag.</summary>
    public bool FontItalic { get; }
    /// <summary>Mapped original font glyph ID.</summary>
    public uint GlyphId { get; }
    /// <summary>Scalar drawing origin in page millimeters.</summary>
    public OfdKeywordPoint BaselineMillimeters { get; }
    /// <summary>Transformed font box, or null for no-ink glyphs such as space.</summary>
    public OfdKeywordQuadrilateral? InkQuadrilateralMillimeters { get; }
    /// <summary>Transformed provider advance; missing source deltas are never replaced by this advance.</summary>
    public double AdvanceXMillimeters { get; }
    /// <summary>Transformed vertical provider advance.</summary>
    public double AdvanceYMillimeters { get; }
}

/// <summary>An immutable, operation-bound hit. It supplies no implemented replacement or geometric-fit guarantee.</summary>
public sealed class OfdKeywordHit
{
    internal OfdKeywordHit(int page, string keyword, List<OfdKeywordFragment> fragments, bool editable,
        OfdKeywordSourceSnapshot snapshot)
    {
        PageIndex = page; MatchedText = keyword; Fragments = new ReadOnlyCollection<OfdKeywordFragment>(fragments.ToArray());
        CanReplace = editable; Source = snapshot; SourceVersion = snapshot.Version;
        OfdGlyphBox? bounds = null;
        foreach (var fragment in fragments)
            if (fragment.InkQuadrilateralMillimeters is { } quad)
            {
                var box = quad.BoundsMillimeters;
                bounds = bounds is null ? box : new OfdGlyphBox(Math.Min(bounds.Left, box.Left), Math.Min(bounds.Top, box.Top),
                    Math.Max(bounds.Right, box.Right), Math.Max(bounds.Bottom, box.Bottom));
            }
        BoundsMillimeters = bounds;
    }
    internal OfdKeywordSourceSnapshot Source { get; }
    /// <summary>Zero-based position in the selected document's page list.</summary>
    public int PageIndex { get; }
    /// <summary>Exact ordinal logical keyword, without normalization or injected whitespace.</summary>
    public string MatchedText { get; }
    /// <summary>Opaque operation/package-bound version. It is not a persistent global ID or authentication token.</summary>
    public string SourceVersion { get; }
    /// <summary>Whether source ownership permits a later editing preflight. Replacement/fit are not implemented or certified by this query.</summary>
    public bool CanReplace { get; }
    /// <summary>Immutable mapped scalar fragments, retaining no-ink spaces.</summary>
    public IReadOnlyList<OfdKeywordFragment> Fragments { get; }
    /// <summary>Union AABB of transformed font boxes, or null for an all-no-ink hit.</summary>
    public OfdGlyphBox? BoundsMillimeters { get; }
    /// <summary>The returned boxes' exact meaning.</summary>
    public OfdGlyphBoundsKind BoundsKind => OfdGlyphBoundsKind.TransformedFontGlyphBounds;
}

/// <summary>An unsupported body-source or named excluded-scope notice.</summary>
public sealed class OfdKeywordDiagnostic
{
    internal OfdKeywordDiagnostic(string code, int? pageIndex, string? path, bool incomplete)
    { Code = code; PageIndex = pageIndex; SourcePath = path; MakesScopeIncomplete = incomplete; }
    /// <summary>Stable reason code, containing no source text or font bytes.</summary>
    public string Code { get; }
    /// <summary>Page position, if the notice belongs to a selected page.</summary>
    public int? PageIndex { get; }
    /// <summary>Structural source locator, if known.</summary>
    public string? SourcePath { get; }
    /// <summary>Whether unsupported content occurs inside the declared search scope.</summary>
    public bool MakesScopeIncomplete { get; }
}

/// <summary>Actual cumulative locator work, separately reporting admission and matching scans.</summary>
public sealed class OfdKeywordLocatorStatistics
{
    internal OfdKeywordLocatorStatistics(OfdKeywordLocatorBudget budget, OfdKeywordOperation operation,
        long nodes, long references, OfdGlyphGeometryStatistics provider)
    {
        TraversalNodes = nodes; TraversalReferences = references; AdmissionNodes = budget.Nodes;
        AdmissionXmlBytes = budget.XmlBytes; ReservedTextUtf16Units = budget.TextUnits;
        ExpandedDeltaValues = budget.DeltaValues; DeltaTokens = budget.DeltaTokens;
        ReturnedFragments = budget.Fragments; ValidationFields = budget.ValidationFields;
        ValidationCharacters = budget.ValidationCharacters; ValidationBytes = budget.ValidationBytes;
        ReservedScalarGlyphs = budget.ReservedGlyphs; ReturnedProviderClusters = budget.ReturnedGlyphs;
        MatchingTextUtf16Units = operation.TextUtf16Units; ClusterVisits = operation.ClusterVisits;
        MatchComparisons = operation.MatchComparisons; FontHashOperations = operation.HashOperations;
        CopiedFontBytes = operation.CopiedFontBytes; ComparedFontBytes = operation.FontComparisonBytes; Provider = provider;
    }
    /// <summary>Observed A0 node visits, across its whole-model preflight.</summary>
    public long TraversalNodes { get; }
    /// <summary>Observed A0 references.</summary>
    public long TraversalReferences { get; }
    /// <summary>Additional admission XML nodes.</summary>
    public long AdmissionNodes { get; }
    /// <summary>Additional bounded admission XML bytes.</summary>
    public long AdmissionXmlBytes { get; }
    /// <summary>Logical/raw source units reserved before provider callbacks, including unsupported source.</summary>
    public long ReservedTextUtf16Units { get; }
    /// <summary>Scalar glyph requests reserved before any provider callback, including subsequently unsupported fonts.</summary>
    public long ReservedScalarGlyphs { get; }
    /// <summary>Actual provider clusters returned for validation, including subsequently clipped/unsupported runs.</summary>
    public long ReturnedProviderClusters { get; }
    /// <summary>Matching scan units charged by P0.</summary>
    public long MatchingTextUtf16Units { get; }
    /// <summary>Mapped cluster visits charged by P0 matching.</summary>
    public long ClusterVisits { get; }
    /// <summary>Expanded explicit and implicit-zero delta values.</summary>
    public long ExpandedDeltaValues { get; }
    /// <summary>Consumed delta tokens.</summary>
    public long DeltaTokens { get; }
    /// <summary>Returned fragments including overlap repetitions.</summary>
    public long ReturnedFragments { get; }
    /// <summary>Fields compared during source revalidation.</summary>
    public long ValidationFields { get; }
    /// <summary>String units compared during source revalidation.</summary>
    public long ValidationCharacters { get; }
    /// <summary>Original-entry bytes compared during source revalidation.</summary>
    public long ValidationBytes { get; }
    /// <summary>Ordinal comparisons including KMP preprocessing.</summary>
    public long MatchComparisons { get; }
    /// <summary>Actual font hash operations.</summary>
    public long FontHashOperations { get; }
    /// <summary>Actual source font bytes copied.</summary>
    public long CopiedFontBytes { get; }
    /// <summary>Actual source font bytes compared.</summary>
    public long ComparedFontBytes { get; }
    /// <summary>Observed provider counters; not a CPU/RSS guarantee.</summary>
    public OfdGlyphGeometryStatistics Provider { get; }
}

/// <summary>Results for selected-document body text only, under a declared continuous-baseline policy.</summary>
public sealed class OfdKeywordSearchResult
{
    internal OfdKeywordSearchResult(List<OfdKeywordHit> hits, List<OfdKeywordDiagnostic> diagnostics,
        int[] pages, OfdKeywordLocatorStatistics statistics)
    {
        Hits = Array.AsReadOnly(hits.ToArray()); Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        PageIndexes = Array.AsReadOnly((int[])pages.Clone()); Statistics = statistics;
        IsComplete = !diagnostics.Exists(d => d.MakesScopeIncomplete);
    }
    /// <summary>All overlapping ordinal hits within the declared search/matching policy.</summary>
    public IReadOnlyList<OfdKeywordHit> Hits { get; }
    /// <summary>Bounded unsupported-source and excluded-scope notices.</summary>
    public IReadOnlyList<OfdKeywordDiagnostic> Diagnostics { get; }
    /// <summary>Completeness only within Scope, PageIndexes, MatchPolicy and VisibilityPolicy, never a whole-file claim.</summary>
    public bool IsComplete { get; }
    /// <summary>Copied zero-based page selection in requested order.</summary>
    public IReadOnlyList<int> PageIndexes { get; }
    /// <summary>The query excludes shared templates and derived annotation appearances.</summary>
    public string Scope => "SelectedDocumentBody";
    /// <summary>No page/layer/line crossing, inferred reading order or implicit whitespace.</summary>
    public string MatchPolicy => "ContinuousBaseline";
    /// <summary>Visible opaque fill geometry before occlusion, requiring uncut font boxes.</summary>
    public string VisibilityPolicy => "VisibleOpaqueFillBeforeOcclusion";
    /// <summary>Observed work from the actual query.</summary>
    public OfdKeywordLocatorStatistics Statistics { get; }
}
