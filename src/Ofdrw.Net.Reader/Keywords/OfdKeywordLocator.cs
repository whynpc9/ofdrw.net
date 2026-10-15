using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Reader.Keywords;

/// <summary>Finds ordinal keywords in supported selected-document body text using exact original-font geometry.</summary>
/// <remarks>Templates/annotations are excluded named scopes. Matching never infers reading order or whitespace.
/// Source-XML deltas follow GB/T33190-2016 table46: absent axes do not move; later X/Y inherit prior TextCode starts.
/// Source-backed empty slots preserve their ordinals and have no glyphs. An all-empty owner is a continuity barrier; font bytes remain verified.
/// All inputs must remain stable. Observable callback changes to captured source/geometry/ownership inputs are rejected; concurrent transient changes
/// restored before observation are not detectable. This service does not implement replacement or write pages.</remarks>
public sealed class OfdKeywordLocator
{
    private readonly IOfdGlyphGeometryProvider provider;
    private readonly OfdKeywordLocatorLimits limits;
    /// <summary>Creates a locator with an explicit provider and finite cumulative source-admission limits.</summary>
    public OfdKeywordLocator(IOfdGlyphGeometryProvider geometryProvider, OfdKeywordLocatorLimits? locatorLimits = null)
    { provider = geometryProvider ?? throw new ArgumentNullException(nameof(geometryProvider)); limits = locatorLimits ?? new(); }

    /// <summary>Searches all current page positions within the declared body scope.</summary>
    public OfdKeywordSearchResult Find(OfdDocumentPackage source, string keyword, OfdKeywordSearchOptions? options = null,
        CancellationToken cancellationToken = default)
        => new OfdKeywordAnalysis(provider, limits).Find(source, keyword, options, cancellationToken);

    /// <summary>Searches a copied distinct page selection. Invalid or duplicate positions reject.</summary>
    public OfdKeywordSearchResult Find(OfdDocumentPackage source, IReadOnlyList<int> pageIndexes, string keyword,
        OfdKeywordSearchOptions? options = null, CancellationToken cancellationToken = default)
        => new OfdKeywordAnalysis(provider, limits).Find(source, pageIndexes, keyword, options, cancellationToken);
}
