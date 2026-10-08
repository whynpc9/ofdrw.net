using System.Collections.Generic;

namespace Ofdrw.Net.Converter.Pdf.Vector;

/// <summary>Actual disposition and object counts for one selected page.</summary>
public sealed class PdfVectorPageResult
{
    internal PdfVectorPageResult(int sourceIndex, bool native, int paths, int texts, int images, string diagnostic)
    { SourcePageIndex = sourceIndex; IsNative = native; PathObjects = paths; TextObjects = texts; ImageObjects = images; Diagnostic = diagnostic; }
    /// <summary>Zero-based source page index; repeated selections have separate results.</summary>
    public int SourcePageIndex { get; }
    /// <summary>True only if the whole page was expressed by native events.</summary>
    public bool IsNative { get; }
    /// <summary>Number of native path objects.</summary>
    public int PathObjects { get; }
    /// <summary>Number of native or transparent semantic text objects.</summary>
    public int TextObjects { get; }
    /// <summary>Number of page image objects.</summary>
    public int ImageObjects { get; }
    /// <summary>Profile decision and, for fallback, the reason and rasterizer diagnostics.</summary>
    public string Diagnostic { get; }
}

/// <summary>Source count and ordered actual output page decisions.</summary>
public sealed class PdfVectorConversionResult
{
    internal PdfVectorConversionResult(int sourceCount, PdfVectorPageResult[] pages) { SourcePageCount = sourceCount; Pages = pages; }
    /// <summary>Number of source pages.</summary>
    public int SourcePageCount { get; }
    /// <summary>Ordered output decisions.</summary>
    public IReadOnlyList<PdfVectorPageResult> Pages { get; }
}
