using System;
using Ofdrw.Net.Core.Constants;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>
/// Options that control PDF to OFD conversion fidelity.
/// </summary>
public sealed class PdfToOfdOptions
{
    /// <summary>Maximum input bytes, enforced while staging the PDF.</summary>
    public long MaxInputBytes { get; set; } = 128L * 1024 * 1024;

    /// <summary>Maximum source or selected output page count.</summary>
    public int MaxPageCount { get; set; } = 10_000;

    /// <summary>Maximum decoded pixels in any one rendered page.</summary>
    public long MaxRasterizedPixelsPerPage { get; set; } = 40_000_000;

    /// <summary>Maximum accumulated compressed page-image bytes.</summary>
    public long MaxTotalImageBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Maximum time allowed for an external pdftoppm attempt.</summary>
    public TimeSpan ExternalRasterizationTimeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Gets or sets the target DPI used when rasterizing PDF pages.
    /// Values outside 72–300 are clamped.
    /// </summary>
    public int Dpi { get; set; } = 144;

    /// <summary>
    /// Gets or sets whether conversion must succeed via page rasterization.
    /// When <c>true</c> (default), text extraction is not used as a visual fallback
    /// because it cannot preserve table/grid layout. It may still be used for the
    /// optional machine-readable text layer.
    /// </summary>
    public bool RequireRasterization { get; set; } = true;

    /// <summary>
    /// Gets or sets whether <c>pdftoppm</c> should be tried before the built-in
    /// Docnet/Pdfium rasterizer. Default is <c>false</c> (Docnet first).
    /// </summary>
    public bool PreferExternalPdfToPpm { get; set; }

    /// <summary>
    /// Gets or sets how extractable PDF text is retained in the OFD output.
    /// The default creates a dual-layer document: a rendered page image for visual
    /// fidelity and transparent, positioned OFD text objects for machine use.
    /// Semantic extraction rejects unverified right-to-left/bidi text. Use
    /// <see cref="PdfTextLayerMode.None"/> with <see cref="RequireRasterization"/>
    /// enabled for an explicit image-only conversion of such input.
    /// </summary>
    public PdfTextLayerMode TextLayerMode { get; set; } = PdfTextLayerMode.Invisible;

    /// <summary>
    /// Gets or sets the maximum number of semantic text objects emitted for one page.
    /// This bounds memory and output growth for untrusted PDFs.
    /// </summary>
    public int MaxTextObjectsPerPage { get; set; } = 50_000;

    /// <summary>Copy local PDF outline destinations into OFD Outlines. Off by default for compatibility.</summary>
    public bool CopyBookmarks { get; set; }

    /// <summary>Copy embedded document files and page FileAttachment annotations. Off by default for compatibility.</summary>
    public bool CopyAttachments { get; set; }

    // Vector raster fallback retains PDF annotation pixels but imports metadata only in its final package.
    internal bool RasterizeCropBox { get; set; }

    /// <summary>Maximum number of copied outline nodes.</summary>
    public int MaxBookmarkCount { get; set; } = 10_000;

    /// <summary>Maximum outline nesting depth.</summary>
    public int MaxBookmarkDepth { get; set; } = 64;

    /// <summary>Maximum number of copied attachment occurrences.</summary>
    public int MaxAttachmentCount { get; set; } = 1_000;

    /// <summary>Maximum decoded bytes in one embedded attachment.</summary>
    public long MaxAttachmentBytes { get; set; } = 32L * 1024 * 1024;

    /// <summary>Maximum decoded attachment bytes across the output document.</summary>
    public long MaxTotalAttachmentBytes { get; set; } = 128L * 1024 * 1024;

    /// <summary>
    /// Gets or sets the XML namespace written into the OFD parts. The default is the
    /// GB/T 33190-2016 URI <see cref="OfdConstants.StandardNamespace"/>. Readers built
    /// around the OFD-H medical profile or pre-2016 drafts only resolve
    /// <see cref="OfdConstants.Namespace"/> (<c>http://www.ofdspec.org</c>) and show
    /// blank pages for <c>/2016</c> packages; use that URI when targeting them.
    /// </summary>
    public string Namespace { get; set; } = OfdConstants.StandardNamespace;
}
