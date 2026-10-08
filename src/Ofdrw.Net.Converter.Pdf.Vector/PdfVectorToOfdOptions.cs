using Ofdrw.Net.Converter.Pdf;

namespace Ofdrw.Net.Converter.Pdf.Vector;

/// <summary>Disposition for a page outside the documented vector profile.</summary>
public enum PdfVectorUnsupportedPagePolicy
{
    /// <summary>Replace the whole page with the existing dual-layer conversion.</summary>
    RasterizePage,
    /// <summary>Fail the conversion before writing the destination.</summary>
    Fail
}

/// <summary>Limits and explicit fallback policy for the optional vector converter.</summary>
public sealed class PdfVectorToOfdOptions
{
    /// <summary>Page policy; budget and cancellation failures never trigger fallback.</summary>
    public PdfVectorUnsupportedPagePolicy UnsupportedPagePolicy { get; set; } = PdfVectorUnsupportedPagePolicy.Fail;
    /// <summary>Existing input/page/image/namespace limits and rasterizer selection, copied at entry.</summary>
    public PdfToOfdOptions Compatibility { get; set; } = new();
    /// <summary>Maximum cumulative unfiltered content bytes for one page.</summary>
    public int MaxContentBytesPerPage { get; set; } = 8_000_000;
    /// <summary>Maximum parsed operations for one page.</summary>
    public int MaxOperationsPerPage { get; set; } = 100_000;
    /// <summary>Maximum native draw events for one page.</summary>
    public int MaxEventsPerPage { get; set; } = 10_000;
    /// <summary>Maximum cumulative path commands for one page.</summary>
    public int MaxPathCommandsPerPage { get; set; } = 100_000;
    /// <summary>Maximum cumulative original Unicode characters for one page.</summary>
    public int MaxTextCharactersPerPage { get; set; } = 100_000;
    /// <summary>Maximum bytes in each original font stream.</summary>
    public int MaxFontBytes { get; set; } = 64_000_000;
    /// <summary>Maximum cumulative embedded font bytes in the resulting package.</summary>
    public long MaxTotalFontBytes { get; set; } = 128_000_000;
    /// <summary>Maximum fonts declared in the effective page resource dictionaries.</summary>
    public int MaxFontsPerPage { get; set; } = 32;
    /// <summary>Maximum graphics-state and inherited-resource depth.</summary>
    public int MaxStackDepth { get; set; } = 64;
}
