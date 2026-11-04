using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>Severity of an OFD-to-PDF conversion diagnostic.</summary>
public enum OfdToPdfDiagnosticSeverity
{
    /// <summary>Informational conversion detail.</summary>
    Information,
    /// <summary>Content was skipped or degraded; the PDF was still written.</summary>
    Warning
}

/// <summary>Describes content that OFD-to-PDF conversion skipped or degraded.</summary>
public sealed class OfdToPdfDiagnostic
{
    internal OfdToPdfDiagnostic(string code, string message, OfdToPdfDiagnosticSeverity severity,
        int? pageIndex = null, string? signatureId = null)
    {
        Code = code;
        Message = message;
        Severity = severity;
        PageIndex = pageIndex;
        SignatureId = signatureId;
    }

    /// <summary>Stable machine-readable code, for example <c>PDF_SIGNATURE_APPEARANCE_SKIPPED</c>.</summary>
    public string Code { get; }

    /// <summary>Human-readable description; it never echoes raw document content.</summary>
    public string Message { get; }

    /// <summary>Severity of the diagnostic.</summary>
    public OfdToPdfDiagnosticSeverity Severity { get; }

    /// <summary>Zero-based source page index when the diagnostic concerns one page.</summary>
    public int? PageIndex { get; }

    /// <summary>OFD signature ID when known.</summary>
    public string? SignatureId { get; }

    /// <inheritdoc/>
    public override string ToString() => $"{Severity}: {Code}: {Message}";
}

/// <summary>Result of an OFD-to-PDF conversion.</summary>
public sealed class OfdToPdfConversionResult
{
    internal OfdToPdfConversionResult(IEnumerable<int> sourcePages, int renderedSignatureAppearances,
        IEnumerable<OfdToPdfDiagnostic> diagnostics)
    {
        SourcePages = sourcePages.ToArray();
        RenderedSignatureAppearances = renderedSignatureAppearances;
        Diagnostics = diagnostics.ToArray();
    }

    /// <summary>Zero-based source page index for each output page, in output order.</summary>
    public IReadOnlyList<int> SourcePages { get; }

    /// <summary>Number of signature appearances drawn into the PDF.</summary>
    public int RenderedSignatureAppearances { get; }

    /// <summary>Skipped or degraded content, in discovery order; empty when nothing was lost.</summary>
    public IReadOnlyList<OfdToPdfDiagnostic> Diagnostics { get; }
}
