namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Severity of a DOCX conversion diagnostic.
/// </summary>
public enum DocxConversionDiagnosticSeverity
{
    /// <summary>Informational conversion detail.</summary>
    Information,
    /// <summary>Content was degraded or substituted.</summary>
    Warning,
    /// <summary>An engine attempt failed.</summary>
    Error
}

/// <summary>
/// Describes an engine fallback, unsupported feature, or rendering substitution.
/// </summary>
public sealed class DocxConversionDiagnostic
{
    /// <summary>Initializes a diagnostic.</summary>
    public DocxConversionDiagnostic(
        string code,
        string message,
        DocxConversionDiagnosticSeverity severity = DocxConversionDiagnosticSeverity.Warning)
    {
        Code = code ?? string.Empty;
        Message = message ?? string.Empty;
        Severity = severity;
    }

    /// <summary>Gets the stable machine-readable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Gets the human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Gets the diagnostic severity.</summary>
    public DocxConversionDiagnosticSeverity Severity { get; }
}
