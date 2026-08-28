using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Reports the engine and degradations involved in a DOCX conversion.
/// </summary>
public sealed class DocxConversionResult
{
    internal DocxConversionResult(
        DocxConversionEngine actualEngine,
        IEnumerable<DocxConversionEngine> attemptedEngines,
        IEnumerable<DocxConversionDiagnostic> diagnostics)
    {
        ActualEngine = actualEngine;
        AttemptedEngines = attemptedEngines.ToArray();
        Diagnostics = diagnostics.ToArray();
    }

    /// <summary>Gets the engine that produced the committed PDF.</summary>
    public DocxConversionEngine ActualEngine { get; }

    /// <summary>Gets engines attempted in order.</summary>
    public IReadOnlyList<DocxConversionEngine> AttemptedEngines { get; }

    /// <summary>Gets fallback, unsupported-feature, and font diagnostics.</summary>
    public IReadOnlyList<DocxConversionDiagnostic> Diagnostics { get; }
}
