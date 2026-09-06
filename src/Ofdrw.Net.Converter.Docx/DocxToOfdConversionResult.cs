using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Converter.Docx;

/// <summary>Reports original-text preservation and the actual source-page selection for DOCX-to-OFD.</summary>
public sealed class DocxToOfdConversionResult
{
    internal DocxToOfdConversionResult(DocxToOfdMode mode, DocxConversionEngine? engine,
        IEnumerable<DocxConversionEngine> attempts, IEnumerable<DocxConversionDiagnostic> diagnostics,
        int sourcePageCount, IEnumerable<int> sourcePages, bool originalTextPreserved, bool pageMappingAccurate)
    {
        Mode = mode; ActualEngine = engine; AttemptedEngines = attempts.ToArray();
        Diagnostics = diagnostics.ToArray(); SourcePageCount = sourcePageCount;
        SourcePages = sourcePages.ToArray(); OriginalTextPreserved = originalTextPreserved;
        PageTextMappingAccurate = pageMappingAccurate;
    }

    public DocxToOfdMode Mode { get; }
    /// <summary>Null for Native or a caller-supplied renderer without engine diagnostics.</summary>
    public DocxConversionEngine? ActualEngine { get; }
    public IReadOnlyList<DocxConversionEngine> AttemptedEngines { get; }
    public IReadOnlyList<DocxConversionDiagnostic> Diagnostics { get; }
    public int SourcePageCount { get; }
    /// <summary>Zero-based source page for every emitted OFD page; order and duplicates are preserved.</summary>
    public IReadOnlyList<int> SourcePages { get; }
    /// <summary>Whether the selected original text was preserved without a known unsupported-content loss.</summary>
    public bool OriginalTextPreserved { get; }
    /// <summary>False if document-scoped supplementary text could not be assigned a page anchor.</summary>
    public bool PageTextMappingAccurate { get; }
}
