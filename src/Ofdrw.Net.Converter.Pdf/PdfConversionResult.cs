using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>Describes the page mapping and rasterizer diagnostics of a PDF-to-OFD conversion.</summary>
public sealed class PdfConversionResult
{
    internal PdfConversionResult(int sourcePageCount, IEnumerable<int> sourcePages, IEnumerable<string> diagnostics)
    {
        SourcePageCount = sourcePageCount;
        SourcePages = sourcePages.ToArray();
        Diagnostics = diagnostics.ToArray();
    }

    public int SourcePageCount { get; }
    /// <summary>Zero-based source index for each output page, including repeated selections.</summary>
    public IReadOnlyList<int> SourcePages { get; }
    public IReadOnlyList<string> Diagnostics { get; }
}

internal sealed class PdfToOfdPackageResult
{
    internal PdfToOfdPackageResult(OfdDocumentPackage package, PdfConversionResult result)
    {
        Package = package;
        Result = result;
    }
    internal OfdDocumentPackage Package { get; }
    internal PdfConversionResult Result { get; }
}
