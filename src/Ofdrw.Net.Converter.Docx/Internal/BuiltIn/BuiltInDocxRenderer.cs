using System.Collections.Generic;
using System.Threading;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

internal sealed class BuiltInDocxRenderer
{
    private readonly DocxConversionOptions _options;

    internal BuiltInDocxRenderer(DocxConversionOptions options)
    {
        _options = options;
    }

    internal IReadOnlyList<DocxConversionDiagnostic> Convert(
        string inputPath,
        string outputPath,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxConversionDiagnostic>();
        var model = new DocxModelReader(_options, diagnostics, cancellationToken).Read(inputPath);
        new BuiltInPdfRenderer(_options, diagnostics, cancellationToken).Render(model, outputPath);
        return diagnostics;
    }
}
