using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Pdf.Converters;

namespace Ofdrw.Net.Converter.Docx.Converters;

/// <summary>
/// Converts DOCX to OFD by composing DOCX-to-PDF and PDF-to-OFD converters.
/// </summary>
public sealed class DocxToOfdConverter : IDocxToOfdConverter
{
    private readonly IDocxToPdfConverter _docxToPdf;
    private readonly IPdfToOfdConverter _pdfToOfd;

    /// <summary>
    /// Initializes a converter using the default DOCX renderer and PDF converter.
    /// </summary>
    public DocxToOfdConverter()
        : this(new DocxToPdfConverter(), new PdfToOfdConverter())
    {
    }

    /// <summary>
    /// Initializes a converter using the supplied DOCX options.
    /// </summary>
    public DocxToOfdConverter(DocxConversionOptions options)
        : this(new DocxToPdfConverter(options), new PdfToOfdConverter())
    {
    }

    /// <summary>
    /// Initializes a converter using explicit pipeline stages.
    /// </summary>
    public DocxToOfdConverter(IDocxToPdfConverter docxToPdf, IPdfToOfdConverter pdfToOfd)
    {
        _docxToPdf = docxToPdf ?? throw new ArgumentNullException(nameof(docxToPdf));
        _pdfToOfd = pdfToOfd ?? throw new ArgumentNullException(nameof(pdfToOfd));
    }

    /// <inheritdoc />
    public async Task ConvertAsync(
        Stream docxInput,
        Stream ofdOutput,
        IReadOnlyList<int>? pages = null,
        CancellationToken cancellationToken = default)
    {
        if (docxInput is null)
        {
            throw new ArgumentNullException(nameof(docxInput));
        }

        if (ofdOutput is null)
        {
            throw new ArgumentNullException(nameof(ofdOutput));
        }

        var tempPdfPath = Path.Combine(Path.GetTempPath(), $"ofdrw-docx-stage-{Guid.NewGuid():N}.pdf");
        try
        {
            using (var pdfOutput = File.Create(tempPdfPath))
            {
                await _docxToPdf.ConvertAsync(docxInput, pdfOutput, cancellationToken).ConfigureAwait(false);
            }

            using var pdfInput = File.OpenRead(tempPdfPath);
            await _pdfToOfd.ConvertAsync(pdfInput, ofdOutput, pages, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            try
            {
                if (File.Exists(tempPdfPath))
                {
                    File.Delete(tempPdfPath);
                }
            }
            catch
            {
                // Best effort cleanup of the private PDF intermediate.
            }
        }
    }
}
