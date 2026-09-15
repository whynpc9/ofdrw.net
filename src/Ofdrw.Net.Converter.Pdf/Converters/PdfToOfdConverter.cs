using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Packaging;
using UglyToad.PdfPig;

namespace Ofdrw.Net.Converter.Pdf.Converters;

/// <summary>
/// Converts PDF to dual-layer OFD. By default each page is rasterized
/// (Docnet/Pdfium, with optional <c>pdftoppm</c>) for visual fidelity and
/// extractable PDF words are retained as transparent OFD text objects.
/// </summary>
public sealed class PdfToOfdConverter : IPdfToOfdConverter
{
    private readonly PdfToOfdOptions _options;

    /// <summary>
    /// Initializes a converter with default options (page rasterization required).
    /// </summary>
    public PdfToOfdConverter()
        : this(new PdfToOfdOptions())
    {
    }

    /// <summary>
    /// Initializes a converter with the supplied options.
    /// </summary>
    public PdfToOfdConverter(PdfToOfdOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        if (_options.MaxTextObjectsPerPage <= 0 || _options.MaxInputBytes <= 0 ||
            _options.MaxPageCount <= 0 || _options.MaxRasterizedPixelsPerPage <= 0 ||
            _options.MaxTotalImageBytes <= 0 || _options.ExternalRasterizationTimeout <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(
                nameof(options),
                "PDF conversion limits and external timeout must be greater than zero.");
        }
    }

    /// <inheritdoc />
    public async Task ConvertAsync(
        Stream pdfInput, Stream ofdOutput, IReadOnlyList<int>? pages = null,
        CancellationToken cancellationToken = default)
    {
        _ = await ConvertWithResultAsync(pdfInput, ofdOutput, pages, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts a PDF and returns the actual page selection and rasterizer diagnostics.</summary>
    public async Task<PdfConversionResult> ConvertWithResultAsync(
        Stream pdfInput, Stream ofdOutput, IReadOnlyList<int>? pages = null,
        CancellationToken cancellationToken = default)
    {
        if (pdfInput is null) throw new ArgumentNullException(nameof(pdfInput));
        if (ofdOutput is null) throw new ArgumentNullException(nameof(ofdOutput));
        var tempPdfPath = Path.Combine(Path.GetTempPath(), $"ofdrw-net-{Guid.NewGuid():N}.pdf");
        try
        {
            using (var staged = File.Create(tempPdfPath))
                await BoundedStreamCopy.CopyAsync(pdfInput, staged, _options.MaxInputBytes, "PDF input", cancellationToken).ConfigureAwait(false);
            var converted = await ConvertFileToPackageAsync(tempPdfPath, pages, cancellationToken).ConfigureAwait(false);
            await new OfdPackageWriter().WriteAsync(converted.Package, ofdOutput, cancellationToken).ConfigureAwait(false);
            return converted.Result;
        }
        finally
        {
            try { File.Delete(tempPdfPath); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    // The DOCX dual-layer pipeline can add original OpenXML text before the only
    // ZIP write, without staging a second PDF or serializing/re-reading an OFD.
    internal async Task<PdfToOfdPackageResult> ConvertFileToPackageAsync(
        string pdfPath, IReadOnlyList<int>? pages, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (new FileInfo(pdfPath).Length > _options.MaxInputBytes)
            throw new InvalidDataException("PDF input exceeds the configured input size limit.");
        using var document = PdfDocument.Open(pdfPath);
        if (document.NumberOfPages > _options.MaxPageCount)
            throw new InvalidDataException("PDF exceeds the configured source page limit.");
        var selected = OfdPageSelection.Normalize(document.NumberOfPages, pages);
        if (selected.Count > _options.MaxPageCount)
            throw new InvalidDataException("PDF selection exceeds the configured output page limit.");
        using var docnet = new DocnetPdfRasterizer(pdfPath, _options.Dpi, _options.MaxRasterizedPixelsPerPage);
        var external = new PdfToPpmRasterizer();
        var diagnostics = new List<string>();
        var package = new OfdDocumentPackage
        {
            Options = new OfdDocumentOptions
            {
                DocType = "OFD-H", DocumentId = "Doc_0", Namespace = OfdConstants.StandardNamespace,
                Metadata = new OfdMetadata
                {
                    Title = "PDF document", Creator = "Ofdrw.Net PdfToOfdConverter",
                    CreationDate = DateTimeOffset.UtcNow, ModificationDate = DateTimeOffset.UtcNow
                }
            }
        };
        long imageBytes = 0;
        var renderedPages = new Dictionary<int, byte[]>();
        foreach (var index in selected)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var pdfPage = document.GetPage(index + 1);
            ValidatePageDimensions(pdfPage.Width, pdfPage.Height);
            var page = new OfdPage
            {
                Index = package.Pages.Count, WidthMillimeters = PointsToMillimeters(pdfPage.Width),
                HeightMillimeters = PointsToMillimeters(pdfPage.Height)
            };
            if (!renderedPages.TryGetValue(index, out var image))
            {
                image = await TryRasterizePageAsync(pdfPath, docnet, external, index, diagnostics, cancellationToken).ConfigureAwait(false);
                if (image is not null)
                {
                    if (image.Length > _options.MaxTotalImageBytes - imageBytes)
                        throw new InvalidDataException("PDF exceeds the configured accumulated image byte limit.");
                    imageBytes += image.Length;
                    renderedPages.Add(index, image);
                }
            }
            if (image is not null)
            {
                page.Elements.Add(new OfdImageElement
                {
                    WidthMillimeters = page.WidthMillimeters, HeightMillimeters = page.HeightMillimeters,
                    Data = image, MediaType = "image/png", FileName = $"pdf_page_{index + 1}.png"
                });
                AddSemanticTextLayer(page, pdfPage);
            }
            else if (_options.RequireRasterization)
            {
                throw new InvalidOperationException($"Failed to rasterize PDF page {index + 1}. " + string.Join("; ", diagnostics));
            }
            else
            {
                AddTextFallbackPage(page, pdfPage, page.WidthMillimeters, page.HeightMillimeters, index);
                AddDiagnostic(diagnostics, $"Page {index + 1}: rendered as a text/placeholder fallback.");
            }
            package.Pages.Add(page);
        }
        return new PdfToOfdPackageResult(package, new PdfConversionResult(document.NumberOfPages, selected, diagnostics));
    }

    private void ValidatePageDimensions(double widthPoints, double heightPoints)
    {
        var dpi = Math.Max(72, Math.Min(_options.Dpi, 300));
        var width = Math.Ceiling(widthPoints * dpi / 72d);
        var height = Math.Ceiling(heightPoints * dpi / 72d);
        if (!(width > 0) || !(height > 0) || width > int.MaxValue || height > int.MaxValue ||
            width * height > _options.MaxRasterizedPixelsPerPage)
            throw new InvalidDataException("PDF page exceeds the configured decoded pixel limit or has invalid dimensions.");
    }

    private void AddSemanticTextLayer(
        OfdPage page,
        UglyToad.PdfPig.Content.Page pdfPage)
    {
        if (_options.TextLayerMode == PdfTextLayerMode.None)
        {
            return;
        }

        var words = pdfPage.GetWords()
            .Where(word => !string.IsNullOrWhiteSpace(word.Text))
            .ToList();
        if (words.Count > _options.MaxTextObjectsPerPage)
        {
            throw new InvalidDataException(
                $"PDF page {pdfPage.Number} contains {words.Count} text objects, " +
                $"which exceeds the configured limit of {_options.MaxTextObjectsPerPage}.");
        }

        foreach (var word in words)
        {
            var bounds = word.BoundingBox;
            var x = PointsToMillimeters(bounds.Left);
            var y = PointsToMillimeters(pdfPage.Height - bounds.Top);
            var width = Math.Max(PointsToMillimeters(bounds.Width), 0.1d);
            var height = Math.Max(PointsToMillimeters(bounds.Height), 0.1d);
            var fontSize = word.Letters.Count > 0
                ? PointsToMillimeters(word.Letters.Max(letter => letter.FontSize))
                : height;
            if (double.IsNaN(fontSize) || double.IsInfinity(fontSize) || fontSize <= 0)
            {
                fontSize = height;
            }

            page.Elements.Add(new OfdTextElement
            {
                LayerId = "semantic-text",
                LayerType = "Foreground",
                XMillimeters = x,
                YMillimeters = y,
                WidthMillimeters = width,
                HeightMillimeters = height,
                FontName = NormalizeFontName(word.FontName),
                FontSizeMillimeters = Math.Max(fontSize, 0.1d),
                FillColor = new OfdColor(0, 0, 0, 0),
                Text = word.Text + " "
            });
        }
    }

    private static string NormalizeFontName(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName))
        {
            return "SimSun";
        }

        var normalized = fontName!.Trim();
        var subsetSeparator = normalized.IndexOf('+');
        if (subsetSeparator == 6 && normalized
            .Substring(0, subsetSeparator)
            .All(character => character >= 'A' && character <= 'Z'))
        {
            normalized = normalized.Substring(subsetSeparator + 1);
        }

        return string.IsNullOrWhiteSpace(normalized) ? "SimSun" : normalized;
    }

    private async Task<byte[]?> TryRasterizePageAsync(
        string pdfPath, DocnetPdfRasterizer docnet, PdfToPpmRasterizer external,
        int pageIndex, IList<string> diagnostics, CancellationToken cancellationToken)
    {
        async Task<byte[]?> ExternalAsync()
        {
            try
            {
                var image = await external.TryRasterizePageAsync(pdfPath, pageIndex, cancellationToken,
                    _options.Dpi, _options.ExternalRasterizationTimeout).ConfigureAwait(false);
                if (image is null) AddDiagnostic(diagnostics, $"Page {pageIndex + 1}: pdftoppm failed: {external.LastFailure}");
                return image;
            }
            catch (TimeoutException exception)
            {
                AddDiagnostic(diagnostics, $"Page {pageIndex + 1}: {exception.Message}");
                return null;
            }
        }
        async Task<byte[]?> DocnetAsync()
        {
            var image = await docnet.TryRasterizePageAsync(pageIndex, cancellationToken).ConfigureAwait(false);
            if (image is null) AddDiagnostic(diagnostics, $"Page {pageIndex + 1}: PDFium unavailable: {docnet.LastFailure}");
            return image;
        }
        return _options.PreferExternalPdfToPpm
            ? await ExternalAsync().ConfigureAwait(false) ?? await DocnetAsync().ConfigureAwait(false)
            : await DocnetAsync().ConfigureAwait(false) ?? await ExternalAsync().ConfigureAwait(false);
    }

    private static void AddDiagnostic(IList<string> diagnostics, string message)
    {
        if (diagnostics.Count < 100) diagnostics.Add(message.Length > 4096 ? message.Substring(0, 4096) + " [truncated]" : message);
        else if (diagnostics.Count == 100) diagnostics.Add("Additional rasterizer diagnostics omitted.");
    }

    private static void AddTextFallbackPage(
        OfdPage page,
        UglyToad.PdfPig.Content.Page pdfPage,
        double widthMm,
        double heightMm,
        int index)
    {
        var words = string.Empty;
        try
        {
            words = string.Join(" ", pdfPage.GetWords().Select(x => x.Text).Where(x => !string.IsNullOrWhiteSpace(x)));
        }
        catch
        {
            words = string.Empty;
        }

        if (!string.IsNullOrWhiteSpace(words))
        {
            page.Elements.Add(new OfdTextElement
            {
                XMillimeters = 10,
                YMillimeters = 12,
                WidthMillimeters = Math.Max(widthMm - 20, 10),
                HeightMillimeters = Math.Max(heightMm - 20, 10),
                FontName = "SimSun",
                FontSizeMillimeters = 4,
                Text = words
            });
        }
        else
        {
            page.Elements.Add(new OfdTextElement
            {
                XMillimeters = 10,
                YMillimeters = 12,
                WidthMillimeters = Math.Max(widthMm - 20, 10),
                HeightMillimeters = 8,
                FontName = "SimSun",
                FontSizeMillimeters = 4,
                Text = $"[fallback] page {index + 1} rendered as placeholder"
            });
        }
    }

    private static double PointsToMillimeters(double points)
    {
        return points * 25.4d / 72d;
    }
}
