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
using Ofdrw.Net.Layout.Builders;
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
    private readonly PdfToPpmRasterizer _pdfToPpm = new();
    private readonly DocnetPdfRasterizer _docnet = new();

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
        if (_options.MaxTextObjectsPerPage <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(options),
                "MaxTextObjectsPerPage must be greater than zero.");
        }
    }

    /// <inheritdoc />
    public async Task ConvertAsync(
        Stream pdfInput,
        Stream ofdOutput,
        IReadOnlyList<int>? pages = null,
        CancellationToken cancellationToken = default)
    {
        if (pdfInput is null)
        {
            throw new ArgumentNullException(nameof(pdfInput));
        }

        if (ofdOutput is null)
        {
            throw new ArgumentNullException(nameof(ofdOutput));
        }

        var tempPdfPath = Path.Combine(Path.GetTempPath(), $"ofdrw-net-{Guid.NewGuid():N}.pdf");
        try
        {
            byte[] pdfBytes;
            using (var temp = File.Create(tempPdfPath))
            {
                await pdfInput.CopyToAsync(temp, 81920, cancellationToken).ConfigureAwait(false);
            }

            pdfBytes = File.ReadAllBytes(tempPdfPath);

            using var document = PdfDocument.Open(tempPdfPath);
            var selected = PageSelection.Normalize(document.NumberOfPages, pages);

            var builder = new OfdDocumentBuilder();
            builder.SetOptions(new OfdDocumentOptions
            {
                DocType = "OFD-H",
                DocumentId = "Doc_0",
                Namespace = OfdConstants.StandardNamespace,
                Metadata = new OfdMetadata
                {
                    Title = Path.GetFileName(tempPdfPath),
                    Creator = "Ofdrw.Net PdfToOfdConverter",
                    CreationDate = DateTimeOffset.UtcNow,
                    ModificationDate = DateTimeOffset.UtcNow
                }
            });

            var outputPageIndex = 0;
            foreach (var index in selected)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var pdfPage = document.GetPage(index + 1);
                var widthMm = PointsToMillimeters(pdfPage.Width);
                var heightMm = PointsToMillimeters(pdfPage.Height);

                var page = new OfdPage
                {
                    Index = outputPageIndex++,
                    WidthMillimeters = widthMm,
                    HeightMillimeters = heightMm
                };

                var image = await TryRasterizePageAsync(tempPdfPath, pdfBytes, index, cancellationToken)
                    .ConfigureAwait(false);
                if (image is not null)
                {
                    page.Elements.Add(new OfdImageElement
                    {
                        ObjectId = $"Img{index + 1}",
                        ResourceId = $"ResImg{index + 1}",
                        XMillimeters = 0,
                        YMillimeters = 0,
                        WidthMillimeters = widthMm,
                        HeightMillimeters = heightMm,
                        Data = image,
                        MediaType = "image/png",
                        FileName = $"pdf_page_{index + 1}.png"
                    });

                    AddSemanticTextLayer(page, pdfPage);

                    builder.AddPage(page);
                    continue;
                }

                if (_options.RequireRasterization)
                {
                    throw new InvalidOperationException(
                        $"Failed to rasterize PDF page {index + 1}. " +
                        "Ensure Docnet/Pdfium natives are available, or install poppler's pdftoppm.");
                }

                AddTextFallbackPage(page, pdfPage, widthMm, heightMm, index);
                builder.AddPage(page);
            }

            var writer = new OfdPackageWriter();
            await writer.WriteAsync(builder.Build(), ofdOutput, cancellationToken).ConfigureAwait(false);
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
                // ignored
            }
        }
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
        string pdfPath,
        byte[] pdfBytes,
        int zeroBasedPageIndex,
        CancellationToken cancellationToken)
    {
        if (_options.PreferExternalPdfToPpm)
        {
            var external = await _pdfToPpm
                .TryRasterizePageAsync(pdfPath, zeroBasedPageIndex, cancellationToken)
                .ConfigureAwait(false);
            if (external is not null)
            {
                return external;
            }

            return await _docnet
                .TryRasterizePageAsync(pdfBytes, zeroBasedPageIndex, _options.Dpi, cancellationToken)
                .ConfigureAwait(false);
        }

        var docnet = await _docnet
            .TryRasterizePageAsync(pdfBytes, zeroBasedPageIndex, _options.Dpi, cancellationToken)
            .ConfigureAwait(false);
        if (docnet is not null)
        {
            return docnet;
        }

        return await _pdfToPpm
            .TryRasterizePageAsync(pdfPath, zeroBasedPageIndex, cancellationToken)
            .ConfigureAwait(false);
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
