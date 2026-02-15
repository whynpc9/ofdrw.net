using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;
using UglyToad.PdfPig;

namespace Ofdrw.Net.Converter.Pdf.Converters;

public sealed class PdfToOfdConverter : IPdfToOfdConverter
{
    private readonly PdfToPpmRasterizer _rasterizer = new();

    public async Task ConvertAsync(Stream pdfInput, Stream ofdOutput, IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default)
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
            using (var temp = File.Create(tempPdfPath))
            {
                await pdfInput.CopyToAsync(temp, 81920, cancellationToken).ConfigureAwait(false);
            }

            using var document = PdfDocument.Open(tempPdfPath);
            var selected = PageSelection.Normalize(document.NumberOfPages, pages);

            var builder = new OfdDocumentBuilder();
            builder.SetOptions(new OfdDocumentOptions
            {
                DocType = "OFD-H",
                DocumentId = "Doc_0",
                Namespace = "http://www.ofdspec.org",
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
                var pdfPage = document.GetPage(index + 1);
                var widthMm = PointsToMillimeters(pdfPage.Width);
                var heightMm = PointsToMillimeters(pdfPage.Height);

                var page = new OfdPage
                {
                    Index = outputPageIndex++,
                    WidthMillimeters = widthMm,
                    HeightMillimeters = heightMm
                };

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
                    var image = await _rasterizer.TryRasterizePageAsync(tempPdfPath, index, cancellationToken).ConfigureAwait(false);
                    if (image is not null)
                    {
                        page.Elements.Add(new OfdImageElement
                        {
                            XMillimeters = 0,
                            YMillimeters = 0,
                            WidthMillimeters = widthMm,
                            HeightMillimeters = heightMm,
                            Data = image,
                            MediaType = "image/png",
                            FileName = $"pdf_page_{index + 1}.png"
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

    private static double PointsToMillimeters(double points)
    {
        return points * 25.4d / 72d;
    }
}
