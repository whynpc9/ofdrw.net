using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Docx.Internal;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Docx.Converters;

/// <summary>
/// Converts DOCX/OpenXML directly to native OFD text objects by default. An optional
/// dual-layer mode uses PDF rendering only for the visual layer while keeping the
/// original DOCX/OpenXML as the machine-readable text source.
/// </summary>
public sealed class DocxToOfdConverter : IDocxToOfdConverter
{
    private readonly IDocxToPdfConverter _docxToPdf;
    private readonly IPdfToOfdConverter _pdfToOfd;
    private readonly DocxConversionOptions _semanticOptions;

    /// <summary>
    /// Initializes a converter using the default DOCX renderer and PDF converter.
    /// </summary>
    public DocxToOfdConverter()
        : this(new DocxConversionOptions())
    {
    }

    /// <summary>
    /// Initializes a converter using the supplied DOCX options.
    /// </summary>
    public DocxToOfdConverter(DocxConversionOptions options)
        : this(options, new PdfToOfdOptions())
    {
    }

    /// <summary>
    /// Initializes a converter using the supplied DOCX and PDF-to-OFD options.
    /// </summary>
    public DocxToOfdConverter(DocxConversionOptions options, PdfToOfdOptions pdfToOfdOptions)
        : this(
            new DocxToPdfConverter(options),
            new PdfToOfdConverter(CreateVisualLayerOptions(pdfToOfdOptions)),
            options)
    {
    }

    /// <summary>
    /// Initializes a converter using explicit pipeline stages.
    /// </summary>
    public DocxToOfdConverter(IDocxToPdfConverter docxToPdf, IPdfToOfdConverter pdfToOfd)
        : this(docxToPdf, pdfToOfd, CreateDualLayerOptions())
    {
    }

    private DocxToOfdConverter(
        IDocxToPdfConverter docxToPdf,
        IPdfToOfdConverter pdfToOfd,
        DocxConversionOptions semanticOptions)
    {
        _docxToPdf = docxToPdf ?? throw new ArgumentNullException(nameof(docxToPdf));
        _pdfToOfd = pdfToOfd ?? throw new ArgumentNullException(nameof(pdfToOfd));
        _semanticOptions = semanticOptions ?? throw new ArgumentNullException(nameof(semanticOptions));
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

        var workDirectory = Path.Combine(
            Path.GetTempPath(),
            $"ofdrw-docx-ofd-{Guid.NewGuid():N}");
        var tempDocxPath = Path.Combine(workDirectory, "input.docx");
        var tempPdfPath = Path.Combine(workDirectory, "visual.pdf");
        var tempOfdPath = Path.Combine(workDirectory, "visual.ofd");
        Directory.CreateDirectory(workDirectory);
        try
        {
            using (var stagedDocx = File.Create(tempDocxPath))
            {
                await docxInput.CopyToAsync(stagedDocx, 81920, cancellationToken).ConfigureAwait(false);
            }

            var sourceTextPages = DocxSemanticTextExtractor.ExtractPages(
                tempDocxPath,
                _semanticOptions,
                cancellationToken);
            if (_semanticOptions.OfdMode == DocxToOfdMode.Native)
            {
                var nativePackage = BuildNativePackage(sourceTextPages, pages);
                await new OfdPackageWriter()
                    .WriteAsync(nativePackage, ofdOutput, cancellationToken)
                    .ConfigureAwait(false);
                return;
            }

            using (var pdfOutput = File.Create(tempPdfPath))
            using (var stagedDocx = File.OpenRead(tempDocxPath))
            {
                await _docxToPdf.ConvertAsync(stagedDocx, pdfOutput, cancellationToken).ConfigureAwait(false);
            }

            using var pdfInput = File.OpenRead(tempPdfPath);
            using (var visualOfd = File.Create(tempOfdPath))
            {
                await _pdfToOfd.ConvertAsync(pdfInput, visualOfd, pages, cancellationToken)
                    .ConfigureAwait(false);
            }

            using var visualInput = File.OpenRead(tempOfdPath);
            var package = await new OfdReader()
                .ReadAsync(visualInput, cancellationToken)
                .ConfigureAwait(false);
            AddSourceTextLayer(package, sourceTextPages, pages);
            await new OfdPackageWriter()
                .WriteAsync(package, ofdOutput, cancellationToken)
                .ConfigureAwait(false);
        }
        finally
        {
            try
            {
                if (Directory.Exists(workDirectory))
                {
                    Directory.Delete(workDirectory, recursive: true);
                }
            }
            catch
            {
                // Best effort cleanup of the private conversion workspace.
            }
        }
    }

    private static DocxConversionOptions CreateDualLayerOptions()
    {
        return new DocxConversionOptions
        {
            OfdMode = DocxToOfdMode.DualLayer
        };
    }

    private static PdfToOfdOptions CreateVisualLayerOptions(PdfToOfdOptions options)
    {
        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        return new PdfToOfdOptions
        {
            Dpi = options.Dpi,
            RequireRasterization = options.RequireRasterization,
            PreferExternalPdfToPpm = options.PreferExternalPdfToPpm,
            MaxTextObjectsPerPage = options.MaxTextObjectsPerPage,
            TextLayerMode = PdfTextLayerMode.None
        };
    }

    private static void AddSourceTextLayer(
        OfdDocumentPackage package,
        IReadOnlyList<string> sourceTextPages,
        IReadOnlyList<int>? selectedPages)
    {
        foreach (var page in package.Pages)
        {
            page.Elements.RemoveAll(element =>
                element is OfdTextElement text &&
                text.FillColor.Alpha == 0 &&
                string.Equals(text.LayerType, "Foreground", StringComparison.OrdinalIgnoreCase));
        }

        package.CustomTags["source-text-origin"] = "DOCX/OpenXML";
        package.CustomTags["source-text-kind"] = "machine-readable";
        package.CustomTags["docx-ofd-mode"] = "DualLayer";

        var outputText = SelectSourceText(sourceTextPages, selectedPages, package.Pages.Count);
        for (var index = 0; index < package.Pages.Count && index < outputText.Count; index++)
        {
            var text = outputText[index];
            if (string.IsNullOrEmpty(text))
            {
                continue;
            }

            var page = package.Pages[index];
            page.Elements.Add(new OfdTextElement
            {
                LayerId = "docx-source-text",
                LayerType = "Foreground",
                XMillimeters = 0,
                YMillimeters = 0,
                WidthMillimeters = page.WidthMillimeters,
                HeightMillimeters = 1,
                FontName = "SimSun",
                FontSizeMillimeters = 1,
                FillColor = new OfdColor(0, 0, 0, 0),
                Text = text
            });
        }
    }

    private static IReadOnlyList<string> SelectSourceText(
        IReadOnlyList<string> sourceTextPages,
        IReadOnlyList<int>? selectedPages,
        int outputPageCount)
    {
        var selected = selectedPages is null || selectedPages.Count == 0
            ? sourceTextPages.ToList()
            : selectedPages
                .Where(index => index >= 0 && index < sourceTextPages.Count)
                .Select(index => sourceTextPages[index])
                .ToList();

        while (selected.Count < outputPageCount)
        {
            selected.Add(string.Empty);
        }

        if (selected.Count > outputPageCount && outputPageCount > 0)
        {
            var overflow = string.Join("\n", selected.Skip(outputPageCount - 1));
            selected.RemoveRange(outputPageCount - 1, selected.Count - outputPageCount + 1);
            selected.Add(overflow);
        }

        return selected;
    }

    private static OfdDocumentPackage BuildNativePackage(
        IReadOnlyList<string> sourceTextPages,
        IReadOnlyList<int>? selectedPages)
    {
        var selected = selectedPages is null || selectedPages.Count == 0
            ? sourceTextPages.ToList()
            : selectedPages
                .Where(index => index >= 0 && index < sourceTextPages.Count)
                .Select(index => sourceTextPages[index])
                .ToList();
        if (selected.Count == 0)
        {
            selected.Add(string.Empty);
        }

        var package = new OfdDocumentPackage
        {
            Options = new OfdDocumentOptions
            {
                DocType = "OFD-H",
                DocumentId = "Doc_0",
                Metadata = new OfdMetadata
                {
                    Title = "DOCX document",
                    Creator = "Ofdrw.Net native DOCX converter",
                    CreationDate = DateTimeOffset.UtcNow,
                    ModificationDate = DateTimeOffset.UtcNow
                }
            }
        };
        package.CustomTags["source-text-origin"] = "DOCX/OpenXML";
        package.CustomTags["source-text-kind"] = "machine-readable";
        package.CustomTags["docx-ofd-mode"] = "Native";

        foreach (var sourcePage in selected)
        {
            AddNativeTextPages(package, sourcePage);
        }

        return package;
    }

    private static void AddNativeTextPages(OfdDocumentPackage package, string sourceText)
    {
        const double pageWidth = 210d;
        const double pageHeight = 297d;
        const double margin = 20d;
        const double fontSize = 3.5d;
        const double lineHeight = 5d;
        var contentWidth = pageWidth - (margin * 2);
        var lines = WrapSourceText(sourceText, contentWidth, fontSize).ToList();
        if (lines.Count == 0)
        {
            lines.Add(string.Empty);
        }

        OfdPage? page = null;
        var y = margin;
        foreach (var line in lines)
        {
            if (page is null || y + lineHeight > pageHeight - margin)
            {
                page = new OfdPage
                {
                    Index = package.Pages.Count,
                    WidthMillimeters = pageWidth,
                    HeightMillimeters = pageHeight
                };
                package.Pages.Add(page);
                y = margin;
            }

            if (!string.IsNullOrEmpty(line))
            {
                page.Elements.Add(new OfdTextElement
                {
                    LayerType = "Body",
                    XMillimeters = margin,
                    YMillimeters = y,
                    WidthMillimeters = contentWidth,
                    HeightMillimeters = lineHeight,
                    FontName = "SimSun",
                    FontSizeMillimeters = fontSize,
                    FillColor = OfdColor.Black,
                    Text = line
                });
            }

            y += lineHeight;
        }
    }

    private static IEnumerable<string> WrapSourceText(
        string sourceText,
        double availableWidth,
        double fontSize)
    {
        var normalized = (sourceText ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
        foreach (var paragraphLine in normalized.Split('\n'))
        {
            if (paragraphLine.Length == 0)
            {
                yield return string.Empty;
                continue;
            }

            var line = new StringBuilder();
            var width = 0d;
            var enumerator = StringInfo.GetTextElementEnumerator(paragraphLine);
            while (enumerator.MoveNext())
            {
                var element = enumerator.GetTextElement();
                var elementWidth = EstimateTextWidth(element, fontSize);
                if (line.Length > 0 && width + elementWidth > availableWidth)
                {
                    yield return line.ToString();
                    line.Clear();
                    width = 0;
                }

                line.Append(element);
                width += elementWidth;
            }

            yield return line.ToString();
        }
    }

    private static double EstimateTextWidth(string textElement, double fontSize)
    {
        if (string.IsNullOrWhiteSpace(textElement))
        {
            return fontSize * 0.35d;
        }

        return textElement.Length == 1 && textElement[0] <= 0x7f
            ? fontSize * 0.55d
            : fontSize;
    }
}
