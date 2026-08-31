using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Docx.Internal;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
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

            if (_semanticOptions.OfdMode == DocxToOfdMode.Native)
            {
                var nativePackage = BuildNativePackageFromModel(tempDocxPath, pages, cancellationToken);
                await new OfdPackageWriter()
                    .WriteAsync(nativePackage, ofdOutput, cancellationToken)
                    .ConfigureAwait(false);
                return;
            }

            var sourceTextPages = DocxSemanticTextExtractor.ExtractPages(
                tempDocxPath,
                _semanticOptions,
                cancellationToken);

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

    private OfdDocumentPackage BuildNativePackageFromModel(
        string docxPath,
        IReadOnlyList<int>? selectedPages,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<DocxConversionDiagnostic>();
        var model = new DocxModelReader(_semanticOptions, diagnostics, cancellationToken).Read(docxPath);
        var package = new BuiltInOfdRenderer(_semanticOptions, diagnostics, cancellationToken).Render(model);
        if (selectedPages is null || selectedPages.Count == 0)
        {
            return package;
        }

        var selected = selectedPages
            .Where(index => index >= 0 && index < package.Pages.Count)
            .Distinct()
            .OrderBy(index => index)
            .Select(index => package.Pages[index])
            .ToList();
        package.Pages.Clear();
        for (var i = 0; i < selected.Count; i++)
        {
            selected[i].Index = i;
            package.Pages.Add(selected[i]);
        }

        if (package.Pages.Count == 0)
        {
            package.Pages.Add(new OfdPage
            {
                Index = 0,
                WidthMillimeters = package.Options.DefaultPageWidthMillimeters,
                HeightMillimeters = package.Options.DefaultPageHeightMillimeters
            });
        }

        return package;
    }
}
