using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Docx.Internal;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.IO;
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
    public async Task ConvertAsync(Stream docxInput, Stream ofdOutput, IReadOnlyList<int>? pages = null,
        CancellationToken cancellationToken = default)
    {
        _ = await ConvertWithResultAsync(docxInput, ofdOutput, pages, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts original OpenXML text and reports page selection and any preservation limitations.</summary>
    public async Task<DocxToOfdConversionResult> ConvertWithResultAsync(
        Stream docxInput, Stream ofdOutput, IReadOnlyList<int>? pages = null,
        CancellationToken cancellationToken = default)
    {
        if (docxInput is null) throw new ArgumentNullException(nameof(docxInput));
        if (ofdOutput is null) throw new ArgumentNullException(nameof(ofdOutput));
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-docx-ofd-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var inputPath = Path.Combine(directory, "input.docx");
        var pdfPath = Path.Combine(directory, "visual.pdf");
        try
        {
            using (var input = File.Create(inputPath))
                await BoundedStreamCopy.CopyAsync(docxInput, input, _semanticOptions.MaxInputBytes,
                    "DOCX input", cancellationToken).ConfigureAwait(false);
            if (_semanticOptions.OfdMode == DocxToOfdMode.Native)
            {
                var diagnostics = new List<DocxConversionDiagnostic>();
                OfdDocumentPackage package;
                using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                {
                    timeout.CancelAfter(_semanticOptions.ProcessTimeout);
                    try
                    {
                        var model = new DocxModelReader(_semanticOptions, diagnostics, timeout.Token).Read(inputPath);
                        package = new BuiltInOfdRenderer(_semanticOptions, diagnostics, timeout.Token).Render(model);
                    }
                    catch (OperationCanceledException exception) when (!cancellationToken.IsCancellationRequested)
                    {
                        throw new TimeoutException("Native DOCX conversion exceeded the configured timeout.", exception);
                    }
                }
                var sourcePageCount = package.Pages.Count;
                var selected = OfdPageSelection.Apply(package, pages, _semanticOptions.MaxPageCount);
                await new OfdPackageWriter().WriteAsync(package, ofdOutput, cancellationToken).ConfigureAwait(false);
                return new DocxToOfdConversionResult(DocxToOfdMode.Native, null, new DocxConversionEngine[0], diagnostics,
                    sourcePageCount, selected, !diagnostics.Any(diagnostic => diagnostic.Code.Contains("UNSUPPORTED")), true);
            }

            var source = DocxSemanticTextExtractor.ExtractDocument(inputPath, _semanticOptions, cancellationToken);
            DocxConversionResult? renderResult = null;
            using (var input = File.OpenRead(inputPath))
            using (var output = File.Create(pdfPath))
            {
                if (_docxToPdf is DocxToPdfConverter renderer)
                    renderResult = await renderer.ConvertWithResultAsync(input, output, cancellationToken).ConfigureAwait(false);
                else await _docxToPdf.ConvertAsync(input, output, cancellationToken).ConfigureAwait(false);
            }
            var mapping = DocxPdfTextMapper.Map(source, pdfPath, _semanticOptions.MaxPageCount,
                renderResult?.ActualEngine == DocxConversionEngine.BuiltIn, cancellationToken);
            if ((mapping.HasUnplacedSupplementalText || mapping.HasUnplacedBodyText) && pages is { Count: > 0 })
                throw new InvalidDataException("DOCX_PAGE_TEXT_SCOPE_AMBIGUOUS: page selection requires a reliable anchor for every original text part.");
            var sourcePages = OfdPageSelection.Normalize(mapping.Pages.Count, pages);
            if (sourcePages.Count > _semanticOptions.MaxPageCount)
                throw new InvalidDataException("DOCX selection exceeds the configured page count limit.");
            OfdDocumentPackage visual;
            var allDiagnostics = renderResult?.Diagnostics.ToList() ?? new List<DocxConversionDiagnostic>();
            allDiagnostics.AddRange(mapping.Diagnostics);
            if (_pdfToOfd is PdfToOfdConverter builtIn)
            {
                var converted = await builtIn.ConvertFileToPackageAsync(pdfPath, sourcePages, cancellationToken).ConfigureAwait(false);
                visual = converted.Package;
                foreach (var diagnostic in converted.Result.Diagnostics)
                    allDiagnostics.Add(new DocxConversionDiagnostic("DOCX_VISUAL_RASTERIZER", diagnostic, DocxConversionDiagnosticSeverity.Information));
            }
            else
            {
                var visualPath = Path.Combine(directory, "visual.ofd");
                using (var input = File.OpenRead(pdfPath))
                using (var output = File.Create(visualPath))
                    await _pdfToOfd.ConvertAsync(input, output, sourcePages, cancellationToken).ConfigureAwait(false);
                using var visualInput = File.OpenRead(visualPath);
                visual = await new OfdReader().ReadAsync(visualInput, cancellationToken).ConfigureAwait(false);
            }
            if (visual.Pages.Count != sourcePages.Count)
                throw new InvalidDataException("The visual converter did not preserve the requested page selection.");
            ApplyDocxPackageProfile(visual);
            AddSourceTextLayer(visual, mapping, sourcePages);
            await new OfdPackageWriter().WriteAsync(visual, ofdOutput, cancellationToken).ConfigureAwait(false);
            return new DocxToOfdConversionResult(DocxToOfdMode.DualLayer, renderResult?.ActualEngine,
                renderResult?.AttemptedEngines ?? new DocxConversionEngine[0], allDiagnostics,
                mapping.Pages.Count, sourcePages, true, !mapping.HasUnplacedSupplementalText && !mapping.HasUnplacedBodyText);
        }
        finally
        {
            try { Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static DocxConversionOptions CreateDualLayerOptions() => new() { OfdMode = DocxToOfdMode.DualLayer };

    private static PdfToOfdOptions CreateVisualLayerOptions(PdfToOfdOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        return new PdfToOfdOptions
        {
            // A DOCX visual stage must not introduce PDF-derived text as a
            // fallback: the semantic layer always comes from OpenXML.
            Dpi = options.Dpi, RequireRasterization = true,
            PreferExternalPdfToPpm = options.PreferExternalPdfToPpm, MaxTextObjectsPerPage = options.MaxTextObjectsPerPage,
            MaxInputBytes = options.MaxInputBytes, MaxPageCount = options.MaxPageCount,
            MaxRasterizedPixelsPerPage = options.MaxRasterizedPixelsPerPage, MaxTotalImageBytes = options.MaxTotalImageBytes,
            ExternalRasterizationTimeout = options.ExternalRasterizationTimeout, TextLayerMode = PdfTextLayerMode.None,
            Namespace = options.Namespace
        };
    }

    /// <summary>
    /// Gives the DualLayer package the same profile as Native output. The visual
    /// stage is a PDF converter whose defaults (GB/T 33190 "/2016" namespace,
    /// "PDF document" metadata) describe a PDF import, but the document is a
    /// DOCX conversion and must open in the same readers as Native output. OFD-H
    /// readers that only resolve the short namespace otherwise drop the page image
    /// and show a blank page.
    /// </summary>
    private void ApplyDocxPackageProfile(OfdDocumentPackage visual)
    {
        var targetNamespace = string.IsNullOrWhiteSpace(_semanticOptions.OfdNamespace)
            ? OfdConstants.Namespace : _semanticOptions.OfdNamespace;
        var sourceNamespace = XNamespace.Get(visual.Options.Namespace);
        if (sourceNamespace.NamespaceName != targetNamespace)
        {
            // A staged OFD read back through OfdReader keeps its resource manifests
            // and unknown elements as raw XML. The writer updates manifests in
            // place, so drop them and remap the raw fragments; payloads are
            // re-emitted from the model under the same content-hash names.
            foreach (var entry in visual.PreservedEntries.Where(pair => IsResourceManifest(pair.Value)).Select(pair => pair.Key).ToList())
                visual.PreservedEntries.Remove(entry);
            RemapNamespace(visual.PreservedDocBodyElements, sourceNamespace, targetNamespace);
            RemapNamespace(visual.PreservedCommonDataElements, sourceNamespace, targetNamespace);
            RemapNamespace(visual.PreservedDocumentElements, sourceNamespace, targetNamespace);
            foreach (var page in visual.Pages)
            {
                RemapNamespace(page.PreservedPageElements, sourceNamespace, targetNamespace);
                foreach (var element in page.Elements.Concat(page.Templates.SelectMany(template => template.Elements)))
                {
                    switch (element)
                    {
                        case OfdImageElement image:
                            image.SourceXml = RemapNamespace(image.SourceXml, sourceNamespace, targetNamespace);
                            image.ClipsXml = RemapNamespace(image.ClipsXml, sourceNamespace, targetNamespace);
                            break;
                        case OfdTextElement text:
                            text.SourceXml = RemapNamespace(text.SourceXml, sourceNamespace, targetNamespace);
                            break;
                        case OfdPathElement path:
                            path.SourceXml = RemapNamespace(path.SourceXml, sourceNamespace, targetNamespace);
                            break;
                        case OfdRawElement raw:
                            raw.Xml = RemapNamespace(raw.Xml, sourceNamespace, targetNamespace) ?? string.Empty;
                            break;
                    }
                }
            }
            visual.Options.Namespace = targetNamespace;
        }
        visual.Options.DocType = OfdConstants.DefaultDocType;
        visual.Options.Metadata.Title = "DOCX document";
        visual.Options.Metadata.Creator = "Ofdrw.Net DocxToOfdConverter";
    }

    private static bool IsResourceManifest(byte[] entry)
    {
        if (entry.Length == 0 || entry.Length > 16 * 1024 * 1024) return false;
        try
        {
            using var stream = new MemoryStream(entry, writable: false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            return reader.MoveToContent() == XmlNodeType.Element && reader.LocalName == "Res";
        }
        catch (XmlException) { return false; }
    }

    private static void RemapNamespace(IList<string> fragments, XNamespace source, string target)
    {
        for (var index = 0; index < fragments.Count; index++)
            fragments[index] = RemapNamespace(fragments[index], source, target) ?? fragments[index];
    }

    private static string? RemapNamespace(string? fragment, XNamespace source, string target)
    {
        if (string.IsNullOrWhiteSpace(fragment)) return fragment;
        var root = XElement.Parse(fragment!, LoadOptions.PreserveWhitespace);
        foreach (var node in root.DescendantsAndSelf())
        {
            if (node.Name.Namespace == source) node.Name = XName.Get(node.Name.LocalName, target);
            node.Attributes().Where(attribute => attribute.IsNamespaceDeclaration && attribute.Value == source.NamespaceName).Remove();
        }
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static void AddSourceTextLayer(OfdDocumentPackage package, DocxPageTextMap textMap, IReadOnlyList<int> sourcePages)
    {
        package.CustomTags["source-text-origin"] = "DOCX/OpenXML";
        package.CustomTags["source-text-kind"] = "machine-readable";
        package.CustomTags["docx-ofd-mode"] = "DualLayer";
        package.CustomTags["source-text-page-map"] = textMap.HasUnplacedBodyText
            ? "includes-unanchored-body-text"
            : textMap.HasUnplacedSupplementalText ? "includes-document-scoped-supplementary-text" : "PDF-position-aligned/OpenXML-content";
        for (var index = 0; index < package.Pages.Count; index++)
        {
            var page = package.Pages[index];
            if (page.Elements.Concat(page.Templates.SelectMany(template => template.Elements))
                .OfType<OfdTextElement>().Any(text => text.FillColor.Alpha != 0))
                throw new InvalidDataException("The DualLayer visual converter must rasterize text; visible PDF-derived text is not an original DOCX semantic source.");
            page.Elements.RemoveAll(element => element is OfdTextElement text && text.FillColor.Alpha == 0 &&
                string.Equals(text.LayerType, "Foreground", StringComparison.OrdinalIgnoreCase));
            var originalText = textMap.Pages[sourcePages[index]];
            if (string.IsNullOrEmpty(originalText)) continue;
            page.Elements.Add(new OfdTextElement
            {
                LayerId = "docx-source-text", LayerType = "Foreground", WidthMillimeters = page.WidthMillimeters,
                HeightMillimeters = 1, FontName = "SimSun", FontSizeMillimeters = 1,
                FillColor = new OfdColor(0, 0, 0, 0), Text = originalText
            });
        }
    }
}
