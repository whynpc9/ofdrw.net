using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using UglyToad.PdfPig;

namespace Ofdrw.Net.Converter.Pdf.Vector;

/// <summary>Explicit optional PDF vector mode with whole-page dual-layer fallback.</summary>
/// <remarks>Only the documented fixed-font, opaque path/text profile is supported. No canvas calls are intercepted.</remarks>
public sealed class PdfVectorToOfdConverter : IPdfToOfdConverter
{
    private readonly PdfVectorToOfdOptions _options;
    /// <summary>Creates the optional converter with explicit failure for unsupported pages.</summary>
    public PdfVectorToOfdConverter() : this(new PdfVectorToOfdOptions()) { }
    /// <summary>Creates the optional converter; options are snapshotted for each conversion.</summary>
    public PdfVectorToOfdConverter(PdfVectorToOfdOptions options) => _options = options ?? throw new ArgumentNullException(nameof(options));
    /// <inheritdoc />
    public async Task ConvertAsync(Stream pdfInput, Stream ofdOutput, IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default)
        => _ = await ConvertWithResultAsync(pdfInput, ofdOutput, pages, cancellationToken).ConfigureAwait(false);

    /// <summary>Converts and reports each actual native or raster page decision.</summary>
    public async Task<PdfVectorConversionResult> ConvertWithResultAsync(Stream pdfInput, Stream ofdOutput,
        IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default)
    {
        if (pdfInput is null) throw new ArgumentNullException(nameof(pdfInput));
        if (ofdOutput is null) throw new ArgumentNullException(nameof(ofdOutput));
        var limits = VectorLimits.Snapshot(_options);
        var path = Path.Combine(Path.GetTempPath(), "ofdrw-vector-" + Guid.NewGuid().ToString("N") + ".pdf");
        try
        {
            using (var staged = File.Create(path))
                await BoundedStreamCopy.CopyAsync(pdfInput, staged, limits.Compatibility.MaxInputBytes, "PDF input", cancellationToken).ConfigureAwait(false);
            int count;
            using (var source = PdfDocument.Open(path)) count = source.NumberOfPages;
            if (count <= 0 || count > limits.Compatibility.MaxPageCount) throw new InvalidDataException("PDF source page limit exceeded or document has no pages.");
            var selection = OfdPageSelection.Normalize(count, pages);
            if (selection.Count > limits.Compatibility.MaxPageCount) throw new InvalidDataException("PDF output page limit exceeded.");
            var package = new OfdDocumentPackage();
            package.Options.Namespace = limits.Compatibility.Namespace;
            package.Options.Metadata.Title = "PDF vector document";
            package.Options.Metadata.Creator = "Ofdrw.Net PdfVectorToOfdConverter";
            var results = new List<PdfVectorPageResult>();
            long imageBytes = 0, fontBytes = 0;
            foreach (var index in selection)
            {
                cancellationToken.ThrowIfCancellationRequested();
                OfdPage page; string diagnostic; bool native;
                try
                {
                    // Discard each parser/resource store after the page, including every rejection path.
                    using var document = PdfDocument.Open(path, new ParsingOptions { UseLenientParsing = false, MaxStackDepth = limits.MaxStackDepth });
                    if (!PlainCatalog(document)) throw new NotSupportedException("PDFV_UNSUPPORTED: catalog color/optional-content/rendering extensions.");
                    document.AddPageFactory<VectorPageContext, VectorContextFactory>();
                    var context = document.GetPage<VectorPageContext>(index + 1);
                    var staged = context.Convert(limits, cancellationToken);
                    page = staged.Pages.Single();
                    var bytes = staged.Fonts.Sum(font => (long)font.Data!.Length);
                    if (bytes > limits.MaxTotalFontBytes - fontBytes) throw new InvalidDataException("PDF accumulated font byte limit exceeded.");
                    fontBytes += bytes;
                    // Give every staged logical handle an ordinal unique package identity.
                    foreach (var font in staged.Fonts)
                    {
                        var old = font.Id; font.Id = "vector-" + package.Pages.Count + "-" + old;
                        foreach (var text in page.Elements.OfType<OfdTextElement>().Where(text => text.FontResourceId == old)) text.FontResourceId = font.Id;
                        package.Fonts.Add(font);
                    }
                    native = true; diagnostic = "PDFV_NATIVE: bounded opaque paths and exact-font Unicode text.";
                }
                catch (NotSupportedException exception) when (limits.UnsupportedPagePolicy == PdfVectorUnsupportedPagePolicy.RasterizePage)
                {
                    OfdPage? originalImage = null;
                    using (var source = PdfDocument.Open(path, new ParsingOptions { UseLenientParsing = false, MaxStackDepth = limits.MaxStackDepth }))
                    {
                        if (PlainCatalog(source))
                        {
                            source.AddPageFactory<VectorPageContext, VectorContextFactory>();
                            originalImage = source.GetPage<VectorPageContext>(index + 1).TryOriginalImagePage(limits,
                                limits.Compatibility.MaxTotalImageBytes - imageBytes, cancellationToken);
                        }
                    }
                    if (originalImage is not null)
                    {
                        page = originalImage;
                        imageBytes += page.Elements.OfType<OfdImageElement>().Sum(image => (long)image.Data.Length);
                        native = false;
                        diagnostic = "PDFV_ORIGINAL_IMAGE_PAGE: exact full-page raw RGB or literal Indexed/DeviceRGB palette samples/resolution and source Interpolate preserved; raster content, no OCR.";
                    }
                    else
                    {
                        // A rejected page contributes no staged vector content or fonts.
                        var fallbackOptions = VectorLimits.CopyCompatibility(limits.Compatibility);
                        fallbackOptions.RequireRasterization = true;
                        fallbackOptions.MaxTotalImageBytes = limits.Compatibility.MaxTotalImageBytes - imageBytes;
                        if (fallbackOptions.MaxTotalImageBytes <= 0) throw new InvalidDataException("PDF accumulated image byte limit exceeded.");
                        using var input = File.OpenRead(path); using var output = new MemoryStream();
                        var fallback = await new PdfToOfdConverter(fallbackOptions).ConvertWithResultAsync(input, output, new[] { index }, cancellationToken).ConfigureAwait(false);
                        output.Position = 0;
                        var rasterPackage = await new OfdReader().ReadAsync(output, cancellationToken).ConfigureAwait(false);
                        page = OfdModelCloner.ClonePage(rasterPackage.Pages.Single(), keepSourcePath: false, clonePayloads: false);
                        foreach (var element in page.Elements) element.ObjectId = null;
                        imageBytes += page.Elements.OfType<OfdImageElement>().Sum(image => (long)image.Data!.Length);
                        if (imageBytes > limits.Compatibility.MaxTotalImageBytes) throw new InvalidDataException("PDF accumulated image byte limit exceeded.");
                        // Default dual-layer text has no embedded font. Re-register by name at final write.
                        foreach (var text in page.Elements.OfType<OfdTextElement>()) { text.SourceXml = null; text.FontResourceId = null; }
                        native = false; diagnostic = "PDFV_RASTER_PAGE: " + exception.Message + "; " + string.Join("; ", fallback.Diagnostics);
                    }
                }
                page.Index = package.Pages.Count; package.Pages.Add(page);
                results.Add(new PdfVectorPageResult(index, native, page.Elements.OfType<OfdPathElement>().Count(),
                    page.Elements.OfType<OfdTextElement>().Count(), page.Elements.OfType<OfdImageElement>().Count(), diagnostic));
            }
            cancellationToken.ThrowIfCancellationRequested();
            await new OfdPackageWriter().WriteAsync(package, ofdOutput, cancellationToken).ConfigureAwait(false);
            return new PdfVectorConversionResult(count, results.ToArray());
        }
        finally { try { File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }
    private static bool PlainCatalog(PdfDocument document)
    {
        var dictionary = document.Structure.Catalog.CatalogDictionary;
        return !dictionary.Data.ContainsKey("OutputIntents") && !dictionary.Data.ContainsKey("OCProperties") &&
            !dictionary.Data.ContainsKey("Extensions") && !dictionary.Data.ContainsKey("Alternates");
    }
}

internal static class VectorLimits
{
    internal static PdfVectorToOfdOptions Snapshot(PdfVectorToOfdOptions value)
    {
        if (value.Compatibility is null) throw new ArgumentException("Compatibility limits are required.");
        var copy = new PdfVectorToOfdOptions { UnsupportedPagePolicy = value.UnsupportedPagePolicy, Compatibility = CopyCompatibility(value.Compatibility),
            MaxContentBytesPerPage = value.MaxContentBytesPerPage, MaxOperationsPerPage = value.MaxOperationsPerPage,
            MaxEventsPerPage = value.MaxEventsPerPage, MaxPathCommandsPerPage = value.MaxPathCommandsPerPage,
            MaxTextCharactersPerPage = value.MaxTextCharactersPerPage, MaxFontBytes = value.MaxFontBytes,
            MaxTotalFontBytes = value.MaxTotalFontBytes, MaxFontsPerPage = value.MaxFontsPerPage, MaxStackDepth = value.MaxStackDepth };
        if (copy.UnsupportedPagePolicy is not (PdfVectorUnsupportedPagePolicy.Fail or PdfVectorUnsupportedPagePolicy.RasterizePage) ||
            copy.MaxContentBytesPerPage <= 0 || copy.MaxOperationsPerPage <= 0 || copy.MaxEventsPerPage <= 0 || copy.MaxPathCommandsPerPage <= 0 ||
            copy.MaxTextCharactersPerPage <= 0 || copy.MaxFontBytes <= 0 || copy.MaxTotalFontBytes <= 0 || copy.MaxFontsPerPage <= 0 || copy.MaxStackDepth <= 0)
            throw new ArgumentOutOfRangeException(nameof(value), "Vector limits must be positive and policy defined.");
        _ = new PdfToOfdConverter(copy.Compatibility); // Validate the existing compatibility limits.
        if (string.IsNullOrWhiteSpace(copy.Compatibility.Namespace)) throw new ArgumentException("Namespace must be specified.");
        return copy;
    }
    internal static PdfToOfdOptions CopyCompatibility(PdfToOfdOptions value) => new()
    { MaxInputBytes = value.MaxInputBytes, MaxPageCount = value.MaxPageCount, MaxRasterizedPixelsPerPage = value.MaxRasterizedPixelsPerPage,
        MaxTotalImageBytes = value.MaxTotalImageBytes, ExternalRasterizationTimeout = value.ExternalRasterizationTimeout, Dpi = value.Dpi,
        RequireRasterization = value.RequireRasterization, PreferExternalPdfToPpm = value.PreferExternalPdfToPpm,
        TextLayerMode = value.TextLayerMode, MaxTextObjectsPerPage = value.MaxTextObjectsPerPage, Namespace = value.Namespace };
}
