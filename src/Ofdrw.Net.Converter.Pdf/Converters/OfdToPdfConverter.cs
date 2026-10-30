using Ofdrw.Net.Core.Fonts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Layout.Internal.Flow;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using PdfSharpCore.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Processing;

namespace Ofdrw.Net.Converter.Pdf.Converters;

public sealed class OfdToPdfConverter : IOfdToPdfConverter
{
    private readonly OfdToPdfOptions _options;
    private readonly OfdGlyphMappingStrategy _glyphMappingStrategy;
    private readonly long _maximumFontParseWork;
    private readonly OfdNameOnlyFontBindings? _nameOnlyBindings;

    public OfdToPdfConverter() : this(new OfdToPdfOptions()) { }

    public OfdToPdfConverter(OfdToPdfOptions options) : this(options, FontParseBudget.DefaultMaximumWork) { }

    /// <summary>Creates an exporter with a cumulative finite cmap budget shared by document, fallback and signature fonts.</summary>
    /// <param name="options">PDF conversion options.</param>
    /// <param name="maximumCmapParseWork">Positive cumulative limit for cmap bytes, parsed records and entries inspected by variation coverage queries.</param>
    public OfdToPdfConverter(OfdToPdfOptions options, long maximumCmapParseWork)
    {
        if (maximumCmapParseWork <= 0) throw new ArgumentOutOfRangeException(nameof(maximumCmapParseWork));
        _maximumFontParseWork = maximumCmapParseWork;
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _nameOnlyBindings = options.NameOnlyFontBindings;
        OfdGlyphMappingStrategyGuard.Validate(options.GlyphMappingStrategy);
        _glyphMappingStrategy = options.GlyphMappingStrategy;
        if (options.PackageLoadOptions is null || options.MaxDecodedImagePixels <= 0)
            throw new ArgumentException("OFD load options and a positive image pixel budget are required.", nameof(options));
        PdfFontRegistry.EnsureInstalled();
    }

    public async Task ConvertAsync(Stream ofdInput, Stream pdfOutput, IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default)
    {
        if (ofdInput is null)
        {
            throw new ArgumentNullException(nameof(ofdInput));
        }

        if (pdfOutput is null)
        {
            throw new ArgumentNullException(nameof(pdfOutput));
        }

        var reader = new OfdReader();
        var package = await reader.ReadAsync(ofdInput, _options.PackageLoadOptions, cancellationToken).ConfigureAwait(false);
        await ConvertPackageAsync(package, pdfOutput, pages, cancellationToken).ConfigureAwait(false);
    }

    // Shared rendering implementation; image export has already loaded and budgeted its package.
    internal async Task ConvertPackageAsync(OfdDocumentPackage package, Stream pdfOutput,
        IReadOnlyList<int>? pages, CancellationToken cancellationToken, bool strictAppearanceBudgets = false, int? maximumSignatureAppearances = null)
    {
        var annotationMaterial = new OfdModelMaterialization(new OfdModelLimits());
        OfdCompositeGraph.Validate(package, annotationMaterial, cancellationToken);
        OfdAnnotationContract.Validate(package, annotationMaterial, cancellationToken);
        var orderedPages = package.Pages.OrderBy(x => x.Index).ToList();
        var selected = OfdPageSelection.Normalize(orderedPages.Count, pages, cancellationToken);
        HashSet<string>? selectedIds = null;
        if (strictAppearanceBudgets)
        {
            var occurrences = orderedPages.Where(page => !string.IsNullOrWhiteSpace(page.Id))
                .GroupBy(page => page.Id!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
            selectedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var index in selected)
            {
                var id = orderedPages[index].Id;
                if (string.IsNullOrWhiteSpace(id) || occurrences[id!] != 1)
                    throw new InvalidDataException("Selected OFD page ID must be nonblank and unique for signature filtering.");
                selectedIds.Add(id!);
            }
        }
        var parseBudget = new FontParseBudget(_maximumFontParseWork, cancellationToken);
        var fonts = new DocumentFontContext(package.Fonts, parseBudget: parseBudget, cancellationToken: cancellationToken,
            nameOnlyBindings: _nameOnlyBindings, glyphMappingStrategy: _glyphMappingStrategy);
        var signatureAppearances = await PrepareSignatureAppearancesAsync(
                package, selectedIds, strictAppearanceBudgets, maximumSignatureAppearances,
                annotationMaterial, parseBudget, cancellationToken)
            .ConfigureAwait(false);

        // PDFsharp font registrations are immutable and process-wide. Gather
        // compatible original-GID mappings before drawing the first page so
        // one embedded face can serve distinct TextObjects using that face.
        var aliasMaterial = new OfdModelMaterialization(new OfdModelLimits());
        fonts.PrepareExplicitGlyphAliases(ExplicitGlyphTexts(
            selected.Select(index => orderedPages[index]), aliasMaterial, cancellationToken));
        var preparedAppearanceFonts = new HashSet<DocumentFontContext>();
        foreach (var appearance in signatureAppearances)
        {
            if (appearance.OfdPage is null || appearance.Fonts is null ||
                !selected.Any(index => string.Equals(orderedPages[index].Id, appearance.PageId,
                    StringComparison.OrdinalIgnoreCase)) ||
                !preparedAppearanceFonts.Add(appearance.Fonts)) continue;
            appearance.Fonts.PrepareExplicitGlyphAliases(ExplicitGlyphTexts(
                new[] { appearance.OfdPage }, aliasMaterial, cancellationToken));
        }

        using var document = new PdfDocument();
        foreach (var selectedIndex in selected)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var pageModel = orderedPages[selectedIndex];
            if (!(pageModel.WidthMillimeters > 0) || !(pageModel.HeightMillimeters > 0) ||
                double.IsInfinity(pageModel.WidthMillimeters) || double.IsInfinity(pageModel.HeightMillimeters))
                throw new InvalidDataException("OFD page has invalid physical dimensions.");
            if (strictAppearanceBudgets && !PdfOperandGeometry.Origin(pageModel.XMillimeters, pageModel.YMillimeters))
                throw new InvalidDataException("Selected OFD page origin must remain finite in PDF coordinates.");
            var pdfPage = document.AddPage();
            pdfPage.Width = MillimetersToPoints(pageModel.WidthMillimeters);
            pdfPage.Height = MillimetersToPoints(pageModel.HeightMillimeters);

            using var graphics = XGraphics.FromPdfPage(pdfPage);
            DrawPage(graphics, pageModel, fonts, _options.MaxDecodedImagePixels, annotationMaterial, cancellationToken);

            DrawSignatureAppearances(
                document,
                graphics,
                pageModel,
                signatureAppearances, _options.MaxDecodedImagePixels, annotationMaterial, cancellationToken, strictAppearanceBudgets);
        }

        cancellationToken.ThrowIfCancellationRequested();
        PdfSupplementaryFontBinding.InstallToUnicodeMaps(document,
            fonts.SupplementaryBindings.Concat(signatureAppearances
                .Where(appearance => appearance.Fonts is not null)
                .SelectMany(appearance => appearance.Fonts!.SupplementaryBindings)));
        var explicitUnicodeBindings =
            new[] { fonts }.Concat(signatureAppearances
                    .Where(appearance => appearance.Fonts is not null)
                    .Select(appearance => appearance.Fonts!))
                .SelectMany(context => context.ExplicitUnicodeBindings).ToArray();
        PdfExplicitGlyphUnicodeMap.Install(document, explicitUnicodeBindings);
        var selectedVerifiedFaces = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var context in new[] { fonts }.Concat(signatureAppearances
                     .Where(appearance => appearance.Fonts is not null)
                     .Select(appearance => appearance.Fonts!)))
        {
            foreach (var face in context.SelectedVerifiedFaces)
                selectedVerifiedFaces[face.Key] = face.Value;
        }
        if (selectedVerifiedFaces.Count == 0)
        {
            document.Save(pdfOutput, false);
        }
        else
        {
            // Keep the caller's output untouched until PDFsharp has saved and
            // each selected FontFile2 contains its complete composite closure.
            using var staged = new MemoryStream();
            document.Save(staged, false);
            var repaired = PdfBoundFontRestorer.Restore(staged.ToArray(), selectedVerifiedFaces);
            await pdfOutput.WriteAsync(repaired, 0, repaired.Length, cancellationToken).ConfigureAwait(false);
        }
        await pdfOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private async Task<IReadOnlyList<PreparedSignatureAppearance>> PrepareSignatureAppearancesAsync(
        OfdDocumentPackage package, HashSet<string>? selectedIds, bool strictAppearanceBudgets,
        int? maximumSignatureAppearances, OfdModelMaterialization annotationMaterial, FontParseBudget parseBudget, CancellationToken cancellationToken)
    {
        var result = new List<PreparedSignatureAppearance>();
        var nestedCache = new Dictionary<string, (OfdDocumentPackage? Package, DocumentFontContext? Fonts)>();
        var payloadIdentities = new Dictionary<byte[], string>();
        var usedBytes = strictAppearanceBudgets ? package.PreservedEntries.Values.Sum(data => (long)data.Length) : 0;
        var usedEntries = strictAppearanceBudgets ? package.PreservedEntries.Count : 0;
        var usedPages = strictAppearanceBudgets ? package.Pages.Count : 0;
        foreach (var appearance in OfdSignatureAppearanceReader.Read(package, selectedIds, maximumSignatureAppearances, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsZip(appearance.Data))
            {
                // Cache only on the image path; preserve legacy PDF appearance tolerance and accounting.
                var key = string.Empty;
                if (strictAppearanceBudgets && !payloadIdentities.TryGetValue(appearance.Data, out key))
                {
                    key = BinaryIdentity.Hash(appearance.Data);
                    payloadIdentities.Add(appearance.Data, key);
                }
                if (!strictAppearanceBudgets || !nestedCache.TryGetValue(key, out var cached))
                {
                    var remaining = strictAppearanceBudgets
                        ? RemainingAppearanceBudget(usedBytes, usedEntries, usedPages)
                        : null;
                    var nested = await ReadAppearanceOfdPackageAsync(appearance.Data, remaining, annotationMaterial, cancellationToken).ConfigureAwait(false);
                    cached = (nested, nested is null ? null : new DocumentFontContext(nested.Fonts,
                        parseBudget: parseBudget, cancellationToken: cancellationToken,
                        nameOnlyBindings: _nameOnlyBindings, glyphMappingStrategy: _glyphMappingStrategy));
                    if (strictAppearanceBudgets)
                    {
                        nestedCache.Add(key, cached);
                        if (nested is not null)
                        {
                            usedBytes += nested.PreservedEntries.Values.Sum(data => (long)data.Length);
                            usedEntries += nested.PreservedEntries.Count;
                            usedPages += nested.Pages.Count;
                        }
                    }
                }
                if (cached.Package is not null) result.Add(new PreparedSignatureAppearance(appearance, cached.Package, cached.Fonts!));
                continue;
            }
            result.Add(new PreparedSignatureAppearance(appearance, appearance.Data));
        }
        return result;
    }

    private OfdPackageLoadOptions RemainingAppearanceBudget(long usedBytes, int usedEntries, int usedPages)
    {
        var load = _options.PackageLoadOptions;
        if (usedBytes >= load.MaxTotalUncompressedBytes || usedEntries >= load.MaxEntryCount || usedPages >= load.MaxPageCount)
            throw new InvalidDataException("Cumulative OFD signature appearance expansion budget exceeded.");
        return new OfdPackageLoadOptions
        {
            MaxInputBytes = load.MaxInputBytes, MaxEntryUncompressedBytes = load.MaxEntryUncompressedBytes,
            MaxAnnotationObjectCount = load.MaxAnnotationObjectCount, MaxAnnotationXmlBytes = load.MaxAnnotationXmlBytes,
            MaxCompressionRatio = load.MaxCompressionRatio,
            MaxTotalUncompressedBytes = load.MaxTotalUncompressedBytes - usedBytes,
            MaxEntryCount = load.MaxEntryCount - usedEntries,
            MaxPageCount = load.MaxPageCount - usedPages
        };
    }

    private async Task<OfdDocumentPackage?> ReadAppearanceOfdPackageAsync(
        byte[] ofdData, OfdPackageLoadOptions? appearanceBudget,
        OfdModelMaterialization annotationMaterial, CancellationToken cancellationToken)
    {
        try
        {
            var reader = new OfdReader();
            using var input = new MemoryStream(ofdData, writable: false);
            var package = await reader
                .ReadAsync(input, appearanceBudget ?? new OfdPackageLoadOptions(), cancellationToken)
                .ConfigureAwait(false);
            OfdCompositeGraph.Validate(package, annotationMaterial, cancellationToken);
            OfdAnnotationContract.Validate(package, annotationMaterial, cancellationToken);
            var pageModel = package.Pages
                .OrderBy(page => page.Index)
                .FirstOrDefault();
            if (pageModel is null ||
                !PdfOperandGeometry.Box(pageModel.XMillimeters, pageModel.YMillimeters,
                    pageModel.WidthMillimeters, pageModel.HeightMillimeters) ||
                !PdfOperandGeometry.PositiveFormExtent(XUnit.FromMillimeter(pageModel.WidthMillimeters).Point) ||
                !PdfOperandGeometry.PositiveFormExtent(XUnit.FromMillimeter(pageModel.HeightMillimeters).Point))
            {
                if (appearanceBudget is not null) throw new InvalidDataException("Nested OFD appearance has invalid page geometry.");
                return null;
            }

            return package;
        }
        catch (OperationCanceledException) { throw; }
        catch (InvalidDataException) when (appearanceBudget is not null) { throw; }
        catch (OutOfMemoryException) { throw; }
        catch (Exception exception) when (appearanceBudget is not null)
        {
            throw new InvalidDataException("Cannot parse selected nested OFD signature appearance.", exception);
        }
        catch
        {
            return null;
        }
    }

    private static void DrawSignatureAppearances(PdfDocument document, XGraphics graphics, OfdPage page,
        IReadOnlyList<PreparedSignatureAppearance> appearances, long maximumPixels,
        OfdModelMaterialization annotationMaterial, CancellationToken cancellationToken, bool reusePayloads = false)
    {
        var forms = new Dictionary<OfdPage, XForm>();
        using var bitmap = new SinglePayloadResource<XImage>(data =>
        {
            ValidateImage(data, maximumPixels);
            return XImage.FromImageSource(new EncodedBitmapImageSource(data, maximumPixels, cancellationToken));
        });
        try
        {
            foreach (var appearance in appearances.Where(item => string.Equals(item.PageId, page.Id, StringComparison.OrdinalIgnoreCase)))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!PdfOperandGeometry.Box(appearance.XMillimeters - page.XMillimeters,
                    appearance.YMillimeters - page.YMillimeters, appearance.WidthMillimeters, appearance.HeightMillimeters))
                {
                    if (reusePayloads) throw new InvalidDataException("Selected signature stamp placement is not representable in PDF coordinates.");
                    continue;
                }
                var target = new XRect(MillimetersToPoints(appearance.XMillimeters - page.XMillimeters),
                    MillimetersToPoints(appearance.YMillimeters - page.YMillimeters),
                    MillimetersToPoints(appearance.WidthMillimeters), MillimetersToPoints(appearance.HeightMillimeters));
                var state = graphics.Save();
                try
                {
                    if (appearance.Clip is { } clip)
                    {
                        var clipX = appearance.XMillimeters + clip.X - page.XMillimeters;
                        var clipY = appearance.YMillimeters + clip.Y - page.YMillimeters;
                        if (!PdfOperandGeometry.Box(clipX, clipY, clip.Width, clip.Height))
                        {
                            if (reusePayloads) throw new InvalidDataException("Selected signature stamp Clip is not representable in PDF coordinates.");
                            continue;
                        }
                        graphics.IntersectClip(new XRect(MillimetersToPoints(clipX),
                            MillimetersToPoints(clipY), MillimetersToPoints(clip.Width),
                            MillimetersToPoints(clip.Height)));
                    }
                    if (appearance.OfdPage is not null)
                    {
                        var formWidth = XUnit.FromMillimeter(appearance.OfdPage.WidthMillimeters).Point;
                        var formHeight = XUnit.FromMillimeter(appearance.OfdPage.HeightMillimeters).Point;
                        if (!PdfOperandGeometry.PositiveExtent(target.Width / formWidth) ||
                            !PdfOperandGeometry.PositiveExtent(target.Height / formHeight))
                        {
                            if (reusePayloads) throw new InvalidDataException("Selected signature form scale is not representable by the PDF writer.");
                            continue;
                        }
                        if (!reusePayloads)
                        {
                            using var form = CreateAppearanceForm(document, appearance, maximumPixels, annotationMaterial, cancellationToken);
                            graphics.DrawImage(form, target);
                        }
                        else
                        {
                            if (!forms.TryGetValue(appearance.OfdPage, out var form))
                            {
                                form = CreateAppearanceForm(document, appearance, maximumPixels, annotationMaterial, cancellationToken);
                                forms.Add(appearance.OfdPage, form);
                            }
                            graphics.DrawImage(form, target);
                        }
                    }
                    else if (!reusePayloads)
                    {
                        ValidateImage(appearance.Data, maximumPixels);
                        using var image = XImage.FromStream(() => new MemoryStream(appearance.Data, writable: false));
                        graphics.DrawImage(image, target);
                    }
                    else graphics.DrawImage(bitmap.Get(appearance.Data), target);
                }
                catch (OperationCanceledException) { throw; }
                catch (InvalidDataException) { throw; }
                catch (OutOfMemoryException) { throw; }
                catch (NotSupportedException exception) when (
                    OfdRenderPreflight.IsPaintFailure(exception) || appearance.OfdPage is not null && !reusePayloads) { throw; }
                catch (Exception exception) when (reusePayloads)
                {
                    throw new InvalidDataException("Cannot render selected signature appearance.", exception);
                }
                catch { /* Unsupported vendor payloads retain the existing tolerant PDF fallback. */ }
                finally { graphics.Restore(state); }
            }
        }
        finally
        {
            foreach (var form in forms.Values) form.Dispose();
        }
    }

    private static XForm CreateAppearanceForm(PdfDocument document, PreparedSignatureAppearance appearance,
        long maximumPixels, OfdModelMaterialization annotationMaterial, CancellationToken token)
    {
        var form = new XForm(document, XUnit.FromMillimeter(appearance.OfdPage!.WidthMillimeters),
            XUnit.FromMillimeter(appearance.OfdPage.HeightMillimeters));
        try
        {
            using var graphics = XGraphics.FromForm(form);
            DrawPage(graphics, appearance.OfdPage, appearance.Fonts!, maximumPixels, annotationMaterial, token);
            return form;
        }
        catch { form.Dispose(); throw; }
    }

    private static bool IsZip(byte[] data)
    {
        return data.Length >= 4 &&
            data[0] == 0x50 &&
            data[1] == 0x4b &&
            data[2] == 0x03 &&
            data[3] == 0x04;
    }

    private static void DrawPage(XGraphics graphics, OfdPage page, DocumentFontContext fonts,
        long maximumPixels, OfdModelMaterialization annotationMaterial, CancellationToken cancellationToken)
    {
        try
        {
            var outlineFonts = fonts.OutlineFonts;
            var elements = OfdCompositeProjection.Expand(EnumerateRenderableElements(page, annotationMaterial, cancellationToken), annotationMaterial, cancellationToken);
            OfdRenderPreflight.Validate(elements);
            elements = PdfRtlPaintOrder.Arrange(elements, fonts, cancellationToken);
            foreach (var element in elements)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!OfdGraphicXmlContract.IsVisible(element)) continue;
                if (element is OfdRawElement glyphXml && EmbeddedFontCoverage.HasExplicitGlyphReferences(glyphXml.Xml))
                    throw new NotSupportedException("PDF export does not model raw CGTransform glyph substitutions.");
                if (element is OfdRawElement { LocalName: "UnsupportedAnnotationAppearance" })
                    throw new NotSupportedException($"Annotation appearance on page {page.Index + 1} contains unsupported drawing; export would lose content.");
                var elementState = graphics.Save();
                try
                {
                    if (element is not OfdImageElement && (element.Clips is not null || !string.IsNullOrWhiteSpace(element.ClippingXml)))
                    {
                        var dx = MillimetersToPoints(element.XMillimeters - page.XMillimeters);
                        var dy = MillimetersToPoints(element.YMillimeters - page.YMillimeters);
                        graphics.TranslateTransform(dx, dy);
                        foreach (var region in OfdClipGeometry.Read(element, annotationMaterial, cancellationToken)) graphics.IntersectClip(OfdPathRenderer.BuildClip(region));
                        graphics.TranslateTransform(-dx, -dy);
                    }
                    if (element is OfdTextElement text)
                    {
                        text = OfdTextRenderProjection.Create(text, annotationMaterial, cancellationToken);
                        var paint = OfdTextPaint.Read(text);
                        var fontSize = Math.Max(0.1, MillimetersToPoints(text.FontSizeMillimeters));
                        var familyName = fonts.Resolve(text, out var resource);
                        if ((text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text)).All(string.IsNullOrEmpty)) continue;
                        var coverage = fonts.Coverage(resource, familyName);
                        Func<string, string> mapPdfText = value => fonts.PdfText(familyName, value);
                        // CT_Text Weight/Italic is the per-object style viewers apply;
                        // the resource flags describe the bound font file.
                        var bold = resource?.Bold == true || text.Weight >= 600;
                        var italic = resource?.Italic == true || text.Italic;
                        var style = (bold ? XFontStyle.Bold : XFontStyle.Regular) |
                            (italic ? XFontStyle.Italic : XFontStyle.Regular);
                        XFont font;
                        FontResolverInfo face;
                        try
                        {
                            font = new XFont(familyName, fontSize, style);
                            face = GlobalFontSettings.FontResolver.ResolveTypeface(familyName, bold, italic);
                        }
                        catch (Exception exception) when (exception is not OutOfMemoryException && exception is not OperationCanceledException)
                        {
                            throw new InvalidDataException($"Font '{resource?.FontName ?? familyName}' could not be initialized without content loss.", exception);
                        }
                        var simulateBold = face.MustSimulateBold;
                        var simulateItalic = face.MustSimulateItalic;
                        SixLabors.Fonts.Font? outlineFont = null;
                        if (simulateBold || paint.Stroke && paint.StrokeColor.Alpha > 0)
                        {
                            if (!outlineFonts.TryGetValue(face.FaceName, out var outlineFace))
                            {
                                var fontBytes = fonts.OutlineFontBytes(familyName, GlobalFontSettings.FontResolver.GetFont(face.FaceName));
                                using var fontStream = new MemoryStream(fontBytes);
                                var family = new SixLabors.Fonts.FontCollection().Add(fontStream);
                                var physical = new OpenTypeFace(fontBytes).Style;
                                var physicalStyle = (physical.Bold ? SixLabors.Fonts.FontStyle.Bold : SixLabors.Fonts.FontStyle.Regular) |
                                    (physical.Italic ? SixLabors.Fonts.FontStyle.Italic : SixLabors.Fonts.FontStyle.Regular);
                                outlineFace = (family, physicalStyle);
                                outlineFonts.Add(face.FaceName, outlineFace);
                            }
                            outlineFont = outlineFace.Family.CreateFont((float)fontSize, outlineFace.Style);
                        }
                        var brush = new XSolidBrush(XColor.FromArgb(
                            paint.Fill ? paint.FillColor.Alpha : 0,
                            paint.FillColor.Red,
                            paint.FillColor.Green,
                            paint.FillColor.Blue));
                        if (text.Runs.Count > 0)
                        {
                            DrawTextRunsMapped(
                                graphics,
                                text,
                                font,
                                brush,
                                page.XMillimeters,
                                page.YMillimeters,
                                outlineFont,
                                simulateBold, simulateItalic, coverage, mapPdfText, cancellationToken, fonts, resource, familyName,
                                explicitGlyphs: EmbeddedFontCoverage.HasExplicitGlyphReferences(text));
                        }
                        else
                        {
                            DrawTextWithMatrix(graphics, text, page.XMillimeters, page.YMillimeters, factor =>
                            {
                                var anchor = OfdTextEmphasis.Anchor(0, 0, factor);
                                DrawStyledString(graphics, text.Text, font, brush, new XPoint(MillimetersToPoints(anchor.X / paint.HScale), MillimetersToPoints(anchor.Y)), outlineFont, simulateBold, simulateItalic, XStringFormats.TopLeft, coverage, paint, mapPdfText, fonts, resource, familyName);
                            });
                        }
                    }

                    if (element is OfdImageElement image && image.Data.Length > 0)
                    {
                        DrawImage(graphics, page, image, maximumPixels, annotationMaterial, cancellationToken);
                    }

                    if (element is OfdPathElement path)
                    {
                        if ((path.Stroke || path.Fill) && !OfdPathRenderer.TryDraw(graphics, path, page.XMillimeters, page.YMillimeters))
                            throw new NotSupportedException("OFD path literal could not be rendered.");
                    }
                }
                finally { graphics.Restore(elementState); }
            }

        }
        catch (OperationCanceledException) { throw; }
        catch (InvalidDataException) { throw; }
        catch (NotSupportedException) { throw; }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            throw new InvalidDataException($"OFD page {page.Index + 1} could not be rendered without content loss.", exception);
        }
    }

    private static IEnumerable<OfdTextElement> ExplicitGlyphTexts(IEnumerable<OfdPage> pages,
        OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        foreach (var page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var elements = OfdCompositeProjection.Expand(
                EnumerateRenderableElements(page, material, cancellationToken), material, cancellationToken);
            foreach (var element in elements)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (OfdGraphicXmlContract.IsVisible(element) &&
                    element is OfdTextElement { SourceXml: { } source } text &&
                    source.Contains("CGTransform", StringComparison.Ordinal))
                    yield return OfdTextRenderProjection.Create(text, material, cancellationToken);
            }
        }
    }

    internal static void DrawImage(XGraphics graphics, OfdPage page, OfdImageElement image, long maximumPixels, OfdModelMaterialization material, CancellationToken cancellationToken,
        Func<byte[], Image<Rgba32>>? decode = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var projected = OfdImageRenderProjection.Prepare(image, maximumPixels, cancellationToken);
        ValidateImage(projected.Data, maximumPixels);
        var state = graphics.Save();
        try
        {
            graphics.TranslateTransform(MillimetersToPoints(image.XMillimeters - page.XMillimeters),
                MillimetersToPoints(image.YMillimeters - page.YMillimeters));
            graphics.IntersectClip(new XRect(0, 0, MillimetersToPoints(image.WidthMillimeters), MillimetersToPoints(image.HeightMillimeters)));
            foreach (var clip in OfdClipGeometry.Read(image, material, cancellationToken))
                graphics.IntersectClip(OfdPathRenderer.BuildClip(clip));
            var matrix = image.Transform ?? new double[] { image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0 };
            if (matrix.Length != 6 || matrix.Any(value => double.IsNaN(value) || double.IsInfinity(value)))
                throw new InvalidDataException("Invalid OFD image transform.");
            var crop = OfdImageRenderingHints.ReadPdfSourceCrop(image);
            var data = projected.Data;
            if (crop is { } sourceCrop)
            {
                var prepared = PdfCropImage.Create(data, sourceCrop, image.Alpha, maximumPixels, cancellationToken, decode);
                data = prepared.Data;
                var sourceMatrix = matrix;
                matrix = new[]
                {
                    sourceMatrix[0] * prepared.Width, sourceMatrix[1] * prepared.Width,
                    sourceMatrix[2] * prepared.Height, sourceMatrix[3] * prepared.Height,
                    sourceMatrix[0] * prepared.U + sourceMatrix[2] * prepared.V + sourceMatrix[4],
                    sourceMatrix[1] * prepared.U + sourceMatrix[3] * prepared.V + sourceMatrix[5]
                };
            }
            graphics.MultiplyTransform(new XMatrix(MillimetersToPoints(matrix[0]), MillimetersToPoints(matrix[1]),
                MillimetersToPoints(matrix[2]), MillimetersToPoints(matrix[3]), MillimetersToPoints(matrix[4]), MillimetersToPoints(matrix[5])));
            if (crop is null && image.Alpha < 255)
            {
                using var pixels = decode is null ? Image.Load<Rgba32>(data) : decode(data);
                cancellationToken.ThrowIfCancellationRequested();
                pixels.Mutate(context => context.Opacity(Math.Max(0, image.Alpha) / 255f));
                using var png = new MemoryStream();
                // RGB PNG metadata can otherwise discard the alpha added by Opacity.
                pixels.Save(png, new PngEncoder { ColorType = PngColorType.RgbWithAlpha });
                data = png.ToArray();
            }
            var interpolate = crop is null && (OfdImageRenderingHints.ReadPdfInterpolate(image) ?? true);
            using var xImage = XImage.FromImageSource(new EncodedBitmapImageSource(data, maximumPixels, cancellationToken, decode, interpolate));
            xImage.Interpolate = interpolate;
            graphics.DrawImage(xImage, 0, 0, 1, 1);
        }
        finally
        {
            graphics.Restore(state);
        }
    }

    private static void ValidateImage(byte[] data, long maximumPixels)
    {
        ImageFormatAdmission.RejectBigTiff(data);
        var info = Image.Identify(data);
        if (info is null || info.Width <= 0 || info.Height <= 0 || (long)info.Width * info.Height > maximumPixels)
            throw new InvalidDataException("OFD image exceeds the configured decoded pixel limit or has invalid dimensions.");
    }

    private static double MillimetersToPoints(double millimeters)
    {
        return millimeters * 72d / 25.4d;
    }

    private static void DrawStyledString(XGraphics graphics, string text, XFont font, XBrush brush,
        XPoint point, SixLabors.Fonts.Font? outlineFont, bool simulateBold, bool italic, XStringFormat? format = null, OpenTypeCmap? coverage = null,
        OfdTextPaint? paint = null, Func<string, string>? mapPdfText = null,
        DocumentFontContext? fontContext = null, OfdFontResource? fontResource = null, string? fontFamily = null,
        bool preserveExplicitSourceSpan = false, double? sourceBoldStrokeWidth = null)
    {
        // PDFsharp does not compose Latin mark sequences (notably dotted i).
        // Match the supported composite used for measurement when drawing one
        // grapheme. OFD source text stays unchanged; PDF text is canonically equivalent.
        // An explicit complete source span must reach its private cmap alias
        // unchanged, including a Hangul filler absent from the ordinary cmap.
        if (!preserveExplicitSourceSpan && text.Length > 1 && text[0] <= 0xFF && char.IsLetter(text[0]))
        {
            var elements = StringInfo.GetTextElementEnumerator(text);
            if (elements.MoveNext() && !elements.MoveNext())
                text = FlowTextMetrics.NormalizeLatinForMeasurement(text);
        }
        if (preserveExplicitSourceSpan) PdfTextControlPolicy.Validate(text);
        else text = PdfTextControlPolicy.VisibleText(text, coverage);
        if (text.Length == 0) return;
        var pdfText = mapPdfText is null ? text : mapPdfText(text);
        var explicitStroke = paint?.Stroke == true && paint.StrokeColor.Alpha > 0;
        var strokeWidthPoints = explicitStroke ? MillimetersToPoints(paint!.LineWidthMillimeters) : 0;
        var opaqueSameColorFill = paint is { Stroke: true, Fill: true } &&
            paint.FillColor.Alpha == 255 && paint.StrokeColor.Alpha == 255 &&
            brush is XSolidBrush opaque && opaque.Color.A == XColors.Black.A &&
            opaque.Color.R == paint.StrokeColor.Red && opaque.Color.G == paint.StrokeColor.Green &&
            opaque.Color.B == paint.StrokeColor.Blue;
        var geometricStroke = opaqueSameColorFill && strokeWidthPoints > 0 && outlineFont is not null;
        // PDFsharp Core 1.3.67 drops resolver style simulations when creating
        // XGlyphTypeface. Apply the missing fallback appearance at draw time.
        var state = graphics.Save();
        try
        {
            // PDFsharp Core can omit the zero-alpha graphics state for text
            // after an image. Keep semantic text extractable, but guarantee
            // that a fully transparent fill cannot paint any pixels.
            // A positioned whitespace glyph must remain blank even when the bound
            // font lacks it and returns a visible .notdef outline (e.g. narrow NBSP).
            // Still emit semantic text and let OFD DeltaX/DeltaY carry its advance.
            var invisibleFill = text.All(char.IsWhiteSpace) || brush is XSolidBrush transparent && transparent.Color.A == 0;
            graphics.TranslateTransform(point.X, point.Y);
            if (italic)
            {
                graphics.MultiplyTransform(new XMatrix(1, 0, -0.21256, 1, 0, 0));
            }
            if (invisibleFill || geometricStroke)
            {
                // Keep one semantic font string. The opaque outline route paints
                // the original glyph once; mixing hinted text and raw contours
                // can close small counters even when their design bounds agree.
                var semanticState = graphics.Save();
                try
                {
                    graphics.IntersectClip(new XRect(0, 0, 0, 0));
                    graphics.DrawString(pdfText, font, brush, new XRect(0, 0, 0, 0), format ?? XStringFormats.Default);
                }
                finally { graphics.Restore(semanticState); }
            }
            else graphics.DrawString(pdfText, font, brush, new XRect(0, 0, 0, 0), format ?? XStringFormats.Default);
            fontContext?.MarkFontFaceDrawn(fontResource, fontFamily);
            // Native DOCX writes its simulated bold stroke into the OFD itself.
            // Recognize the same-color stroke emitted for the whole TextObject,
            // even when positioned runs are drawn one glyph at a time. Older
            // documents used a 2.5%-em stroke for Han text; retain that contract.
            // A different color or width remains an independent OFD outline.
            var ownBoldStroke = simulateBold && explicitStroke && paint!.Fill &&
                paint.FillColor.Red == paint.StrokeColor.Red &&
                paint.FillColor.Green == paint.StrokeColor.Green &&
                paint.FillColor.Blue == paint.StrokeColor.Blue &&
                paint.FillColor.Alpha == paint.StrokeColor.Alpha &&
                (Math.Abs(strokeWidthPoints - (sourceBoldStrokeWidth ?? FontEmphasisMetrics.BoldStrokeWidth(font.Size, text))) <= 0.01 ||
                 Math.Abs(strokeWidthPoints - FontEmphasisMetrics.BoldStrokeWidth(font.Size)) <= 0.01);
            var simulatedStroke = simulateBold && paint?.Fill == true && !ownBoldStroke;
            if (!text.All(char.IsWhiteSpace) && outlineFont is not null && brush is XSolidBrush solid && (explicitStroke || simulatedStroke))
            {
                var topBearing = 0;
                for (var index = 0; index < text.Length; index++)
                {
                    var codePoint = char.ConvertToUtf32(text, index);
                    if (char.IsHighSurrogate(text[index])) index++;
                    if (outlineFont.TryGetGlyphs(new SixLabors.Fonts.Unicode.CodePoint(codePoint),
                            SixLabors.Fonts.ColorFontSupport.None, out var glyphs))
                    {
                        foreach (var glyph in glyphs)
                            topBearing = Math.Min(topBearing, glyph.GlyphMetrics.TopSideBearing);
                    }
                }
                // Fonts 1.0.1 centers the em box within the horizontal line metrics,
                // then shifts the ascender for negative top bearings. Undo both
                // offsets; otherwise the outline and PDF text have different baselines.
                var metrics = outlineFont.FontMetrics;
                var layoutAscender = metrics.HorizontalMetrics.Ascender - topBearing -
                    (metrics.HorizontalMetrics.LineHeight - metrics.UnitsPerEm) * 0.5f;
                var baselineOffset = -layoutAscender * outlineFont.Size / metrics.UnitsPerEm;
                if (format == XStringFormats.TopLeft)
                {
                    // Use the same top-to-baseline distance as PDFsharp DrawString.
                    baselineOffset += (float)(font.GetHeight() * font.CellAscent / font.CellSpace);
                }
                void RenderOutline(XColor color, double width, bool fill, bool strokeAsFill = false) =>
                    SixLabors.Fonts.TextRenderer.RenderTextTo(
                        new PdfGlyphOutlineRenderer(graphics, color, width, fill, strokeAsFill), text,
                        new SixLabors.Fonts.TextOptions(outlineFont)
                        {
                            Dpi = 72,
                            Origin = new Vector2(0, baselineOffset),
                            KerningMode = SixLabors.Fonts.KerningMode.None
                        });
                if (simulatedStroke && !geometricStroke)
                    RenderOutline(solid.Color, FontEmphasisMetrics.BoldStrokeWidth(font.Size, text), true);
                if (explicitStroke)
                {
                    // Outline-only text has no fill for the usual bold overlay.
                    // Thicken its existing stroke while keeping the counters open.
                    var strokeWidth = MillimetersToPoints(paint!.LineWidthMillimeters) +
                        (simulateBold && !paint.Fill ? FontEmphasisMetrics.BoldStrokeWidth(font.Size, text) : 0);
                    // Two opaque same-color centered strokes cover the wider
                    // region. Emit their union once, including explicit widths
                    // other than the native writer's historical 1% and 2.5%.
                    // Mixing an additional subpixel XPen stroke with geometry
                    // would reintroduce renderer-dependent boundary expansion.
                    if (geometricStroke && simulatedStroke)
                        strokeWidth = Math.Max(strokeWidth, FontEmphasisMetrics.BoldStrokeWidth(font.Size, text));
                    // A font fill plus a separately rasterized contour boundary
                    // can leave light seams inside thin glyph strokes. Repaint
                    // the original contour only when both paints are opaque and
                    // identical; repeating that fill cannot change compositing.
                    // Different colors, translucent paints and outline-only text
                    // retain their independent boundary-stroke semantics.
                    RenderOutline(XColor.FromArgb(paint.StrokeColor.Alpha, paint.StrokeColor.Red,
                        paint.StrokeColor.Green, paint.StrokeColor.Blue), strokeWidth, opaqueSameColorFill, geometricStroke);
                }
            }
        }
        finally
        {
            graphics.Restore(state);
        }
    }

    private static void DrawTextRuns(
        XGraphics graphics,
        OfdTextElement text,
        XFont font,
        XBrush brush,
        double pageOriginX,
        double pageOriginY,
        SixLabors.Fonts.Font? outlineFont,
        bool simulateBold, bool simulateItalic, OpenTypeCmap? coverage, CancellationToken cancellationToken) =>
        DrawTextRunsMapped(graphics, text, font, brush, pageOriginX, pageOriginY, outlineFont,
            simulateBold, simulateItalic, coverage, static value => value, cancellationToken);

    private static void DrawTextRunsMapped(
        XGraphics graphics,
        OfdTextElement text,
        XFont font,
        XBrush brush,
        double pageOriginX,
        double pageOriginY,
        SixLabors.Fonts.Font? outlineFont,
        bool simulateBold, bool simulateItalic, OpenTypeCmap? coverage, Func<string, string> mapPdfText, CancellationToken cancellationToken,
        DocumentFontContext? fontContext = null, OfdFontResource? fontResource = null, string? fontFamily = null,
        bool explicitGlyphs = false)
    {
        var paint = OfdTextPaint.Read(text);
        var explicitRun = explicitGlyphs
            ? (fontContext ?? throw new InvalidOperationException("Explicit glyph rendering requires a document font context."))
                .ReadExplicitRun(text, fontResource)
            : null;
        var groupedFont = explicitRun?.Groups is not null
            ? OfdExplicitGlyphFont.Read(fontResource, text, explicitRun, embedFont: false) : null;
        // Positioned text can call DrawStyledString once per glyph. The OFD
        // writer chose one stroke width for the whole TextObject, so scan its
        // source text once instead of once for every glyph.
        double? sourceBoldStrokeWidth = simulateBold && paint.Stroke && paint.Fill && paint.StrokeColor.Alpha > 0
            ? FontEmphasisMetrics.BoldStrokeWidth(font.Size, text.Text) : null;
        DrawTextWithMatrix(graphics, text, pageOriginX, pageOriginY, factor =>
        {
            for (var runIndex = 0; runIndex < text.Runs.Count; runIndex++)
            {
                var run = text.Runs[runIndex];
                var plannedRun = explicitRun?.Groups is { } groups ? groups[runIndex].Projection : explicitRun;
                if (plannedRun?.PathGlyphSourceSpan is { } sourceSpan)
                {
                    var anchor = OfdTextEmphasis.Anchor(run.XMillimeters, run.YMillimeters, factor);
                    DrawPathGlyphOutlines(graphics, text, plannedRun, fontResource, factor, paint,
                        cancellationToken, groupedFont?.Slice(explicitRun!.Groups![runIndex].GlyphOffset,
                            plannedRun.GlyphIds.Count));
                    if (plannedRun.VisualHebrewSourceSpan is not null)
                    {
                        // Keep character semantics at each independently verified original-GID
                        // origin. A whole-word CID cannot expose per-character geometry.
                        var origins = plannedRun.PathGlyphOrigins(cancellationToken);
                        for (var index = 0; index < origins.Count; index++)
                        {
                            var point = OfdTextEmphasis.Anchor(origins[index].X, origins[index].Y, factor);
                            DrawStyledString(graphics, sourceSpan[sourceSpan.Length - 1 - index].ToString(), font,
                                new XSolidBrush(XColor.FromArgb(0, 0, 0, 0)),
                                new XPoint(MillimetersToPoints(point.X / paint.HScale), MillimetersToPoints(point.Y)),
                                null, false, false, coverage: coverage, mapPdfText: mapPdfText,
                                fontContext: fontContext, fontResource: fontResource, fontFamily: fontFamily);
                        }
                        continue;
                    }
                    DrawStyledString(graphics, sourceSpan, font,
                        new XSolidBrush(XColor.FromArgb(0, 0, 0, 0)),
                        new XPoint(MillimetersToPoints(anchor.X / paint.HScale), MillimetersToPoints(anchor.Y)),
                        // This glyph supplies semantics only. Passing the
                        // source paint/outline font would redraw its ordinary
                        // Unicode outline when Stroke is enabled.
                        null, false, false, coverage: coverage,
                        mapPdfText: mapPdfText, fontContext: fontContext, fontResource: fontResource,
                        fontFamily: fontFamily, preserveExplicitSourceSpan: true);
                    continue;
                }
                if (explicitRun?.SingleGlyphSourceSpan is not null)
                {
                    var anchor = OfdTextEmphasis.Anchor(run.XMillimeters, run.YMillimeters, factor);
                    DrawStyledString(graphics, run.Text, font, brush,
                        new XPoint(MillimetersToPoints(anchor.X / paint.HScale), MillimetersToPoints(anchor.Y)),
                        outlineFont, simulateBold, simulateItalic, coverage: coverage, paint: paint,
                        mapPdfText: mapPdfText, fontContext: fontContext, fontResource: fontResource,
                        fontFamily: fontFamily, preserveExplicitSourceSpan: true,
                        sourceBoldStrokeWidth: sourceBoldStrokeWidth);
                    continue;
                }
                // CGTransform maps each validated BMP character to one original
                // GID. Grapheme segmentation can combine adjacent characters
                // (for example Hangul Jamo) and lose a glyph-slot advance.
                var deltas = explicitGlyphs
                    ? (X: OfdTextGeometry.ExpandDeltas(run.DeltaX, run.Text.Length - 1, cancellationToken),
                        Y: OfdTextGeometry.ExpandDeltas(run.DeltaY, run.Text.Length - 1, cancellationToken))
                    : OfdTextGeometry.PositioningDeltas(run, cancellationToken);
                var deltaX = deltas.X; var deltaY = deltas.Y;
                if (deltaX.Count == 0 && deltaY.Count == 0)
                {
                    var anchor = OfdTextEmphasis.Anchor(run.XMillimeters, run.YMillimeters, factor);
                    if (explicitGlyphs && run.Text.Length > 1)
                    {
                        // The standard says absent DeltaX/DeltaY leaves every
                        // subsequent glyph at the same origin. A whole PDF
                        // DrawString would add font advances.
                        foreach (var code in run.Text)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            DrawStyledString(graphics, code.ToString(), font, brush,
                                new XPoint(MillimetersToPoints(anchor.X / paint.HScale), MillimetersToPoints(anchor.Y)),
                                outlineFont, simulateBold, simulateItalic, coverage: coverage, paint: paint,
                                mapPdfText: mapPdfText, fontContext: fontContext, fontResource: fontResource,
                                fontFamily: fontFamily, sourceBoldStrokeWidth: sourceBoldStrokeWidth);
                        }
                        continue;
                    }
                    DrawStyledString(graphics, run.Text, font, brush,
                        new XPoint(MillimetersToPoints(anchor.X / paint.HScale), MillimetersToPoints(anchor.Y)), outlineFont, simulateBold, simulateItalic, coverage: coverage, paint: paint, mapPdfText: mapPdfText,
                        fontContext: fontContext, fontResource: fontResource, fontFamily: fontFamily,
                        sourceBoldStrokeWidth: sourceBoldStrokeWidth);
                    continue;
                }
                var x = run.XMillimeters; var y = run.YMillimeters;
                var i = 0;
                foreach (var glyph in explicitGlyphs
                    ? run.Text.Select(character => character.ToString())
                    : OfdTextGeometry.EnumerateGlyphs(run.Text, cancellationToken))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var anchor = OfdTextEmphasis.Anchor(x, y, factor);
                    DrawStyledString(graphics, glyph, font, brush,
                        new XPoint(MillimetersToPoints(anchor.X / paint.HScale), MillimetersToPoints(anchor.Y)), outlineFont, simulateBold, simulateItalic, coverage: coverage, paint: paint, mapPdfText: mapPdfText,
                        fontContext: fontContext, fontResource: fontResource, fontFamily: fontFamily,
                        sourceBoldStrokeWidth: sourceBoldStrokeWidth);
                    if (i < deltaX.Count) x += deltaX[i];
                    if (i < deltaY.Count) y += deltaY[i];
                    i++;
                }
            }
        });
    }

    private static void DrawPathGlyphOutlines(XGraphics graphics, OfdTextElement text,
        OfdExplicitGlyphRun run, OfdFontResource? resource, double[]? factor,
        OfdTextPaint paint, CancellationToken token, OfdExplicitGlyphFont? plannedFont = null)
    {
        var font = plannedFont ?? OfdExplicitGlyphFont.Read(resource, text, run, embedFont: false);
        var origins = run.PathGlyphOrigins(token);
        for (var index = 0; index < font.Outlines.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            var outline = font.Outlines[index];
            var origin = origins[index];
            var anchor = OfdTextEmphasis.Anchor(origin.X, origin.Y, factor);
            // Every source GID uses its explicit OFD draw point. Absent
            // deltas in the previously accepted 1:2 profile share one origin.
            var scale = text.FontSizeMillimeters / outline.UnitsPerEm;
            var path = new XGraphicsPath { FillMode = XFillMode.Winding };
            double x = 0, y = 0;
            foreach (var command in outline.ToPathCommands())
            {
                var nextX = MillimetersToPoints(anchor.X / paint.HScale + command.X * scale);
                var nextY = MillimetersToPoints(anchor.Y - command.Y * scale);
                switch (command.Verb)
                {
                    case TrueTypePathVerb.Move:
                        path.StartFigure(); x = nextX; y = nextY;
                        break;
                    case TrueTypePathVerb.Line:
                        path.AddLine(x, y, nextX, nextY); x = nextX; y = nextY;
                        break;
                    case TrueTypePathVerb.Quadratic:
                        var controlX = MillimetersToPoints(anchor.X / paint.HScale + command.ControlX * scale);
                        var controlY = MillimetersToPoints(anchor.Y - command.ControlY * scale);
                        var cubic1X = x + (controlX - x) * 2d / 3d;
                        var cubic1Y = y + (controlY - y) * 2d / 3d;
                        var cubic2X = nextX + (controlX - nextX) * 2d / 3d;
                        var cubic2Y = nextY + (controlY - nextY) * 2d / 3d;
                        path.AddBezier(x, y, cubic1X, cubic1Y, cubic2X, cubic2Y, nextX, nextY);
                        x = nextX; y = nextY;
                        break;
                    case TrueTypePathVerb.Close:
                        path.CloseFigure();
                        break;
                }
            }
            XBrush? fill = paint.Fill && paint.FillColor.Alpha > 0 ?
                new XSolidBrush(XColor.FromArgb(paint.FillColor.Alpha,
                    paint.FillColor.Red, paint.FillColor.Green, paint.FillColor.Blue)) : null;
            XPen? stroke = paint.Stroke && paint.StrokeColor.Alpha > 0 ?
                new XPen(XColor.FromArgb(paint.StrokeColor.Alpha,
                    paint.StrokeColor.Red, paint.StrokeColor.Green, paint.StrokeColor.Blue),
                    MillimetersToPoints(paint.LineWidthMillimeters)) : null;
            if (fill is not null && stroke is not null) graphics.DrawPath(stroke, fill, path);
            else if (fill is not null) graphics.DrawPath(fill, path);
            else if (stroke is not null) graphics.DrawPath(stroke, path);
        }
    }

    private static void DrawTextWithMatrix(XGraphics graphics, OfdTextElement text,
        double pageOriginX, double pageOriginY, Action<double[]?> draw)
    {
        var state = graphics.Save();
        try
        {
            graphics.TranslateTransform(MillimetersToPoints(text.XMillimeters - pageOriginX), MillimetersToPoints(text.YMillimeters - pageOriginY));
            var transform = OfdTextEmphasis.DrawingTransform(text);
            if (transform.Matrix is { Length: 6 } matrix)
            {
                if (matrix.Any(value => double.IsNaN(value) || double.IsInfinity(value))) throw new InvalidDataException("Invalid OFD text transform.");
                graphics.MultiplyTransform(new XMatrix(matrix[0], matrix[1], matrix[2], matrix[3], MillimetersToPoints(matrix[4]), MillimetersToPoints(matrix[5])));
            }
            var paint = OfdTextPaint.Read(text);
            if (paint.HScale != 1) graphics.ScaleTransform(paint.HScale, 1);
            draw(transform.Factor);
        }
        finally { graphics.Restore(state); }
    }

    private static IEnumerable<OfdElement> EnumerateRenderableElements(OfdPage page, OfdModelMaterialization annotationMaterial, CancellationToken token)
    {
        foreach (var template in page.Templates.Where(x =>
            string.Equals(x.ZOrder, "Background", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var element in template.Elements)
            {
                yield return element;
            }
        }

        foreach (var element in page.Elements)
        {
            yield return element;
        }

        foreach (var template in page.Templates.Where(x =>
            !string.Equals(x.ZOrder, "Background", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var element in template.Elements)
            {
                yield return element;
            }
        }
        foreach (var occurrence in OfdAnnotationProjection.Resolve(page, annotationMaterial, token)) yield return occurrence.RenderElement;
    }

    private sealed class PreparedSignatureAppearance
    {
        public PreparedSignatureAppearance(
            OfdSignatureAppearance source,
            OfdDocumentPackage package, DocumentFontContext fonts)
        {
            PageId = source.PageId;
            XMillimeters = source.XMillimeters;
            YMillimeters = source.YMillimeters;
            WidthMillimeters = source.WidthMillimeters;
            HeightMillimeters = source.HeightMillimeters;
            Clip = source.Clip;
            Data = Array.Empty<byte>();
            OfdPage = package.Pages.OrderBy(page => page.Index).First();
            Fonts = fonts;
        }

        public PreparedSignatureAppearance(
            OfdSignatureAppearance source,
            byte[] data)
        {
            PageId = source.PageId;
            XMillimeters = source.XMillimeters;
            YMillimeters = source.YMillimeters;
            WidthMillimeters = source.WidthMillimeters;
            HeightMillimeters = source.HeightMillimeters;
            Clip = source.Clip;
            Data = data;
        }

        public string PageId { get; }

        public double XMillimeters { get; }

        public double YMillimeters { get; }

        public double WidthMillimeters { get; }

        public double HeightMillimeters { get; }

        public OfdSignatureClip? Clip { get; }

        public byte[] Data { get; }

        public OfdPage? OfdPage { get; }
        internal DocumentFontContext? Fonts { get; }
    }
}
