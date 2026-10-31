using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Converter.Svg.Internal;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Svg.Converters;

/// <summary>
/// Converts one OFD page to a self-contained SVG document.
/// </summary>
public sealed class OfdToSvgConverter
{
    private static readonly Regex PathTokenPattern = new(
        @"[A-Za-z]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly OfdPackageLoadOptions _loadOptions;
    private readonly OfdGlyphMappingStrategy _glyphMappingStrategy;
    private readonly long _maximumFontParseWork;
    private readonly OfdNameOnlyFontBindings? _nameOnlyBindings;

    public OfdToSvgConverter() : this(new OfdPackageLoadOptions()) { }

    public OfdToSvgConverter(OfdPackageLoadOptions loadOptions) : this(loadOptions, FontParseBudget.DefaultMaximumWork) { }

    /// <summary>Creates an SVG exporter with a finite cumulative cmap parse budget.</summary>
    /// <param name="loadOptions">OFD loading limits.</param>
    /// <param name="maximumCmapParseWork">Positive cumulative limit for cmap bytes, parsed records and entries inspected by variation coverage queries.</param>
    public OfdToSvgConverter(OfdPackageLoadOptions loadOptions, long maximumCmapParseWork)
        : this(loadOptions, maximumCmapParseWork, null) { }

    /// <summary>Creates an SVG exporter with strict conversion-local bindings for name-only fonts.</summary>
    public OfdToSvgConverter(OfdPackageLoadOptions loadOptions, OfdNameOnlyFontBindings nameOnlyBindings)
        : this(loadOptions, FontParseBudget.DefaultMaximumWork,
            nameOnlyBindings ?? throw new ArgumentNullException(nameof(nameOnlyBindings))) { }

    /// <summary>Creates an SVG exporter with explicit name-only face bindings and a finite parse budget.</summary>
    public OfdToSvgConverter(OfdPackageLoadOptions loadOptions, long maximumCmapParseWork,
        OfdNameOnlyFontBindings? nameOnlyBindings)
        : this(loadOptions, maximumCmapParseWork, nameOnlyBindings, OfdGlyphMappingStrategy.LegacyConcatenatedGlobal) { }

    /// <summary>Creates an SVG exporter with an explicit CGTransform interpretation.</summary>
    public OfdToSvgConverter(OfdGlyphMappingStrategy glyphMappingStrategy)
        : this(new OfdPackageLoadOptions(), FontParseBudget.DefaultMaximumWork, null, glyphMappingStrategy) { }

    /// <summary>Creates an SVG exporter with explicit resource limits, face bindings and mapping strategy.</summary>
    public OfdToSvgConverter(OfdPackageLoadOptions loadOptions, long maximumCmapParseWork,
        OfdNameOnlyFontBindings? nameOnlyBindings, OfdGlyphMappingStrategy glyphMappingStrategy)
    {
        OfdGlyphMappingStrategyGuard.Validate(glyphMappingStrategy);
        _glyphMappingStrategy = glyphMappingStrategy;
        if (maximumCmapParseWork <= 0) throw new ArgumentOutOfRangeException(nameof(maximumCmapParseWork));
        _maximumFontParseWork = maximumCmapParseWork;
        _loadOptions = loadOptions ?? throw new ArgumentNullException(nameof(loadOptions));
        _nameOnlyBindings = nameOnlyBindings;
    }

    public async Task ConvertAsync(
        Stream ofdInput,
        Stream svgOutput,
        int pageIndex = 0,
        CancellationToken cancellationToken = default)
    {
        if (ofdInput is null)
        {
            throw new ArgumentNullException(nameof(ofdInput));
        }

        if (svgOutput is null)
        {
            throw new ArgumentNullException(nameof(svgOutput));
        }

        var package = await new OfdReader()
            .ReadAsync(ofdInput, _loadOptions, cancellationToken)
            .ConfigureAwait(false);
        var pages = package.Pages.OrderBy(page => page.Index).ToList();
        if (pageIndex < 0 || pageIndex >= pages.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(pageIndex));
        }

        var material = new OfdModelMaterialization(new OfdModelLimits());
        OfdCompositeGraph.Validate(package, material, cancellationToken);
        OfdAnnotationContract.Validate(package, material, cancellationToken);
        var svg = BuildSvg(package, pages[pageIndex], material, new FontParseBudget(_maximumFontParseWork, cancellationToken),
            cancellationToken, nameOnlyBindings: _nameOnlyBindings, glyphMappingStrategy: _glyphMappingStrategy);
        var bytes = Encoding.UTF8.GetBytes(
            new XDocument(new XDeclaration("1.0", "utf-8", null), svg)
                .ToString(SaveOptions.DisableFormatting));
        await svgOutput.WriteAsync(bytes, 0, bytes.Length, cancellationToken)
            .ConfigureAwait(false);
        await svgOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    internal static XElement BuildSvg(OfdDocumentPackage package, OfdPage page, OfdModelMaterialization material, FontParseBudget parseBudget,
        CancellationToken cancellationToken, HtmlRenderContext? preview = null, OfdNameOnlyFontBindings? nameOnlyBindings = null,
        OfdGlyphMappingStrategy glyphMappingStrategy = OfdGlyphMappingStrategy.LegacyConcatenatedGlobal)
    {
        preview?.StartPage(page);
        var svgNs = XNamespace.Get("http://www.w3.org/2000/svg");
        var root = new XElement(
            svgNs + "svg",
            new XAttribute("version", "1.1"),
            new XAttribute("width", $"{Invariant(page.WidthMillimeters)}mm"),
            new XAttribute("height", $"{Invariant(page.HeightMillimeters)}mm"),
            new XAttribute(
                "viewBox",
                $"0 0 {Invariant(page.WidthMillimeters)} {Invariant(page.HeightMillimeters)}"));

        var imageIndex = 0;
        var objectIndex = 0;
        var elements = OfdCompositeProjection.Expand(EnumerateElements(page, material, cancellationToken, package.Options.Namespace), material, cancellationToken).ToList();
        OfdRenderPreflight.Validate(elements);
        var renderTexts = new Dictionary<OfdTextElement, OfdTextElement>();
        var previewReserved = new HashSet<OfdElement>();
        void BeforePreviewElement(OfdElement element)
        {
            if (preview is not null && previewReserved.Add(element))
                preview.BeforeElement(element);
        }
        OfdTextElement RenderText(OfdTextElement source)
        {
            if (renderTexts.TryGetValue(source, out var projected)) return projected;
            BeforePreviewElement(source);
            projected = OfdTextRenderProjection.Create(source, material, cancellationToken);
            renderTexts.Add(source, projected);
            return projected;
        }
        var explicitCmaps = new Dictionary<OfdFontResource, OpenTypeCmap>();
        var families = AddEmbeddedFonts(root, svgNs, package.Fonts,
            elements.OfType<OfdTextElement>().Where(OfdGraphicXmlContract.IsVisible).ToList(),
            parseBudget, explicitCmaps, preview, nameOnlyBindings, RenderText, glyphMappingStrategy);
        foreach (var element in elements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!OfdGraphicXmlContract.IsVisible(element)) continue;
            BeforePreviewElement(element);
            if (element is OfdRawElement glyphXml && EmbeddedFontCoverage.HasExplicitGlyphReferences(glyphXml.Xml))
                throw new NotSupportedException("SVG export does not model raw CGTransform glyph substitutions.");
            if (element is OfdRawElement { LocalName: "UnsupportedAnnotationAppearance" })
                throw new NotSupportedException($"Annotation appearance on page {page.Index + 1} contains unsupported drawing; export would lose content.");
            var target = root;
            if (element is not OfdImageElement && (element.Clips is not null || !string.IsNullOrWhiteSpace(element.ClippingXml)))
            {
                var defs = root.Element(svgNs + "defs");
                if (defs is null) { defs = new XElement(svgNs + "defs"); root.AddFirst(defs); }
                var regionIndex = 0;
                foreach (var region in OfdClipGeometry.Read(element, material, cancellationToken))
                {
                    var id = (preview?.Prefix ?? "") + $"object-{objectIndex}-clip-{regionIndex++}";
                    defs.Add(new XElement(svgNs + "clipPath", new XAttribute("id", id), new XAttribute("clipPathUnits", "userSpaceOnUse"),
                        region.Paths.Select(path => new XElement(svgNs + "path", new XAttribute("d", NormalizePathData(path.AbbreviatedData)),
                            new XAttribute("clip-rule", region.EvenOdd ? "evenodd" : "nonzero"),
                            new XAttribute("transform", BuildTransform(element.XMillimeters - page.XMillimeters, element.YMillimeters - page.YMillimeters, path.Transform))))));
                    var group = new XElement(svgNs + "g", new XAttribute("clip-path", $"url(#{id})"));
                    target.Add(group); target = group;
                }
            }
            objectIndex++;
            if (element is OfdTextElement text)
            {
                AddText(target, svgNs, page, RenderText(text), package.Fonts,
                    families, parseBudget, explicitCmaps, cancellationToken, preview, nameOnlyBindings, glyphMappingStrategy);
            }
            else if (element is OfdImageElement image && image.Data.Length > 0)
            {
                AddImage(root, svgNs, page, image, imageIndex++, material, cancellationToken, preview);
            }
            else if (element is OfdPathElement path &&
                !string.IsNullOrWhiteSpace(path.AbbreviatedData))
            {
                var style = OfdPathStyle.Read(path);
                var pathNode = new XElement(
                    svgNs + "path",
                    new XAttribute("fill-rule", style.IsEvenOdd ? "evenodd" : "nonzero"),
                    new XAttribute("d", NormalizePathData(path.AbbreviatedData)),
                    new XAttribute(
                        "stroke",
                        path.Stroke ? ToCssColor(path.StrokeColor) : "none"),
                    new XAttribute(
                        "stroke-width",
                        Invariant(path.LineWidthMillimeters)),
                    new XAttribute(
                        "fill",
                        path.Fill
                            ? ToCssColor(path.FillColor ?? path.StrokeColor)
                            : "none"),
                    new XAttribute("transform", BuildTransform(
                        path.XMillimeters - page.XMillimeters,
                        path.YMillimeters - page.YMillimeters,
                        path.Transform)));
                if (path.Stroke)
                {
                    if (style.Cap is { } cap) pathNode.SetAttributeValue("stroke-linecap", cap);
                    if (style.Join is { } join) pathNode.SetAttributeValue("stroke-linejoin", join);
                    if (style.MiterLimit is { } miter) pathNode.SetAttributeValue("stroke-miterlimit", OfdNumericFormat.Plain(miter));
                    if (style.DashPattern is { } dash)
                    {
                        pathNode.SetAttributeValue("stroke-dasharray", string.Join(" ", dash.Select(OfdNumericFormat.Plain)));
                        pathNode.SetAttributeValue("stroke-dashoffset", OfdNumericFormat.Plain(style.DashOffset));
                    }
                }
                var strokeAlpha = path.StrokeColor.Alpha / 255d * style.Alpha;
                var fillAlpha = (path.FillColor ?? path.StrokeColor).Alpha / 255d * style.Alpha;
                if (path.Stroke && strokeAlpha != 1) pathNode.SetAttributeValue("stroke-opacity", OfdNumericFormat.Plain(strokeAlpha));
                if (path.Fill && fillAlpha != 1) pathNode.SetAttributeValue("fill-opacity", OfdNumericFormat.Plain(fillAlpha));

                target.Add(pathNode);
            }
        }

        return root;
    }

    private static IReadOnlyDictionary<OfdFontResource, (string Family, OpenTypeCmap? Coverage, bool UnboldOutline)> AddEmbeddedFonts(
        XElement root, XNamespace ns, IReadOnlyList<OfdFontResource> fonts, IReadOnlyList<OfdTextElement> texts,
        FontParseBudget parseBudget, IDictionary<OfdFontResource, OpenTypeCmap> explicitCmaps,
        HtmlRenderContext? preview, OfdNameOnlyFontBindings? nameOnlyBindings,
        Func<OfdTextElement, OfdTextElement> renderText, OfdGlyphMappingStrategy glyphMappingStrategy)
    {
        var families = new Dictionary<OfdFontResource, (string Family, OpenTypeCmap? Coverage, bool UnboldOutline)>();
        var emitted = new Dictionary<string, OpenTypeCmap?>(StringComparer.Ordinal);
        var css = new StringBuilder();
        var selected = new HashSet<OfdFontResource>();
        var explicitOnly = new HashSet<OfdFontResource>();
        var selectedTexts = new Dictionary<OfdFontResource, List<OfdTextElement>>();
        foreach (var text in texts)
        {
            var selectedFont = OfdFontSelection.Resolve(fonts, text);
            if (selectedFont is null) continue;
            if (OfdExplicitGlyphRun.Read(renderText(text), selectedFont, explicitCmaps, parseBudget, glyphMappingStrategy) is not null)
            { explicitOnly.Add(selectedFont); continue; }
            if (!selectedTexts.TryGetValue(selectedFont, out var uses)) selectedTexts.Add(selectedFont, uses = new());
            uses.Add(text);
        }
        if (nameOnlyBindings is not null)
            foreach (var text in texts.Where(text => text.Runs.Count == 0 ? !string.IsNullOrEmpty(text.Text) : text.Runs.Any(run => !string.IsNullOrEmpty(run.Text))))
                if (OfdFontSelection.Resolve(fonts, text) is { } font && font.Data.Length == 0) selected.Add(font);
        foreach (var font in fonts.Where(font => font.Data.Length > 0 || selected.Contains(font)))
        {
            // GID paths carry their own geometry and do not need a browser
            // font face or Unicode cmap. A mixed-use face is still emitted.
            if (explicitOnly.Contains(font) && !selectedTexts.ContainsKey(font)) continue;
            parseBudget.Spend(0);
            var cached = font.Data.Length > 0 ? preview?.Font(font, parseBudget) : null;
            byte[] data;
            OpenTypeCmap? coverage = null;
            if (font.Data.Length == 0)
            {
                var bound = nameOnlyBindings!.Resolve(font, parseBudget);
                data = bound.Data;
                coverage = bound.Coverage;
            }
            else data = cached?.Data ?? OpenTypeCollection.SelectFace(font.Data, font.CollectionFaceIndex);
            (bool Bold, bool Italic) style;
            if (font.Data.Length == 0) style = (false, false);
            else if (cached.HasValue) { style = (cached.Value.Bold, cached.Value.Italic); coverage = cached.Value.Coverage; }
            else
            {
                var face = new OpenTypeFace(data); style = face.Style;
                try { coverage = new OpenTypeCmap(face, parseBudget); }
                catch (NotSupportedException) { /* Unused valid encodings may still be emitted for the browser. */ }
            }
            if (selectedTexts.TryGetValue(font, out var cffUses) &&
                data.Length >= 4 && data[0] == 'O' && data[1] == 'T' && data[2] == 'T' && data[3] == 'O')
            {
                var cffFace = new OpenTypeFace(data);
                if (Cff1SvgProfile.IsCffOpenType(cffFace))
                {
                    var count = Cff1SvgProfile.Validate(data, cffFace);
                    coverage ??= new OpenTypeCmap(cffFace, parseBudget);
                    if (coverage.IsSymbol)
                        throw new NotSupportedException("SVG CFF1 text requires a direct Unicode cmap.");
                    Cff1SvgProfile.ValidateVariationGlyphs(coverage, count, parseBudget);
                    foreach (var use in cffUses)
                        Cff1SvgProfile.ValidateSelectedText(renderText(use), font, coverage, count, parseBudget);
                }
            }
            if (coverage?.IsSymbol == true && texts.Any(text => ReferenceEquals(OfdFontSelection.Resolve(fonts, text), font)))
                data = SymbolCmapProfile.CreateUnicodeAlias(data, coverage, parseBudget);
            var family = (preview?.FontPrefix ?? "ofd-font-") + BinaryIdentity.Hash(data);
            var unboldOutline = !style.Bold && OpenTypeFace.HasStaticTrueTypeOutlines(data);
            if (emitted.TryGetValue(family, out var known)) { families[font] = (family, known, unboldOutline); continue; }
            preview?.BeforeFontEmission(data);
            families[font] = (family, coverage, unboldOutline); emitted.Add(family, coverage);
            var openType = data.Length >= 4 && data[0] == 'O' && data[1] == 'T' && data[2] == 'T' && data[3] == 'O';
            css.Append("@font-face{font-family:'").Append(family).Append("';src:url(data:font/")
                .Append(openType ? "otf" : "ttf").Append(";base64,").Append(Convert.ToBase64String(data))
                .Append(") format('").Append(openType ? "opentype" : "truetype").Append("');font-weight:")
                .Append(style.Bold ? "bold" : "normal").Append(";font-style:")
                .Append(style.Italic ? "italic" : "normal").Append(";}\n");
        }
        if (css.Length > 0) root.Add(new XElement(ns + "style", new XAttribute("type", "text/css"), css.ToString()));
        return families;
    }

    private static void AddImage(XElement root, XNamespace ns, OfdPage page, OfdImageElement image, int index, OfdModelMaterialization material, CancellationToken cancellationToken, HtmlRenderContext? preview)
    {
        var projected = preview is null
            ? OfdImageRenderProjection.Prepare(image, 16_000_000, cancellationToken)
            : preview.PreparedImage(image);
        var group = new XElement(ns + "g", new XAttribute("transform", BuildTransform(
            image.XMillimeters - page.XMillimeters, image.YMillimeters - page.YMillimeters, null)),
            new XAttribute("opacity", Invariant(Math.Max(0, Math.Min(255, image.Alpha)) / 255d)));
        var defs = root.Element(ns + "defs");
        if (defs is null) { defs = new XElement(ns + "defs"); root.AddFirst(defs); }
        var boundaryId = (preview?.Prefix ?? "") + $"image-{index}-boundary";
        defs.Add(new XElement(ns + "clipPath", new XAttribute("id", boundaryId), new XAttribute("clipPathUnits", "userSpaceOnUse"),
            new XElement(ns + "rect", new XAttribute("width", Invariant(image.WidthMillimeters)), new XAttribute("height", Invariant(image.HeightMillimeters)))));
        var content = new XElement(ns + "g", new XAttribute("clip-path", $"url(#{boundaryId})"));
        group.Add(content);
        var clipIndex = 0;
        foreach (var region in OfdClipGeometry.Read(image, material, cancellationToken))
        {
            var id = (preview?.Prefix ?? "") + $"image-{index}-clip-{clipIndex++}";
            defs.Add(new XElement(ns + "clipPath", new XAttribute("id", id), new XAttribute("clipPathUnits", "userSpaceOnUse"),
                region.Paths.Select(path => new XElement(ns + "path", new XAttribute("d", NormalizePathData(path.AbbreviatedData)),
                    new XAttribute("clip-rule", region.EvenOdd ? "evenodd" : "nonzero"),
                    new XAttribute("transform", BuildTransform(0, 0, path.Transform))))));
            var clipped = new XElement(ns + "g", new XAttribute("clip-path", $"url(#{id})"));
            content.Add(clipped);
            content = clipped;
        }
        var imageNode = new XElement(ns + "image", new XAttribute("width", "1"), new XAttribute("height", "1"),
            new XAttribute("preserveAspectRatio", "none"),
            new XAttribute("transform", BuildTransform(0, 0, image.Transform ?? new double[] { image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0 })),
            new XAttribute("href", $"data:{projected.MediaType};base64,{Convert.ToBase64String(projected.Data)}"));
        var interpolate = OfdImageRenderingHints.ReadPdfInterpolate(image);
        if (interpolate.HasValue) imageNode.SetAttributeValue("image-rendering", interpolate.Value ? "smooth" : "crisp-edges");
        content.Add(imageNode);
        root.Add(group);
    }

    private static void AddText(
        XElement root,
        XNamespace svgNs,
        OfdPage page,
        OfdTextElement text,
        IReadOnlyList<OfdFontResource> fonts,
        IReadOnlyDictionary<OfdFontResource, (string Family, OpenTypeCmap? Coverage, bool UnboldOutline)> families,
        FontParseBudget parseBudget, IDictionary<OfdFontResource, OpenTypeCmap> explicitCmaps,
        CancellationToken cancellationToken, HtmlRenderContext? preview,
        OfdNameOnlyFontBindings? nameOnlyBindings, OfdGlyphMappingStrategy glyphMappingStrategy)
    {
        var resource = OfdFontSelection.Resolve(fonts, text);
        if (OfdExplicitGlyphRun.Read(text, resource, explicitCmaps, parseBudget, glyphMappingStrategy) is { } explicitRun)
        {
            AddExplicitGlyphs(root, svgNs, page, text, explicitRun,
                OfdExplicitGlyphFont.Read(resource, text, explicitRun, embedFont: false), cancellationToken, preview);
            return;
        }
        if (EmbeddedFontCoverage.HasExplicitGlyphReferences(text))
            throw new NotSupportedException("SVG export cannot render this CGTransform source.");
        if (nameOnlyBindings is not null &&
            (text.Runs.Count == 0 ? !string.IsNullOrEmpty(text.Text) : text.Runs.Any(run => !string.IsNullOrEmpty(run.Text))) &&
            (resource is null || resource.Data.Length == 0 && !families.ContainsKey(resource)))
            throw new NotSupportedException($"Selected name-only font '{text.FontName}' has no explicit face binding.");
        if (nameOnlyBindings is not null && resource?.Data.Length == 0)
            OfdNameOnlyFontBindings.ValidateBmpText(text);
        // Resource names can distinguish styles (e.g. native DOCX's "|bold"
        // identity). CSS needs the declared family, with style carried below.
        var family = !string.IsNullOrWhiteSpace(resource?.FamilyName)
            ? resource!.FamilyName! : text.FontName;
        var knownUnboldOutline = false;
        if (resource is not null && families.TryGetValue(resource, out var embedded))
        {
            family = embedded.Family;
            knownUnboldOutline = embedded.UnboldOutline;
            if (embedded.Coverage is null) throw new NotSupportedException("Selected embedded font has unmodeled cmap character semantics; the original OFD font remains preserved.");
            if (embedded.Coverage.IsSymbol) SymbolCmapProfile.ValidateText(text, embedded.Coverage, parseBudget, cancellationToken);
            else EmbeddedFontCoverage.Validate(text, embedded.Coverage, resource.FontName, cancellationToken);
        }
        var fontWeight = resource?.Bold == true || text.Weight >= 600 ? "bold" : "normal";
        var fontStyle = resource?.Italic == true || text.Italic ? "italic" : "normal";
        var paint = OfdTextPaint.Read(text);
        if (fontWeight == "bold" && knownUnboldOutline && paint.Fill && paint.Stroke &&
            paint.FillColor.Alpha == 255 && paint.StrokeColor.Alpha == 255 &&
            paint.FillColor.Red == paint.StrokeColor.Red && paint.FillColor.Green == paint.StrokeColor.Green &&
            paint.FillColor.Blue == paint.StrokeColor.Blue)
        {
            var value = text.Runs.Count == 0 ? text.Text : string.Concat(text.Runs.Select(run => run.Text));
            // The declared contour already supplies the existing script-aware
            // emphasis. Avoid a second host synthesis; retain thin/hairline
            // strokes and unknown faces on their original route. Compare on
            // the writer's serialized grid without changing the stroke itself.
            var emphasis = double.Parse(Invariant(FontEmphasisMetrics.BoldStrokeWidth(text.FontSizeMillimeters, value)), CultureInfo.InvariantCulture);
            if (emphasis > 0 && paint.LineWidthMillimeters >= emphasis) fontWeight = "normal";
        }
        var drawing = OfdTextEmphasis.DrawingTransform(text);
        var transform = BuildTransform(
            text.XMillimeters - page.XMillimeters,
            text.YMillimeters - page.YMillimeters,
            paint.HScale == 1 ? drawing.Matrix : OfdCompositeProjection.Multiply(
                drawing.Matrix ?? OfdCompositeProjection.Identity(), new[] { paint.HScale, 0d, 0d, 1d, 0d, 0d }));
        void ApplyPaint(XElement node, string value)
        {
            node.SetAttributeValue("fill", paint.Fill ? ToCssColor(paint.FillColor) : "none");
            node.SetAttributeValue("stroke", paint.Stroke && paint.StrokeColor.Alpha > 0 ? ToCssColor(paint.StrokeColor) : "none");
            if (paint.Stroke)
            {
                node.SetAttributeValue("stroke-width", Invariant(paint.LineWidthMillimeters));
                if (paint.StrokeColor.Alpha != 255) node.SetAttributeValue("stroke-opacity", Invariant(paint.StrokeColor.Alpha / 255d));
            }
            if (paint.FillColor.Alpha != 255) node.SetAttributeValue("fill-opacity", Invariant(paint.FillColor.Alpha / 255d));
            if (OfdTextGeometry.IsWhiteSpace(value, cancellationToken))
            { node.SetAttributeValue("fill-opacity", "0"); node.SetAttributeValue("stroke-opacity", "0"); }
        }
        if (text.Runs.Count == 0)
        {
            var anchor = OfdTextEmphasis.Anchor(0, text.FontSizeMillimeters, drawing.Factor);
            var node = new XElement(
                svgNs + "text",
                new XAttribute("x", OfdNumericFormat.Plain(anchor.X / paint.HScale)),
                new XAttribute("y", OfdNumericFormat.Plain(anchor.Y)),
                new XAttribute("font-family", family),
                new XAttribute("font-weight", fontWeight),
                new XAttribute("font-style", fontStyle),
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                new XAttribute("style", "white-space:pre"),
                new XAttribute("font-size", Invariant(text.FontSizeMillimeters)),
                new XAttribute("transform", transform),
                text.Text);
            ApplyPaint(node, text.Text);
            root.Add(node);
            return;
        }

        foreach (var run in text.Runs)
        {
            var anchor = OfdTextEmphasis.Anchor(run.XMillimeters, run.YMillimeters, drawing.Factor);
            var node = new XElement(
                svgNs + "text",
                new XAttribute("x", OfdNumericFormat.Plain(anchor.X / paint.HScale)),
                new XAttribute("y", OfdNumericFormat.Plain(anchor.Y)),
                new XAttribute("font-family", family),
                new XAttribute("font-weight", fontWeight),
                new XAttribute("font-style", fontStyle),
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                new XAttribute("style", "white-space:pre"),
                new XAttribute("font-size", Invariant(text.FontSizeMillimeters)),
                new XAttribute("transform", transform),
                run.Text);
            var deltas = OfdTextGeometry.PositioningDeltas(run, cancellationToken);
            var deltaX = deltas.X; var deltaY = deltas.Y;
            if (deltaX.Count > 0 || deltaY.Count > 0)
            {
                // OFD deltas locate adjacent origins. SVG dx/dy add to the
                // browser's own advance, so use explicit glyph origins instead.
                node.RemoveNodes();
                var x = run.XMillimeters;
                var y = run.YMillimeters;
                var index = 0;
                foreach (var glyph in OfdTextGeometry.EnumerateGlyphs(run.Text, cancellationToken))
                {
                    var glyphAnchor = OfdTextEmphasis.Anchor(x, y, drawing.Factor);
                    var span = new XElement(svgNs + "tspan", new XAttribute("x", OfdNumericFormat.Plain(glyphAnchor.X / paint.HScale)),
                        new XAttribute("y", OfdNumericFormat.Plain(glyphAnchor.Y)), glyph);
                    if (OfdTextGeometry.IsWhiteSpace(glyph, cancellationToken))
                    { span.SetAttributeValue("fill-opacity", "0"); span.SetAttributeValue("stroke-opacity", "0"); }
                    node.Add(span);
                    if (index < deltaX.Count) x += deltaX[index];
                    if (index < deltaY.Count) y += deltaY[index];
                    index++;
                }
            }

            ApplyPaint(node, run.Text);
            root.Add(node);
        }
    }

    private static void AddExplicitGlyphs(XElement root, XNamespace ns, OfdPage page,
        OfdTextElement text, OfdExplicitGlyphRun run, OfdExplicitGlyphFont font, CancellationToken token,
        HtmlRenderContext? preview)
    {
        if (run.Groups is { } groups)
        {
            foreach (var plannedGroup in groups)
                AddExplicitGlyphs(root, ns, page, text, plannedGroup.Projection,
                    font.Slice(plannedGroup.GlyphOffset, plannedGroup.Projection.GlyphIds.Count), token, preview);
            return;
        }
        var paint = OfdTextPaint.Read(text);
        var drawing = OfdTextEmphasis.DrawingTransform(text);
        var transform = BuildTransform(text.XMillimeters - page.XMillimeters,
            text.YMillimeters - page.YMillimeters,
            paint.HScale == 1 ? drawing.Matrix : OfdCompositeProjection.Multiply(
                drawing.Matrix ?? OfdCompositeProjection.Identity(),
                new[] { paint.HScale, 0d, 0d, 1d, 0d, 0d }));
        var group = new XElement(ns + "g", new XAttribute("transform", transform),
            new XAttribute("role", "img"), new XAttribute("aria-label", run.LogicalText),
            new XAttribute("data-ofd-text", run.LogicalText),
            new XElement(ns + "title", run.LogicalText));
        var pathOrigins = run.PathGlyphSourceSpan is null ? null : run.PathGlyphOrigins(token);
        var globalIndex = 0;
        foreach (var textRun in run.Runs)
        {
            var glyphCount = run.SingleGlyphSourceSpan is not null ? 1 :
                pathOrigins is not null ? pathOrigins.Count : textRun.Text.Length;
            var hasX = !string.IsNullOrWhiteSpace(textRun.DeltaX);
            var hasY = !string.IsNullOrWhiteSpace(textRun.DeltaY);
            var deltaX = hasX ? OfdTextGeometry.ExpandDeltas(textRun.DeltaX, glyphCount - 1, token) : [];
            var deltaY = hasY ? OfdTextGeometry.ExpandDeltas(textRun.DeltaY, glyphCount - 1, token) : [];
            var x = textRun.XMillimeters; var y = textRun.YMillimeters;
            for (var index = 0; index < glyphCount; index++, globalIndex++)
            {
                token.ThrowIfCancellationRequested();
                if (pathOrigins is not null)
                {
                    x = pathOrigins[index].X;
                    y = pathOrigins[index].Y;
                }
                var outline = font.Outlines[globalIndex];
                preview?.BeforeExplicitGlyphOutline(outline);
                var anchor = OfdTextEmphasis.Anchor(x, y, drawing.Factor);
                var scale = text.FontSizeMillimeters / outline.UnitsPerEm;
                var path = new StringBuilder();
                foreach (var command in outline.ToPathCommands())
                {
                    double px = anchor.X / paint.HScale + command.X * scale;
                    double py = anchor.Y - command.Y * scale;
                    switch (command.Verb)
                    {
                        case TrueTypePathVerb.Move: path.Append("M ").Append(Invariant(px)).Append(' ').Append(Invariant(py)).Append(' '); break;
                        case TrueTypePathVerb.Line: path.Append("L ").Append(Invariant(px)).Append(' ').Append(Invariant(py)).Append(' '); break;
                        case TrueTypePathVerb.Quadratic:
                            path.Append("Q ").Append(Invariant(anchor.X / paint.HScale + command.ControlX * scale)).Append(' ')
                                .Append(Invariant(anchor.Y - command.ControlY * scale)).Append(' ')
                                .Append(Invariant(px)).Append(' ').Append(Invariant(py)).Append(' ');
                            break;
                        case TrueTypePathVerb.Close: path.Append("Z "); break;
                    }
                }
                var node = new XElement(ns + "path", new XAttribute("d", path.ToString().TrimEnd()),
                    new XAttribute("data-original-gid", run.GlyphIds[globalIndex]),
                    new XAttribute("fill", paint.Fill ? ToCssColor(paint.FillColor) : "none"),
                    new XAttribute("stroke", paint.Stroke && paint.StrokeColor.Alpha > 0 ? ToCssColor(paint.StrokeColor) : "none"));
                if (paint.Stroke) node.SetAttributeValue("stroke-width", Invariant(paint.LineWidthMillimeters));
                if (paint.FillColor.Alpha != 255) node.SetAttributeValue("fill-opacity", Invariant(paint.FillColor.Alpha / 255d));
                if (paint.StrokeColor.Alpha != 255) node.SetAttributeValue("stroke-opacity", Invariant(paint.StrokeColor.Alpha / 255d));
                group.Add(node);
                if (index + 1 == glyphCount || pathOrigins is not null) continue;
                if (hasX) x += deltaX[index];
                if (hasY) y += deltaY[index];
            }
        }
        root.Add(group);
    }

    private static string NormalizePathData(string data)
    {
        var tokens = PathTokenPattern.Matches(data)
            .Cast<Match>()
            .Select(match =>
            {
                if (match.Value == "B")
                {
                    return "C";
                }

                if (match.Value == "b")
                {
                    return "c";
                }

                if (match.Value == "C")
                {
                    return "Z";
                }

                return match.Value == "c" ? "z" : match.Value;
            });
        return string.Join(" ", tokens);
    }

    private static IEnumerable<OfdElement> EnumerateElements(OfdPage page, OfdModelMaterialization material, CancellationToken token, string documentNamespace)
    {
        return page.Templates
            .Where(template => string.Equals(
                template.ZOrder,
                "Background",
                StringComparison.OrdinalIgnoreCase))
            .SelectMany(template => template.Elements)
            .Concat(page.Elements)
            .Concat(page.Templates
                .Where(template => !string.Equals(
                    template.ZOrder,
                    "Background",
                    StringComparison.OrdinalIgnoreCase))
                .SelectMany(template => template.Elements))
            .Concat(OfdAnnotationProjection.Resolve(page, material, token, documentNamespace: documentNamespace).Select(occurrence => occurrence.RenderElement));
    }

    private static string BuildTransform(
        double x,
        double y,
        double[]? matrix)
    {
        var translation = $"translate({OfdNumericFormat.Plain(x)} {OfdNumericFormat.Plain(y)})";
        return matrix is { Length: 6 }
            ? $"{translation} matrix({string.Join(" ", matrix.Select(OfdNumericFormat.Plain))})"
            : translation;
    }

    private static string ToCssColor(OfdColor color)
    {
        return $"rgb({color.Red},{color.Green},{color.Blue})";
    }

    private static string Invariant(double value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }
}
