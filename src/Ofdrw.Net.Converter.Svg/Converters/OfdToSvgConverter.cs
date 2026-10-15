using Ofdrw.Net.Core.Fonts;
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
    private readonly long _maximumFontParseWork;

    public OfdToSvgConverter() : this(new OfdPackageLoadOptions()) { }

    public OfdToSvgConverter(OfdPackageLoadOptions loadOptions) : this(loadOptions, FontParseBudget.DefaultMaximumWork) { }

    /// <summary>Creates an SVG exporter with a finite cumulative cmap parse budget.</summary>
    /// <param name="loadOptions">OFD loading limits.</param>
    /// <param name="maximumCmapParseWork">Positive cumulative limit for cmap bytes, parsed records and entries inspected by variation coverage queries.</param>
    public OfdToSvgConverter(OfdPackageLoadOptions loadOptions, long maximumCmapParseWork)
    {
        if (maximumCmapParseWork <= 0) throw new ArgumentOutOfRangeException(nameof(maximumCmapParseWork));
        _maximumFontParseWork = maximumCmapParseWork;
        _loadOptions = loadOptions ?? throw new ArgumentNullException(nameof(loadOptions));
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
        var svg = BuildSvg(package, pages[pageIndex], material, new FontParseBudget(_maximumFontParseWork, cancellationToken), cancellationToken);
        var bytes = Encoding.UTF8.GetBytes(
            new XDocument(new XDeclaration("1.0", "utf-8", null), svg)
                .ToString(SaveOptions.DisableFormatting));
        await svgOutput.WriteAsync(bytes, 0, bytes.Length, cancellationToken)
            .ConfigureAwait(false);
        await svgOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static XElement BuildSvg(OfdDocumentPackage package, OfdPage page, OfdModelMaterialization material, FontParseBudget parseBudget, CancellationToken cancellationToken)
    {
        var svgNs = XNamespace.Get("http://www.w3.org/2000/svg");
        var root = new XElement(
            svgNs + "svg",
            new XAttribute("version", "1.1"),
            new XAttribute("width", $"{Invariant(page.WidthMillimeters)}mm"),
            new XAttribute("height", $"{Invariant(page.HeightMillimeters)}mm"),
            new XAttribute(
                "viewBox",
                $"0 0 {Invariant(page.WidthMillimeters)} {Invariant(page.HeightMillimeters)}"));

        var families = AddEmbeddedFonts(root, svgNs, package.Fonts, parseBudget);
        var imageIndex = 0;
        var objectIndex = 0;
        foreach (var element in OfdCompositeProjection.Expand(EnumerateElements(page, material, cancellationToken, package.Options.Namespace), material, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!OfdGraphicXmlContract.IsVisible(element)) continue;
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
                    var id = $"object-{objectIndex}-clip-{regionIndex++}";
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
                AddText(target, svgNs, page, OfdTextRenderProjection.Create(text, material, cancellationToken), package.Fonts, families, cancellationToken);
            }
            else if (element is OfdImageElement image && image.Data.Length > 0)
            {
                AddImage(root, svgNs, page, image, imageIndex++, material, cancellationToken);
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

    private static IReadOnlyDictionary<OfdFontResource, (string Family, OpenTypeCmap? Coverage)> AddEmbeddedFonts(
        XElement root, XNamespace ns, IReadOnlyList<OfdFontResource> fonts, FontParseBudget parseBudget)
    {
        var families = new Dictionary<OfdFontResource, (string Family, OpenTypeCmap? Coverage)>();
        var emitted = new Dictionary<string, OpenTypeCmap?>(StringComparer.Ordinal);
        var css = new StringBuilder();
        foreach (var font in fonts.Where(font => font.Data.Length > 0))
        {
            parseBudget.Spend(0);
            var data = OpenTypeCollection.SelectFace(font.Data, font.CollectionFaceIndex);
            var family = "ofd-font-" + BinaryIdentity.Hash(data);
            if (emitted.TryGetValue(family, out var known)) { families[font] = (family, known); continue; }
            var face = new OpenTypeFace(data); OpenTypeCmap? coverage = null;
            try { coverage = new OpenTypeCmap(face, parseBudget); }
            catch (NotSupportedException) { /* Unused valid encodings may still be emitted for the browser. */ }
            families[font] = (family, coverage); emitted.Add(family, coverage);
            var style = face.Style;
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

    private static void AddImage(XElement root, XNamespace ns, OfdPage page, OfdImageElement image, int index, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        var group = new XElement(ns + "g", new XAttribute("transform", BuildTransform(
            image.XMillimeters - page.XMillimeters, image.YMillimeters - page.YMillimeters, null)),
            new XAttribute("opacity", Invariant(Math.Max(0, Math.Min(255, image.Alpha)) / 255d)));
        var defs = root.Element(ns + "defs");
        if (defs is null) { defs = new XElement(ns + "defs"); root.AddFirst(defs); }
        var boundaryId = $"image-{index}-boundary";
        defs.Add(new XElement(ns + "clipPath", new XAttribute("id", boundaryId), new XAttribute("clipPathUnits", "userSpaceOnUse"),
            new XElement(ns + "rect", new XAttribute("width", Invariant(image.WidthMillimeters)), new XAttribute("height", Invariant(image.HeightMillimeters)))));
        var content = new XElement(ns + "g", new XAttribute("clip-path", $"url(#{boundaryId})"));
        group.Add(content);
        var clipIndex = 0;
        foreach (var region in OfdClipGeometry.Read(image, material, cancellationToken))
        {
            var id = $"image-{index}-clip-{clipIndex++}";
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
            new XAttribute("href", $"data:{image.MediaType};base64,{Convert.ToBase64String(image.Data)}"));
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
        IReadOnlyDictionary<OfdFontResource, (string Family, OpenTypeCmap? Coverage)> families, CancellationToken cancellationToken)
    {
        if (EmbeddedFontCoverage.HasExplicitGlyphReferences(text))
            throw new NotSupportedException("SVG export does not model CGTransform glyph substitutions; the OFD font and XML remain preserved.");
        var resource = OfdFontSelection.Resolve(fonts, text);
        var family = text.FontName;
        if (resource is not null && families.TryGetValue(resource, out var embedded))
        {
            family = embedded.Family;
            if (embedded.Coverage is null || embedded.Coverage.IsSymbol) throw new NotSupportedException("Selected embedded font has unmodeled cmap character semantics; the original OFD font remains preserved.");
            EmbeddedFontCoverage.Validate(text, embedded.Coverage, resource.FontName, cancellationToken);
        }
        var fontWeight = resource?.Bold == true || text.Weight >= 600 ? "bold" : "normal";
        var fontStyle = resource?.Italic == true || text.Italic ? "italic" : "normal";
        var drawing = OfdTextEmphasis.DrawingTransform(text);
        var transform = BuildTransform(
            text.XMillimeters - page.XMillimeters,
            text.YMillimeters - page.YMillimeters,
            drawing.Matrix);
        if (text.Runs.Count == 0)
        {
            var anchor = OfdTextEmphasis.Anchor(0, text.FontSizeMillimeters, drawing.Factor);
            var node = new XElement(
                svgNs + "text",
                new XAttribute("x", OfdNumericFormat.Plain(anchor.X)),
                new XAttribute("y", OfdNumericFormat.Plain(anchor.Y)),
                new XAttribute("font-family", family),
                new XAttribute("font-weight", fontWeight),
                new XAttribute("font-style", fontStyle),
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                new XAttribute("style", "white-space:pre"),
                new XAttribute("font-size", Invariant(text.FontSizeMillimeters)),
                new XAttribute("fill", ToCssColor(text.FillColor)),
                new XAttribute("transform", transform),
                text.Text);
            if (text.FillColor.Alpha != 255)
            {
                node.SetAttributeValue(
                    "fill-opacity",
                    Invariant(text.FillColor.Alpha / 255d));
            }

            if (text.Text.All(char.IsWhiteSpace)) node.SetAttributeValue("fill-opacity", "0");
            root.Add(node);
            return;
        }

        foreach (var run in text.Runs)
        {
            var anchor = OfdTextEmphasis.Anchor(run.XMillimeters, run.YMillimeters, drawing.Factor);
            var node = new XElement(
                svgNs + "text",
                new XAttribute("x", OfdNumericFormat.Plain(anchor.X)),
                new XAttribute("y", OfdNumericFormat.Plain(anchor.Y)),
                new XAttribute("font-family", family),
                new XAttribute("font-weight", fontWeight),
                new XAttribute("font-style", fontStyle),
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                new XAttribute("style", "white-space:pre"),
                new XAttribute("font-size", Invariant(text.FontSizeMillimeters)),
                new XAttribute("fill", ToCssColor(text.FillColor)),
                new XAttribute("transform", transform),
                run.Text);
            var glyphs = OfdTextGeometry.Glyphs(run.Text);
            var deltaX = OfdTextGeometry.ExpandDeltas(run.DeltaX, Math.Max(0, glyphs.Count - 1));
            var deltaY = OfdTextGeometry.ExpandDeltas(run.DeltaY, Math.Max(0, glyphs.Count - 1));
            if (deltaX.Count > 0 || deltaY.Count > 0)
            {
                // OFD deltas locate adjacent origins. SVG dx/dy add to the
                // browser's own advance, so use explicit glyph origins instead.
                node.RemoveNodes();
                var x = run.XMillimeters;
                var y = run.YMillimeters;
                for (var index = 0; index < glyphs.Count; index++)
                {
                    var glyphAnchor = OfdTextEmphasis.Anchor(x, y, drawing.Factor);
                    var span = new XElement(svgNs + "tspan", new XAttribute("x", OfdNumericFormat.Plain(glyphAnchor.X)),
                        new XAttribute("y", OfdNumericFormat.Plain(glyphAnchor.Y)), glyphs[index]);
                    if (glyphs[index].All(char.IsWhiteSpace)) span.SetAttributeValue("fill-opacity", "0");
                    node.Add(span);
                    if (index < deltaX.Count) x += deltaX[index];
                    if (index < deltaY.Count) y += deltaY[index];
                }
            }

            if (text.FillColor.Alpha != 255)
            {
                node.SetAttributeValue(
                    "fill-opacity",
                    Invariant(text.FillColor.Alpha / 255d));
            }

            if (run.Text.All(char.IsWhiteSpace)) node.SetAttributeValue("fill-opacity", "0");
            root.Add(node);
        }
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
