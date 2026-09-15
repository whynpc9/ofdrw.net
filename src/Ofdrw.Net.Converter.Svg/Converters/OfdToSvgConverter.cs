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

    public OfdToSvgConverter() : this(new OfdPackageLoadOptions()) { }

    public OfdToSvgConverter(OfdPackageLoadOptions loadOptions)
    {
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

        var svg = BuildSvg(package, pages[pageIndex], cancellationToken);
        var bytes = Encoding.UTF8.GetBytes(
            new XDocument(new XDeclaration("1.0", "utf-8", null), svg)
                .ToString(SaveOptions.DisableFormatting));
        await svgOutput.WriteAsync(bytes, 0, bytes.Length, cancellationToken)
            .ConfigureAwait(false);
        await svgOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static XElement BuildSvg(OfdDocumentPackage package, OfdPage page, CancellationToken cancellationToken)
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

        var families = AddEmbeddedFonts(root, svgNs, package.Fonts);
        var imageIndex = 0;
        foreach (var element in EnumerateElements(page))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (element is OfdTextElement text)
            {
                AddText(root, svgNs, page, text, package.Fonts, families);
            }
            else if (element is OfdImageElement image && image.Data.Length > 0)
            {
                AddImage(root, svgNs, page, image, imageIndex++);
            }
            else if (element is OfdPathElement path &&
                !string.IsNullOrWhiteSpace(path.AbbreviatedData))
            {
                var pathNode = new XElement(
                    svgNs + "path",
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
                if (path.Stroke && path.StrokeColor.Alpha != 255)
                {
                    pathNode.SetAttributeValue(
                        "stroke-opacity",
                        Invariant(path.StrokeColor.Alpha / 255d));
                }

                if (path.Fill && (path.FillColor ?? path.StrokeColor).Alpha != 255)
                {
                    pathNode.SetAttributeValue(
                        "fill-opacity",
                        Invariant((path.FillColor ?? path.StrokeColor).Alpha / 255d));
                }

                root.Add(pathNode);
            }
        }

        return root;
    }

    private static IReadOnlyDictionary<OfdFontResource, string> AddEmbeddedFonts(
        XElement root, XNamespace ns, IReadOnlyList<OfdFontResource> fonts)
    {
        var families = new Dictionary<OfdFontResource, string>();
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        var css = new StringBuilder();
        foreach (var font in fonts.Where(font => font.Data.Length > 0))
        {
            var family = "ofd-font-" + BinaryIdentity.Hash(font.Data);
            families[font] = family;
            if (!emitted.Add(family)) continue;
            var style = OfdFontStyle.Read(font.Data);
            var openType = font.Data.Length >= 4 && font.Data[0] == 'O' && font.Data[1] == 'T' && font.Data[2] == 'T' && font.Data[3] == 'O';
            css.Append("@font-face{font-family:'").Append(family).Append("';src:url(data:font/")
                .Append(openType ? "otf" : "ttf").Append(";base64,").Append(Convert.ToBase64String(font.Data))
                .Append(") format('").Append(openType ? "opentype" : "truetype").Append("');font-weight:")
                .Append(style.Bold ? "bold" : "normal").Append(";font-style:")
                .Append(style.Italic ? "italic" : "normal").Append(";}\n");
        }
        if (css.Length > 0) root.Add(new XElement(ns + "style", new XAttribute("type", "text/css"), css.ToString()));
        return families;
    }

    private static void AddImage(XElement root, XNamespace ns, OfdPage page, OfdImageElement image, int index)
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
        foreach (var region in OfdClipGeometry.Read(image.ClipsXml))
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
        content.Add(new XElement(ns + "image", new XAttribute("width", "1"), new XAttribute("height", "1"),
            new XAttribute("preserveAspectRatio", "none"),
            new XAttribute("transform", BuildTransform(0, 0, image.Transform ?? new double[] { image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0 })),
            new XAttribute("href", $"data:{image.MediaType};base64,{Convert.ToBase64String(image.Data)}")));
        root.Add(group);
    }

    private static void AddText(
        XElement root,
        XNamespace svgNs,
        OfdPage page,
        OfdTextElement text,
        IReadOnlyList<OfdFontResource> fonts,
        IReadOnlyDictionary<OfdFontResource, string> families)
    {
        var resource = fonts.FirstOrDefault(font => !string.IsNullOrEmpty(text.FontResourceId) && font.Id == text.FontResourceId)
            ?? fonts.FirstOrDefault(font => string.Equals(font.FontName, text.FontName, StringComparison.OrdinalIgnoreCase) && !font.Bold && !font.Italic)
            ?? fonts.FirstOrDefault(font => string.Equals(font.FontName, text.FontName, StringComparison.OrdinalIgnoreCase));
        var family = resource is not null && families.TryGetValue(resource, out var embeddedFamily) ? embeddedFamily : text.FontName;
        var transform = BuildTransform(
            text.XMillimeters - page.XMillimeters,
            text.YMillimeters - page.YMillimeters,
            text.Transform);
        if (text.Runs.Count == 0)
        {
            var node = new XElement(
                svgNs + "text",
                new XAttribute("x", "0"),
                new XAttribute("y", Invariant(text.FontSizeMillimeters)),
                new XAttribute("font-family", family),
                new XAttribute("font-weight", resource?.Bold == true ? "bold" : "normal"),
                new XAttribute("font-style", resource?.Italic == true ? "italic" : "normal"),
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

            root.Add(node);
            return;
        }

        foreach (var run in text.Runs)
        {
            var node = new XElement(
                svgNs + "text",
                new XAttribute("x", Invariant(run.XMillimeters)),
                new XAttribute("y", Invariant(run.YMillimeters)),
                new XAttribute("font-family", family),
                new XAttribute("font-weight", resource?.Bold == true ? "bold" : "normal"),
                new XAttribute("font-style", resource?.Italic == true ? "italic" : "normal"),
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
                    node.Add(new XElement(svgNs + "tspan", new XAttribute("x", Invariant(x)),
                        new XAttribute("y", Invariant(y)), glyphs[index]));
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

    private static IEnumerable<OfdElement> EnumerateElements(OfdPage page)
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
                .SelectMany(template => template.Elements));
    }

    private static string BuildTransform(
        double x,
        double y,
        double[]? matrix)
    {
        var translation = $"translate({Invariant(x)} {Invariant(y)})";
        return matrix is { Length: 6 }
            ? $"{translation} matrix({string.Join(" ", matrix.Select(Invariant))})"
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
