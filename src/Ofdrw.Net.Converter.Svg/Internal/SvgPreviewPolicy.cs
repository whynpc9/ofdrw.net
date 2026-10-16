using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Converter.Svg.Internal;

// Validates only our own generated DOM. This is intentionally not an arbitrary SVG sanitizer.
internal static class SvgPreviewPolicy
{
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private static readonly IReadOnlyDictionary<string, string[]> Attributes = new Dictionary<string, string[]>
    {
        ["svg"] = new[] { "version", "width", "height", "viewBox" }, ["defs"] = Array.Empty<string>(),
        ["g"] = new[] { "transform", "opacity", "clip-path" }, ["clipPath"] = new[] { "id", "clipPathUnits" },
        ["path"] = new[] { "d", "transform", "clip-rule", "fill-rule", "stroke", "fill", "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-miterlimit", "stroke-dasharray", "stroke-dashoffset", "stroke-opacity", "fill-opacity" },
        ["rect"] = new[] { "width", "height" },
        ["image"] = new[] { "width", "height", "preserveAspectRatio", "transform", "href", "image-rendering" },
        ["text"] = new[] { "x", "y", "font-family", "font-size", "font-weight", "font-style", "space", "style", "fill", "fill-opacity", "transform" },
        ["tspan"] = new[] { "x", "y", "fill-opacity" }, ["style"] = new[] { "type" }
    };
    private static readonly Regex FontCss = new(@"\G@font-face\{font-family:'(?<family>[a-z0-9-]+)';src:url\(data:font/(?:ttf|otf);base64,[A-Za-z0-9+/]+={0,2}\) format\('(?:truetype|opentype)'\);font-weight:(?:normal|bold);font-style:(?:normal|italic);\}\n", RegexOptions.CultureInvariant);
    private static readonly Regex Color = new(@"\Argb\((?:[0-9]{1,3},){2}[0-9]{1,3}\)\z", RegexOptions.CultureInvariant);
    private static readonly Regex Transform = new(@"\Atranslate\([^()]+\)(?: matrix\([^()]+\))?\z", RegexOptions.CultureInvariant);
    private static readonly Regex Path = new(@"\A[MLHVCSQTAZmlhvcsqtaz0-9eE+.,\s-]*\z", RegexOptions.CultureInvariant);
    internal static void Validate(XElement root, string prefix, string fontPrefix, CancellationToken token)
    {
        var ids = new Dictionary<string, XElement>(StringComparer.Ordinal);
        var references = new List<string>(); var fonts = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in root.DescendantsAndSelf())
        {
            token.ThrowIfCancellationRequested();
            if (node.Name.Namespace != Svg) Fail();
            if (!Attributes.TryGetValue(node.Name.LocalName, out var allowed)) Fail();
            foreach (var attribute in node.Attributes())
            {
                token.ThrowIfCancellationRequested();
                if (attribute.IsNamespaceDeclaration || !allowed!.Contains(attribute.Name.LocalName) ||
                    (attribute.Name.Namespace != XNamespace.None && attribute.Name != XNamespace.Xml + "space")) Fail();
                var value = attribute.Value; HtmlRenderContext.Characters(value);
                switch (attribute.Name.LocalName)
                {
                    case "id": if (!value.StartsWith(prefix, StringComparison.Ordinal) || !ids.TryAddCompat(value, node)) Fail(); break;
                    case "clip-path":
                        if (!value.StartsWith("url(#" + prefix, StringComparison.Ordinal) || !value.EndsWith(")", StringComparison.Ordinal)) Fail();
                        references.Add(value.Substring(5, value.Length - 6)); break;
                    case "href":
                        if (!(value.StartsWith("data:image/png;base64,", StringComparison.Ordinal) || value.StartsWith("data:image/jpeg;base64,", StringComparison.Ordinal))) Fail();
                        break;
                    case "style": if (value != "white-space:pre") Fail(); break;
                    case "type": if (value != "text/css") Fail(); break;
                    case "space": if (value != "preserve") Fail(); break;
                    case "version": if (value != "1.1") Fail(); break;
                    case "clipPathUnits": if (value != "userSpaceOnUse") Fail(); break;
                    case "preserveAspectRatio": if (value != "none") Fail(); break;
                    case "image-rendering": if (value != "smooth" && value != "crisp-edges") Fail(); break;
                    case "font-weight": if (value != "bold" && value != "normal") Fail(); break;
                    case "font-style": if (value != "italic" && value != "normal") Fail(); break;
                    case "font-family": break; // presentation attribute, never a CSS source or URI
                    case "fill": case "stroke": if (value != "none" && !Color.IsMatch(value)) Fail(); break;
                    case "fill-rule": case "clip-rule": if (value != "nonzero" && value != "evenodd") Fail(); break;
                    case "stroke-linecap": if (value != "butt" && value != "round" && value != "square") Fail(); break;
                    case "stroke-linejoin": if (value != "miter" && value != "round" && value != "bevel") Fail(); break;
                    case "d": if (!Path.IsMatch(value)) Fail(); FinitePath(value); break;
                    case "transform":
                        if (!Transform.IsMatch(value)) Fail();
                        Numbers(value.Replace("translate(", "").Replace("matrix(", "").Replace(")", "")); break;
                    case "opacity": case "fill-opacity": case "stroke-opacity":
                        var alpha = Number(value); if (alpha < 0 || alpha > 1) Fail(); break;
                    case "width": case "height":
                        if (value.EndsWith("mm", StringComparison.Ordinal)) value = value.Substring(0, value.Length - 2);
                        if (Number(value) <= 0) Fail(); break;
                    case "viewBox": if (Numbers(value) != 4) Fail(); break;
                    case "stroke-dasharray": Numbers(value); break;
                    default: Number(value); break;
                }
            }
            foreach (var child in node.Nodes())
            {
                if (child is XElement) continue;
                if (child is not XText text || node.Name.LocalName is not ("text" or "tspan" or "style")) Fail();
                else HtmlRenderContext.Characters(text.Value);
            }
            if (node.Name.LocalName == "style")
            {
                var css = node.Value; var position = 0;
                while (position < css.Length)
                {
                    token.ThrowIfCancellationRequested(); var match = FontCss.Match(css, position);
                    if (!match.Success || match.Index != position || !match.Groups["family"].Value.StartsWith(fontPrefix, StringComparison.Ordinal)) Fail();
                    fonts.Add(match.Groups["family"].Value); position += match.Length;
                }
            }
        }
        foreach (var id in references)
            if (!ids.TryGetValue(id, out var target) || target.Name != Svg + "clipPath") Fail();
        foreach (var family in root.Descendants(Svg + "text").Attributes("font-family"))
            if (family.Value.StartsWith("ofd-preview-font-", StringComparison.Ordinal) && !fonts.Contains(family.Value)) Fail();
    }
    private static bool TryAddCompat(this Dictionary<string, XElement> ids, string id, XElement node)
    { if (ids.ContainsKey(id)) return false; ids.Add(id, node); return true; }
    private static void FinitePath(string value)
    { foreach (Match n in Regex.Matches(value, @"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?")) Number(n.Value); }
    private static int Numbers(string value)
    { var numbers = value.Split(new[] { ' ', ',', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries); foreach (var n in numbers) Number(n); return numbers.Length; }
    private static double Number(string value)
    { if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) || double.IsNaN(n) || double.IsInfinity(n)) { Fail(); return 0; } return n; }
    private static void Fail() => throw new NotSupportedException("Generated SVG is outside the static HTML preview policy.");
}
