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
            .ReadAsync(ofdInput, cancellationToken)
            .ConfigureAwait(false);
        var pages = package.Pages.OrderBy(page => page.Index).ToList();
        if (pageIndex < 0 || pageIndex >= pages.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(pageIndex));
        }

        var svg = BuildSvg(pages[pageIndex]);
        var bytes = Encoding.UTF8.GetBytes(
            new XDocument(new XDeclaration("1.0", "utf-8", null), svg)
                .ToString(SaveOptions.DisableFormatting));
        await svgOutput.WriteAsync(bytes, 0, bytes.Length, cancellationToken)
            .ConfigureAwait(false);
        await svgOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static XElement BuildSvg(OfdPage page)
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

        foreach (var element in EnumerateElements(page))
        {
            if (element is OfdTextElement text)
            {
                AddText(root, svgNs, page, text);
            }
            else if (element is OfdImageElement image && image.Data.Length > 0)
            {
                root.Add(new XElement(
                    svgNs + "image",
                    new XAttribute("x", Invariant(image.XMillimeters - page.XMillimeters)),
                    new XAttribute("y", Invariant(image.YMillimeters - page.YMillimeters)),
                    new XAttribute("width", Invariant(image.WidthMillimeters)),
                    new XAttribute("height", Invariant(image.HeightMillimeters)),
                    new XAttribute(
                        "href",
                        $"data:{image.MediaType};base64,{Convert.ToBase64String(image.Data)}")));
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

    private static void AddText(
        XElement root,
        XNamespace svgNs,
        OfdPage page,
        OfdTextElement text)
    {
        var transform = BuildTransform(
            text.XMillimeters - page.XMillimeters,
            text.YMillimeters - page.YMillimeters,
            text.Transform);
        if (text.Runs.Count == 0)
        {
            root.Add(new XElement(
                svgNs + "text",
                new XAttribute("x", "0"),
                new XAttribute("y", Invariant(text.FontSizeMillimeters)),
                new XAttribute("font-family", text.FontName),
                new XAttribute("font-size", Invariant(text.FontSizeMillimeters)),
                new XAttribute("fill", ToCssColor(text.FillColor)),
                new XAttribute("transform", transform),
                text.Text));
            return;
        }

        foreach (var run in text.Runs)
        {
            var node = new XElement(
                svgNs + "text",
                new XAttribute("x", Invariant(run.XMillimeters)),
                new XAttribute("y", Invariant(run.YMillimeters)),
                new XAttribute("font-family", text.FontName),
                new XAttribute("font-size", Invariant(text.FontSizeMillimeters)),
                new XAttribute("fill", ToCssColor(text.FillColor)),
                new XAttribute("transform", transform),
                run.Text);
            var deltaX = ExpandDeltas(run.DeltaX);
            var deltaY = ExpandDeltas(run.DeltaY);
            if (deltaX.Count > 0)
            {
                node.SetAttributeValue(
                    "dx",
                    string.Join(" ", ToSvgDeltas(run.Text, deltaX).Select(Invariant)));
            }

            if (deltaY.Count > 0)
            {
                node.SetAttributeValue(
                    "dy",
                    string.Join(" ", ToSvgDeltas(run.Text, deltaY).Select(Invariant)));
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

    private static IReadOnlyList<double> ExpandDeltas(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return Array.Empty<double>();
        }

        var tokens = value!.Split(
            new[] { ' ', '\t', '\r', '\n' },
            StringSplitOptions.RemoveEmptyEntries);
        var result = new List<double>();
        for (var index = 0; index < tokens.Length;)
        {
            if (string.Equals(tokens[index], "g", StringComparison.OrdinalIgnoreCase) &&
                index + 2 < tokens.Length &&
                int.TryParse(tokens[index + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var count) &&
                double.TryParse(tokens[index + 2], NumberStyles.Float, CultureInfo.InvariantCulture, out var repeated))
            {
                for (var repeat = 0; repeat < Math.Max(count, 0); repeat++)
                {
                    result.Add(repeated);
                }

                index += 3;
                continue;
            }

            if (double.TryParse(tokens[index], NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                result.Add(parsed);
            }

            index++;
        }

        return result;
    }

    private static IReadOnlyList<double> ToSvgDeltas(
        string text,
        IReadOnlyList<double> ofdDeltas)
    {
        var glyphCount = StringInfo.ParseCombiningCharacters(text).Length;
        if (glyphCount == 0)
        {
            return Array.Empty<double>();
        }

        var result = new List<double>(glyphCount) { 0 };
        for (var index = 0;
             index < ofdDeltas.Count && result.Count < glyphCount;
             index++)
        {
            result.Add(ofdDeltas[index]);
        }

        return result;
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
