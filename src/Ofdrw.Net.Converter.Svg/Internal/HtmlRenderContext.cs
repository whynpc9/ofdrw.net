using System;
using System.IO;
using System.Collections.Generic;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.IO;

using System.Linq;
using System.Threading;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Svg.Internal;

// Reservations precede the renderer's projection, normalization, glyph splitting, DOM,
// base64 and CSS allocations. They deliberately overestimate, rather than measuring a
// large allocation afterwards. Reader/composite materialization has independent limits.
internal sealed class HtmlRenderContext
{
    private static readonly Regex PathToken = new(@"\G[\s,]*(?:([MLBQACmlbqac])|([-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?))", RegexOptions.CultureInvariant);
    private readonly OfdHtmlOptions _options;
    private readonly CancellationToken _token;
    private long _nodes, _characters, _resources, _pixels;
    private readonly Dictionary<OfdFontResource, (byte[] Data, string Hash, OpenTypeCmap? Coverage, bool Bold, bool Italic)> _fontCache = new();
    private readonly Dictionary<string, (byte[] Data, string Hash, OpenTypeCmap? Coverage, bool Bold, bool Italic)> _faceCache = new(StringComparer.Ordinal);
    internal (byte[] Data, string Hash, OpenTypeCmap? Coverage, bool Bold, bool Italic) Font(OfdFontResource font, FontParseBudget budget)
    {
        if (_fontCache.TryGetValue(font, out var cached)) return cached;
        Reserve(0, 512);
        var key = BinaryIdentity.Hash(font.Data) + ":" + font.CollectionFaceIndex.ToString(CultureInfo.InvariantCulture);
        if (_faceCache.TryGetValue(key, out cached)) { _fontCache.Add(font, cached); return cached; }
        Reserve(0, checked(4096L + 2L * font.Data.LongLength));
        var data = OpenTypeCollection.SelectFace(font.Data, font.CollectionFaceIndex);
        var face = new OpenTypeFace(data); OpenTypeCmap? coverage = null;
        try { coverage = new OpenTypeCmap(face, budget); } catch (NotSupportedException) { }
        var style = face.Style;
        var value = (data, BinaryIdentity.Hash(data), coverage, style.Bold, style.Italic);
        _fontCache.Add(font, value); _faceCache.Add(key, value); return value;
    }
    internal string Prefix { get; private set; } = "";
    internal string FontPrefix { get; private set; } = "";
    internal HtmlRenderContext(OfdHtmlOptions options, CancellationToken token)
    { _options = options; _token = token; Reserve(32, checked(8192L + options.Title.Length * 24L)); }
    internal void Occurrence(int index)
    { Prefix = "ofd-preview-svg-" + index + "-"; FontPrefix = "ofd-preview-font-" + index + "-"; }
    internal void StartPage(OfdPage page)
    {
        Finite(page.XMillimeters); Finite(page.YMillimeters);
        if (!Positive(page.WidthMillimeters) || !Positive(page.HeightMillimeters))
            throw new InvalidDataException("HTML preview requires finite positive page dimensions.");
        Reserve(32, 2048);
    }
    internal void BeforeFontEmission(byte[] data)
    {
        Resource(data.LongLength);
        // Includes collection face extraction, base64, builders, style text, validation
        // and UTF-8 escaping. Collection source is the safe upper bound of a face.
        Reserve(4, checked(4096L + 8L * data.LongLength));
    }
    internal void BeforeElement(OfdElement element)
    {
        _token.ThrowIfCancellationRequested();
        Finite(element.XMillimeters); Finite(element.YMillimeters);
        Finite(element.WidthMillimeters); Finite(element.HeightMillimeters);
        string? source = element switch
        { OfdTextElement t => t.SourceXml, OfdPathElement p => p.SourceXml, OfdImageElement i => i.SourceXml, _ => null };
        // Raw visible drawing is never silently omitted by this new API.
        if (element is not (OfdTextElement or OfdPathElement or OfdImageElement))
            throw new NotSupportedException("HTML preview cannot render an opaque drawing element.");
        var sourceLength = source?.Length ?? 0;
        long textLength = 0, runs = 0, pathLength = 0, clipLength = element.ClippingXml?.Length ?? 0;
        if (element is OfdTextElement text)
        {
            textLength = text.Text.Length; runs = text.Runs.Count;
            foreach (var run in text.Runs)
            { _token.ThrowIfCancellationRequested(); textLength = checked(textLength + run.Text.Length + (run.DeltaX?.Length ?? 0) + (run.DeltaY?.Length ?? 0)); }
            Finite(text.FontSizeMillimeters); Matrix(text.Transform);
            Characters(text.Text); Characters(text.FontName);
        }
        if (element is OfdPathElement path)
        { pathLength = path.AbbreviatedData.Length; Matrix(path.Transform); Finite(path.LineWidthMillimeters); }
        if (element is OfdImageElement image)
        { Matrix(image.Transform); Resource(image.Data.LongLength); }
        if (element.Clips is { } clips)
            foreach (var clip in clips.Items)
                foreach (var area in clip.Areas)
                    clipLength = checked(clipLength + area.Path.AbbreviatedData.Length + 128);
        // At most one tspan per UTF-16 unit; each has <= 4 attributes plus text.
        // Clip XML characters conservatively bound paths/groups/attributes as well.
        Reserve(checked(96L + 8L * textLength + 24L * runs + 8L * clipLength),
            checked(8192L + 384L * textLength + 4096L * runs + 64L * (sourceLength + pathLength + clipLength) +
                (element is OfdImageElement img ? 16L * img.Data.LongLength : 0)));
        if (source is { Length: > 0 })
        {
            var node = OfdXmlEnvelope.Parse(source);
            if (!OfdGraphicXmlContract.HasKnownChildren(node) || OfdGraphicXmlContract.HasUnsupportedReferences(node))
                throw new NotSupportedException("HTML preview does not support unknown drawing children or references.");
            if (element is OfdTextElement &&
                (((OfdTextElement)element).Weight < 600 && node.Attribute("Stroke")?.Value is "true" or "1" || node.Attribute("Fill")?.Value is "false" or "0" ||
                 node.Attribute("HScale") is { Value: not "1" } || node.Attribute("ReadDirection") is { Value: not "0" } ||
                 node.Attribute("CharDirection") is { Value: not "0" } || node.Attribute("Alpha") is { Value: not "255" }))
                throw new NotSupportedException("Text stroke, scale, direction or object alpha is outside the HTML render profile.");
            if (node.DescendantsAndSelf().Attributes().Any(a => !OfdGraphicXmlContract.IsKnownAttribute(a) &&
                !(a.Name == XNamespace.Xml + "space" && a.Parent!.Name.LocalName == "TextCode" && a.Value == "preserve")))
                throw new NotSupportedException("HTML preview does not support unknown drawing attributes.");
        }
        if (element is OfdPathElement drawingPath) ValidatePath(drawingPath.AbbreviatedData);
        if (element.Clips is { } typedClips)
            foreach (var clip in typedClips.Items) foreach (var area in clip.Areas) ValidatePath(area.Path.AbbreviatedData);
        if (element.ClippingXml is { Length: > 0 } clipXml)
            foreach (var clipPath in OfdXmlEnvelope.Parse(clipXml).Descendants().Where(n => n.Name.LocalName == "AbbreviatedData")) ValidatePath(clipPath.Value);
        if (element is OfdImageElement media) HtmlImageAdmission.Validate(media.Data, media.MediaType, _token, pixels =>
        { _token.ThrowIfCancellationRequested(); Spend(ref _pixels, pixels, _options.MaxDecodedImagePixels, "decoded image pixels"); });
    }
    private void ValidatePath(string data)
    {
        var position = 0; var arity = -1; var numbers = 0;
        while (position < data.Length)
        {
            _token.ThrowIfCancellationRequested(); var match = PathToken.Match(data, position);
            if (!match.Success) { if (string.IsNullOrWhiteSpace(data.Substring(position))) break; throw new NotSupportedException("Unsupported OFD path syntax."); }
            position += match.Length;
            if (match.Groups[1].Success)
            {
                if (arity > 0 && (numbers < arity || numbers % arity != 0)) throw new InvalidDataException("Incomplete OFD path command.");
                var command = char.ToUpperInvariant(match.Groups[1].Value[0]);
                if (arity == -1 && command != 'M') throw new InvalidDataException("OFD path must begin with Move.");
                arity = command switch { 'M' or 'L' => 2, 'B' => 6, 'Q' => 4, 'A' => 7, 'C' => 0, _ => -1 }; numbers = 0;
            }
            else
            {
                if (arity <= 0 || !double.TryParse(match.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                    throw new InvalidDataException("Invalid OFD path coordinate.");
                Finite(number); numbers++;
            }
        }
        if (arity == -1 || arity > 0 && (numbers < arity || numbers % arity != 0)) throw new InvalidDataException("Incomplete OFD path.");
    }
    internal void Reserve(long nodes, long characters)
    { _token.ThrowIfCancellationRequested(); Spend(ref _nodes, nodes, _options.MaxGeneratedNodes, "nodes"); Spend(ref _characters, characters, _options.MaxGeneratedCharacters, "characters"); }
    private void Resource(long bytes) => Spend(ref _resources, bytes, _options.MaxEmbeddedResourceBytes, "embedded resources");
    private static void Spend(ref long used, long value, long limit, string name)
    { if (value < 0 || value > limit - used) throw new InvalidDataException("HTML preview capacity exceeded: " + name + "."); used += value; }
    internal static bool Positive(double value) => value > 0 && !double.IsInfinity(value) && !double.IsNaN(value);
    private static void Finite(double value)
    { if (double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidDataException("Non-finite HTML drawing coordinate."); }
    private static void Matrix(double[]? matrix)
    { if (matrix is null) return; if (matrix.Length != 6) throw new InvalidDataException("Invalid drawing matrix."); foreach (var n in matrix) Finite(n); }
    internal static void Characters(string value)
    {
        try { XmlConvert.VerifyXmlChars(value); }
        catch (XmlException ex) { throw new InvalidDataException("Invalid preview characters.", ex); }
    }
}
