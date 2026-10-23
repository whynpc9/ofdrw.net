using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// A transient rendering view. Raw Reader runs and owned editing identities stay untouched.
internal static class OfdTextRenderProjection
{
    internal static OfdTextElement Create(OfdTextElement text, OfdModelMaterialization material, CancellationToken token)
        => CreateCore(text, material, token, renderGeometry: true);

    internal static OfdTextElement CreateLogical(OfdTextElement text, OfdModelMaterialization material, CancellationToken token)
        => CreateCore(text, material, token, renderGeometry: false);

    private static OfdTextElement CreateCore(OfdTextElement text, OfdModelMaterialization material, CancellationToken token, bool renderGeometry)
    {
        var source = ReadSource(text, material, token);
        if (source is null) return text;
        var codes = source.Value.Codes; var raw = source.Value.Raw;
        var view = (OfdTextElement)OfdModelCloner.CloneElement(text, clonePayloads: false, material: material, cancellationToken: token);
        view.Runs.Clear(); var joined = new StringBuilder(); double? previousX = null, previousY = null;
        for (var i = 0; i < codes.Length; i++)
        {
            token.ThrowIfCancellationRequested(); var code = codes[i];
            if (renderGeometry && code.Elements().Any()) throw new NotSupportedException("Nested source TextCode content cannot be rendered safely.");
            var x = renderGeometry ? Axis(code, "X", previousX) : text.Runs[i].XMillimeters;
            var y = renderGeometry ? Axis(code, "Y", previousY) : text.Runs[i].YMillimeters;
            previousX = x; previousY = y; // Inherit the previous START, never its final glyph.
            var decoded = Decode(raw[i], token); joined.Append(decoded);
            view.Runs.Add(new OfdTextRun { Text = decoded, XMillimeters = x, YMillimeters = y,
                DeltaX = code.Attribute("DeltaX")?.Value, DeltaY = code.Attribute("DeltaY")?.Value });
        }
        view.Text = joined.ToString(); return view;
    }

    // Logical extraction shares decoding and source/model consistency, without adopting
    // rendering restrictions on unknown children or requiring usable glyph coordinates.
    internal static string LogicalText(OfdTextElement text, OfdModelMaterialization material, CancellationToken token)
    {
        var source = ReadSource(text, material, token);
        if (source is null) return text.Text;
        var joined = new StringBuilder();
        foreach (var raw in source.Value.Raw) { token.ThrowIfCancellationRequested(); joined.Append(Decode(raw, token)); }
        return joined.ToString();
    }

    private static (XElement[] Codes, string[] Raw)? ReadSource(OfdTextElement text, OfdModelMaterialization material, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (text.SourceXml is not { Length: > 0 } source) return null;
        material.GuardFragmentXml(source, token);
        material.ChargeCopy(checked(16384L + 16L * source.Length + 256L * text.Runs.Count), token);
        material.ChargeWork(checked(8L * source.Length), token);
        var node = OfdXmlEnvelope.Parse(source);
        if (node.Name.LocalName != "TextObject") throw new InvalidDataException("Text source must be a TextObject.");
        var codes = node.Elements(node.Name.Namespace + "TextCode").ToArray();
        // Only source text/position fields must still match the Reader projection.
        // Unrelated typed style/transform changes remain usable by the renderer.
        var raw = codes.Select(code => string.Concat(code.Nodes().OfType<XText>().Select(value => value.Value))).ToArray();
        var originalSize = RawNumber(node.Attribute("Size")?.Value, 4);
        var sameRuns = codes.Length == text.Runs.Count;
        double? previousX = null, previousY = null;
        for (var i = 0; sameRuns && i < codes.Length; i++)
        {
            token.ThrowIfCancellationRequested(); var run = text.Runs[i]; var code = codes[i];
            var x = RawNumber(code.Attribute("X")?.Value, previousX ?? 0);
            var y = RawNumber(code.Attribute("Y")?.Value, previousY ?? originalSize);
            sameRuns = run.Text == raw[i] &&
                SameSourceAxis(run.XMillimeters, text.SourceInheritedTextCodeAxes ? x : RawNumber(code.Attribute("X")?.Value, 0)) &&
                SameSourceAxis(run.YMillimeters, text.SourceInheritedTextCodeAxes ? y : RawNumber(code.Attribute("Y")?.Value, originalSize)) &&
                run.DeltaX == code.Attribute("DeltaX")?.Value && run.DeltaY == code.Attribute("DeltaY")?.Value;
            previousX = x; previousY = y;
        }
        if (!sameRuns) return null; // Explicit typed run edits take precedence over the old envelope.
        if (text.Text != string.Concat(raw))
            throw new InvalidDataException("Text and source runs diverged; update the typed runs or clear the source envelope before rendering.");
        return (codes, raw);
    }

    internal static string EncodeLiteral(string text, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeWork(checked(2L * text.Length), token);
        var slashes = 0;
        for (var index = 0; index < text.Length; index++)
        { if ((index & 255) == 0) token.ThrowIfCancellationRequested(); if (text[index] == '\\') slashes++; }
        if (slashes == 0) return text;
        var length = checked(text.Length + 4L * slashes);
        if (length > int.MaxValue) throw new InvalidDataException("Encoded TextCode exceeds string capacity.");
        // Reserve both the builder buffer and final UTF-16 string before either allocation.
        material.ChargeCopy(checked(128L + 4L * length), token);
        material.ChargeWork(checked(2L * length), token);
        var encoded = new StringBuilder((int)length);
        for (var index = 0; index < text.Length; index++)
        {
            if ((index & 255) == 0) token.ThrowIfCancellationRequested();
            if (text[index] == '\\') encoded.Append("\\005C"); else encoded.Append(text[index]);
        }
        token.ThrowIfCancellationRequested(); return encoded.ToString();
    }

    private static double RawNumber(string? value, double fallback) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : fallback;

    // Reader can retain NaN from an invalid source axis. It is still the same
    // source run for logical decoding; render admission rejects the axis later.
    private static bool SameSourceAxis(double actual, double source) =>
        actual == source || double.IsNaN(actual) && double.IsNaN(source);

    private static double Axis(XElement code, string name, double? previous)
    {
        if (code.Attribute(name) is not { } axis)
            return previous ?? throw new InvalidDataException("First source TextCode requires " + name + ".");
        if (!double.TryParse(axis.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) || double.IsNaN(value) || double.IsInfinity(value))
            throw new InvalidDataException("Source TextCode requires a finite " + name + ".");
        return value;
    }

    private static string Decode(string raw, CancellationToken token)
    {
        var decoded = new StringBuilder(raw.Length);
        for (var index = 0; index < raw.Length; index++)
        {
            if ((index & 255) == 0) token.ThrowIfCancellationRequested();
            var value = raw[index];
            if (value == '\\')
            {
                if (raw.Length - index <= 4 || !ushort.TryParse(raw.Substring(index + 1, 4), NumberStyles.AllowHexSpecifier,
                        CultureInfo.InvariantCulture, out var number))
                    throw new InvalidDataException("Unsupported source TextCode escape; four hexadecimal digits are required.");
                value = (char)number; index += 4;
                if (value < 0x20) throw new InvalidDataException("Escaped source control characters cannot be rendered safely.");
            }
            decoded.Append(value);
        }
        var result = decoded.ToString();
        for (var index = 0; index < result.Length; index++)
        {
            if ((index & 255) == 0) token.ThrowIfCancellationRequested();
            if (char.IsHighSurrogate(result[index]))
            { if (++index >= result.Length || !char.IsLowSurrogate(result[index])) throw new InvalidDataException("Unpaired source TextCode surrogate."); }
            else if (char.IsLowSurrogate(result[index])) throw new InvalidDataException("Unpaired source TextCode surrogate.");
        }
        return result;
    }
}
