using System;
using System.Globalization;
using System.Linq;
using System.Xml;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Text;

/// <summary>Consumes a completed query result for bounded legacy or typed original-GID OFD output.</summary>
public static class TextPrelayoutOfdWriter
{
    /// <summary>Creates one page from the exact result used for measurement. The caller must have font embedding rights.</summary>
    public static OfdDocumentPackage CreateSinglePage(TextPrelayoutResult result, double pageWidthMillimeters,
        double pageHeightMillimeters, double leftMillimeters, double topMillimeters)
        => CreateSinglePageCore(result, pageWidthMillimeters, pageHeightMillimeters,
            leftMillimeters, topMillimeters, positionedGlyphs: false);

    /// <summary>Creates a bounded single-TextCode page with actual original GIDs and a GID-preserving font subset.</summary>
    public static OfdDocumentPackage CreateSinglePageWithPositionedGlyphs(TextPrelayoutResult result,
        double pageWidthMillimeters, double pageHeightMillimeters, double leftMillimeters, double topMillimeters)
        => CreateSinglePageCore(result, pageWidthMillimeters, pageHeightMillimeters,
            leftMillimeters, topMillimeters, positionedGlyphs: true);

    private static OfdDocumentPackage CreateSinglePageCore(TextPrelayoutResult result, double pageWidthMillimeters,
        double pageHeightMillimeters, double leftMillimeters, double topMillimeters, bool positionedGlyphs)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        if (!Finite(pageWidthMillimeters) || !Finite(pageHeightMillimeters) ||
            !Finite(leftMillimeters) || !Finite(topMillimeters) ||
            pageWidthMillimeters <= 0 || pageHeightMillimeters <= 0 ||
            pageWidthMillimeters > 1_000_000 || pageHeightMillimeters > 1_000_000 ||
            leftMillimeters < 0 || topMillimeters < 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid destination page geometry.");
        if (result.Overflows || leftMillimeters + result.WidthMillimeters > pageWidthMillimeters ||
            topMillimeters + result.HeightMillimeters > pageHeightMillimeters)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "Prelayout result exceeds its destination page.");
        foreach (var line in result.Lines)
        {
            var ink = line.InkBoundsMillimeters;
            if (ink is null) continue;
            if (leftMillimeters + ink.X < -1e-8 || topMillimeters + ink.Y < -1e-8 ||
                leftMillimeters + ink.X + ink.Width > pageWidthMillimeters + 1e-8 ||
                topMillimeters + ink.Y + ink.Height > pageHeightMillimeters + 1e-8)
                throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                    "Measured glyph ink would extend outside the destination page.");
        }
        if (result.Request.Script != "Latn" || result.Face.Bold || result.Face.Italic)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "Positioned OFD output requires the bounded regular, horizontal Latin profile.");
        try { XmlConvert.VerifyXmlChars(result.Face.Name); }
        catch (XmlException exception)
        {
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The selected font name cannot be written as OFD XML.", exception);
        }
        var scale = result.Request.FontSizeMillimeters / result.Face.UnitsPerEm;
        // Complete preflight before creating even an in-memory package.
        foreach (var line in result.Lines)
        {
            var text = result.Request.Text.Substring(line.SourceStart, line.SourceLength);
            if (line.ShapedRun.Request.Text != text ||
                line.ShapedRun.Request.SourceStart != line.SourceStart ||
                !ReferenceEquals(line.ShapedRun.Request.Face, result.Face))
                throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "The shaped line has no lossless Text/Delta representation.");
            if (positionedGlyphs)
            {
                if (text.Length == 0) continue;
                var latinCombiningPair = OfdPositionedGlyphProfile.IsLatinCombiningPair(text);
                if (text.Length > 8 || line.ShapedRun.Glyphs.Count is < 1 or > 8 ||
                    (!latinCombiningPair && text.Any(value => value < 0x21 || value > 0x7e)))
                    throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                        "The positioned output requires a short ASCII line or one bounded Latin combining pair.");
                var glyphs = line.ShapedRun.Glyphs;
                if (latinCombiningPair && glyphs.Count != 1)
                    throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                        "The Latin combining profile requires one complete shaped glyph cluster.");
                if (glyphs.Count != text.Length && !(glyphs.Count == 1 && text.Length is 2 or 3))
                    throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                        "The positioned output needs one-to-one glyphs or a complete short ligature.");
                if (glyphs.Count == text.Length)
                {
                    var sourceByGid = new System.Collections.Generic.Dictionary<uint, char>();
                    var gidBySource = new System.Collections.Generic.Dictionary<char, uint>();
                    for (var index = 0; index < glyphs.Count; index++)
                    {
                        if (sourceByGid.TryGetValue(glyphs[index].GlyphId, out var prior) && prior != text[index])
                            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                                "One shaped GID cannot represent distinct source characters in PDF/SVG output.");
                        if (gidBySource.TryGetValue(text[index], out var priorGid) && priorGid != glyphs[index].GlyphId)
                            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                                "One source character cannot use distinct GIDs in the bounded PDF font alias.");
                        sourceByGid[glyphs[index].GlyphId] = text[index];
                        gidBySource[text[index]] = glyphs[index].GlyphId;
                    }
                }
                for (var i = 0; i < glyphs.Count; i++)
                    if (glyphs[i].YAdvance != 0 ||
                        (glyphs.Count == text.Length
                            ? glyphs[i].ClusterStart != line.SourceStart + i || glyphs[i].ClusterLength != 1
                            : glyphs[i].ClusterStart != line.SourceStart || glyphs[i].ClusterLength != text.Length))
                        throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                            "The positioned glyph clusters exceed the verified single-TextCode profile.");
            }
            else
            {
                if (line.ShapedRun.Glyphs.Count != text.Length)
                    throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                        "The shaped line has no lossless Text/Delta representation.");
                for (var i = 0; i < text.Length; i++)
                {
                    var glyph = line.ShapedRun.Glyphs[i];
                    if (text[i] < 0x20 || text[i] > 0x7e || glyph.ClusterStart != line.SourceStart + i ||
                        glyph.ClusterLength != 1 || glyph.GlyphId != result.Face.NominalGlyph(text[i]) ||
                        glyph.XOffset != 0 || glyph.YOffset != 0 || glyph.YAdvance != 0)
                        throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                            "A shaped glyph needs typed GID/offset output, reserved for the later output stage.");
                }
            }
        }
        var package = new OfdDocumentPackage();
        package.Options.DefaultPageWidthMillimeters = pageWidthMillimeters;
        package.Options.DefaultPageHeightMillimeters = pageHeightMillimeters;
        // The typed path preserves the measured face and GIDs for subset closure;
        // the legacy path continues to embed the complete selected face.
        if (!positionedGlyphs) package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        package.Fonts.Add(new OfdFontResource { Id = "F1", FontName = result.Face.Name,
            FamilyName = result.Face.Name, Charset = "unicode", Data = result.Face.GetFaceBytes(), CollectionFaceIndex = 0 });
        var page = new OfdPage { Id = "1", Index = 0, WidthMillimeters = pageWidthMillimeters,
            HeightMillimeters = pageHeightMillimeters };
        foreach (var line in result.Lines.Where(value => value.SourceLength > 0))
        {
            var text = result.Request.Text.Substring(line.SourceStart, line.SourceLength);
            var ink = line.InkBoundsMillimeters;
            var inkTop = ink?.Y - line.TopMillimeters ?? 0;
            var boundaryLeft = Math.Min(0, ink?.X ?? 0);
            var boundaryTop = Math.Min(0, inkTop);
            var boundaryRight = Math.Max(line.AdvanceMillimeters, ink is null ? 0 : ink.X + ink.Width);
            var boundaryBottom = Math.Max(line.HeightMillimeters, ink is null ? 0 : inkTop + ink.Height);
            var element = new OfdTextElement
            {
                LayerType = "Body", XMillimeters = leftMillimeters + boundaryLeft,
                YMillimeters = topMillimeters + line.TopMillimeters + boundaryTop,
                WidthMillimeters = Math.Max(0.1, boundaryRight - boundaryLeft),
                HeightMillimeters = Math.Max(0.1, boundaryBottom - boundaryTop),
                FontName = result.Face.Name, FontResourceId = "F1",
                FontSizeMillimeters = result.Request.FontSizeMillimeters, Text = text,
            };
            var outputRun = new OfdTextRun
            {
                Text = text, XMillimeters = -boundaryLeft,
                YMillimeters = line.BaselineMillimeters - line.TopMillimeters - boundaryTop,
                DeltaX = text.Length > 1 ? string.Join(" ", line.ShapedRun.Glyphs.Take(text.Length - 1)
                    .Select(glyph => (glyph.XAdvance * scale).ToString("0.###############", CultureInfo.InvariantCulture))) : null,
            };
            if (positionedGlyphs)
            {
                var glyphs = new OfdPositionedGlyph[line.ShapedRun.Glyphs.Count];
                long penX = 0, penY = 0;
                for (var index = 0; index < glyphs.Length; index++)
                {
                    var glyph = line.ShapedRun.Glyphs[index];
                    var projection = TextGlyphProjection.Create(glyph, penX, penY, scale,
                        -boundaryLeft, line.BaselineMillimeters - line.TopMillimeters - boundaryTop);
                    glyphs[index] = new OfdPositionedGlyph(glyph.GlyphId, glyph.ClusterStart, glyph.ClusterLength,
                        projection.X, projection.Y, projection.AdvanceX,
                        // Preflight admits only YAdvance == 0; retain the legacy positive zero.
                        glyph.YAdvance * scale);
                    penX = projection.NextPenX; penY = projection.NextPenY;
                }
                outputRun.XMillimeters = glyphs[0].OriginXMillimeters;
                outputRun.YMillimeters = glyphs[0].OriginYMillimeters;
                outputRun.DeltaX = glyphs.Length > 1 ? string.Join(" ", glyphs.Skip(1).Select((glyph, index) =>
                    (glyph.OriginXMillimeters - glyphs[index].OriginXMillimeters)
                    .ToString("0.###############", CultureInfo.InvariantCulture))) : null;
                outputRun.DeltaY = Enumerable.Range(1, glyphs.Length - 1).Any(index =>
                    glyphs[index].OriginYMillimeters != glyphs[index - 1].OriginYMillimeters)
                    ? string.Join(" ", glyphs.Skip(1).Select((glyph, index) =>
                        (glyph.OriginYMillimeters - glyphs[index].OriginYMillimeters)
                        .ToString("0.###############", CultureInfo.InvariantCulture))) : null;
                outputRun.PositionedGlyphs = new OfdPositionedGlyphRun(text, line.SourceStart,
                    result.Face.FaceSha256, glyphs);
            }
            element.Runs.Add(outputRun);
            page.Elements.Add(element);
        }
        package.Pages.Add(page);
        return package;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
