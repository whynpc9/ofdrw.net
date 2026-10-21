using System;
using System.Globalization;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Text;

/// <summary>Consumes a completed query result for the deliberately narrow, one-glyph-per-ASCII-character OFD profile.</summary>
public static class TextPrelayoutOfdWriter
{
    /// <summary>Creates one page from the exact result used for measurement. The caller must have font embedding rights.</summary>
    public static OfdDocumentPackage CreateSinglePage(TextPrelayoutResult result, double pageWidthMillimeters,
        double pageHeightMillimeters, double leftMillimeters, double topMillimeters)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        if (!Finite(pageWidthMillimeters) || !Finite(pageHeightMillimeters) ||
            !Finite(leftMillimeters) || !Finite(topMillimeters) ||
            pageWidthMillimeters <= 0 || pageHeightMillimeters <= 0 ||
            pageWidthMillimeters > 1_000_000 || pageHeightMillimeters > 1_000_000 ||
            leftMillimeters < 0 || topMillimeters < 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid destination page geometry.");
        if (result.Overflows || leftMillimeters + result.Request.WidthMillimeters > pageWidthMillimeters ||
            topMillimeters + result.HeightMillimeters > pageHeightMillimeters)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "Prelayout result exceeds its destination page.");
        if (result.Request.Script != "Latn" || result.Face.Bold || result.Face.Italic)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The first OFD output profile is regular, horizontal Latin ASCII only.");
        var scale = result.Request.FontSizeMillimeters / result.Face.UnitsPerEm;
        // Complete preflight before creating even an in-memory package. The typed OFD model
        // cannot yet carry arbitrary GIDs, clusters, offsets, or bidi semantics.
        foreach (var line in result.Lines)
        {
            var text = result.Request.Text.Substring(line.SourceStart, line.SourceLength);
            if (line.ShapedRun.Glyphs.Count != text.Length || line.ShapedRun.Request.Text != text ||
                !ReferenceEquals(line.ShapedRun.Request.Face, result.Face))
                throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "The shaped line has no lossless Text/Delta representation.");
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
        var package = new OfdDocumentPackage();
        package.Options.DefaultPageWidthMillimeters = pageWidthMillimeters;
        package.Options.DefaultPageHeightMillimeters = pageHeightMillimeters;
        // Preserves the exact face bytes and GIDs that were measured. A later typed-GID
        // writer can add a verified subset path without changing this result contract.
        package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        package.Fonts.Add(new OfdFontResource { Id = "F1", FontName = result.Face.Name,
            FamilyName = result.Face.Name, Charset = "unicode", Data = result.Face.GetFaceBytes(), CollectionFaceIndex = 0 });
        var page = new OfdPage { Id = "1", Index = 0, WidthMillimeters = pageWidthMillimeters,
            HeightMillimeters = pageHeightMillimeters };
        foreach (var line in result.Lines.Where(value => value.SourceLength > 0))
        {
            var text = result.Request.Text.Substring(line.SourceStart, line.SourceLength);
            var element = new OfdTextElement
            {
                LayerType = "Body", XMillimeters = leftMillimeters,
                YMillimeters = topMillimeters + line.TopMillimeters,
                WidthMillimeters = Math.Max(0.1, line.AdvanceMillimeters),
                HeightMillimeters = line.HeightMillimeters,
                FontName = result.Face.Name, FontResourceId = "F1",
                FontSizeMillimeters = result.Request.FontSizeMillimeters, Text = text,
            };
            element.Runs.Add(new OfdTextRun
            {
                Text = text, YMillimeters = line.BaselineMillimeters - line.TopMillimeters,
                DeltaX = text.Length > 1 ? string.Join(" ", line.ShapedRun.Glyphs.Take(text.Length - 1)
                    .Select(glyph => (glyph.XAdvance * scale).ToString("0.###############", CultureInfo.InvariantCulture))) : null,
            });
            page.Elements.Add(element);
        }
        package.Pages.Add(page);
        return package;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
