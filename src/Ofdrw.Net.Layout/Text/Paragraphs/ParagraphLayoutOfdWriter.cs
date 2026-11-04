using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Text.Paragraphs;

/// <summary>Writes a final <see cref="ParagraphLayoutResult"/> to OFD pages without reshaping or re-measuring.</summary>
/// <remarks>
/// <para>Each run becomes one TextObject whose TextCode keeps the run's unmodified logical text.
/// Every source cluster has its own CGTransform (logical order) carrying the original GIDs, and every
/// glyph slot has an explicit origin, so RTL and many-to-many clusters need no reader inference.</para>
/// <para>Output profile: BMP text; fill paint in DeviceRGB with alpha; no stroke; faces must permit
/// embedding (OS/2.fsType). Fonts are embedded as GID-preserving subsets by default. Anything else is
/// rejected with <see cref="TextLayoutError.UnsupportedOutput"/> before a package is created.</para>
/// </remarks>
public static class ParagraphLayoutOfdWriter
{
    /// <summary>Creates one OFD page per layout page. The content box starts at (left, top) on each page.</summary>
    public static OfdDocumentPackage CreateDocument(ParagraphLayoutResult result, double pageWidthMillimeters,
        double pageHeightMillimeters, double leftMillimeters, double topMillimeters,
        bool embedFullFonts = false, CancellationToken cancellationToken = default)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        if (!Finite(pageWidthMillimeters) || !Finite(pageHeightMillimeters) || !Finite(leftMillimeters) ||
            !Finite(topMillimeters) || pageWidthMillimeters <= 0 || pageHeightMillimeters <= 0 ||
            pageWidthMillimeters > 1_000_000 || pageHeightMillimeters > 1_000_000 || leftMillimeters < 0 || topMillimeters < 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid destination page geometry.");
        if (leftMillimeters + result.WidthMillimeters > pageWidthMillimeters + 1e-9 ||
            result.Request.Options.PageHeightMillimeters is { } contentHeight &&
            topMillimeters + contentHeight > pageHeightMillimeters + 1e-9)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "The layout frame exceeds the destination page.");

        // Complete preflight before creating any package object.
        var faces = new List<TextFontFace>();
        foreach (var line in result.Lines)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (topMillimeters + line.TopMillimeters + line.HeightMillimeters > pageHeightMillimeters + 1e-9)
                throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "A line box leaves the destination page.",
                    sourceStart: line.SourceStart, sourceLength: line.SourceLength);
            foreach (var run in line.Runs)
            {
                Preflight(result, run);
                if (run.InkBoundsMillimeters is { } ink && (leftMillimeters + ink.X < -1e-9 ||
                        topMillimeters + ink.Y < -1e-9 || leftMillimeters + ink.X + ink.Width > pageWidthMillimeters + 1e-9 ||
                        topMillimeters + ink.Y + ink.Height > pageHeightMillimeters + 1e-9))
                    throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "Glyph ink would leave the destination page.",
                        sourceStart: run.SourceStart, sourceLength: run.SourceLength);
                if (!faces.Any(face => face.FaceSha256 == run.Face.FaceSha256)) faces.Add(run.Face);
            }
        }

        var package = new OfdDocumentPackage();
        package.Options.DefaultPageWidthMillimeters = pageWidthMillimeters;
        package.Options.DefaultPageHeightMillimeters = pageHeightMillimeters;
        if (embedFullFonts) package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        var resources = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 0; index < faces.Count; index++)
        {
            var face = faces[index];
            var id = "F" + (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            resources[face.FaceSha256] = id;
            package.Fonts.Add(new OfdFontResource
            {
                Id = id, FontName = face.Name, FamilyName = face.Name, Charset = "unicode",
                Bold = face.Bold, Italic = face.Italic, Data = face.GetFaceBytes(), CollectionFaceIndex = 0,
            });
        }
        foreach (var layoutPage in result.Pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = new OfdPage
            {
                Id = (layoutPage.Index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                Index = layoutPage.Index, WidthMillimeters = pageWidthMillimeters, HeightMillimeters = pageHeightMillimeters,
            };
            foreach (var line in layoutPage.Lines)
                foreach (var run in line.Runs)
                    page.Elements.Add(CreateElement(run, line, leftMillimeters, topMillimeters, resources[run.Face.FaceSha256]));
            package.Pages.Add(page);
        }
        return package;
    }

    private static void Preflight(ParagraphLayoutResult result, ParagraphLayoutRun run)
    {
        var paint = run.Style.Paint;
        if (!paint.Fill || paint.Stroke)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The OFD text writer profile supports filled text without stroke.",
                sourceStart: run.SourceStart, sourceLength: run.SourceLength);
        if (run.Style.Synthetic != TextSyntheticStyle.None || run.Style.RequestedBold != run.Style.FaceBold ||
            run.Style.RequestedItalic != run.Style.FaceItalic)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "Synthetic styles are not written.",
                sourceStart: run.SourceStart, sourceLength: run.SourceLength);
        if (run.Text.Any(char.IsSurrogate))
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "OFD CodePosition counting for supplementary characters needs a separately verified profile.",
                sourceStart: run.SourceStart, sourceLength: run.SourceLength);
        if (run.Text.Length > 4096 || run.Glyphs.Count > 16384)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "A run exceeds the OFD TextCode budget.",
                sourceStart: run.SourceStart, sourceLength: run.SourceLength);
        if (result.Text.Substring(run.SourceStart, run.SourceLength) != run.Text ||
            run.ShapedRun.Request.Text != run.Text || run.ShapedRun.Request.SourceStart != run.SourceStart ||
            run.ShapedRun.Glyphs.Count != run.Glyphs.Count ||
            run.Glyphs.Where((glyph, index) => glyph.GlyphId != run.ShapedRun.Glyphs[index].GlyphId ||
                glyph.ClusterStart != run.ShapedRun.Glyphs[index].ClusterStart).Any())
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The run lost its source or shaped-glyph identity.", sourceStart: run.SourceStart, sourceLength: run.SourceLength);
        if (run.Glyphs.Any(glyph => glyph.GlyphId > ushort.MaxValue))
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "A GID exceeds the OFD glyph range.",
                sourceStart: run.SourceStart, sourceLength: run.SourceLength);
        try { XmlConvert.VerifyXmlChars(run.Face.Name); XmlConvert.VerifyXmlChars(run.Text); }
        catch (XmlException error)
        {
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "The run cannot be written as OFD XML.", error,
                run.SourceStart, run.SourceLength);
        }
        // Clusters must partition the run and keep their glyphs contiguous (validated again by the model).
        try
        {
            _ = new OfdPositionedGlyphRun(run.Text, run.SourceStart, run.Face.FaceSha256,
                run.Glyphs.Select(glyph => new OfdPositionedGlyph(glyph.GlyphId, glyph.ClusterStart, glyph.ClusterLength,
                    0, 0, 0, 0)).ToArray(), OfdPositionedGlyphMapping.ClusterMap);
        }
        catch (ArgumentException error)
        {
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The run's glyph clusters have no complete OFD cluster mapping.", error, run.SourceStart, run.SourceLength);
        }
    }

    private static OfdTextElement CreateElement(ParagraphLayoutRun run, ParagraphLayoutLine line,
        double left, double top, string fontId)
    {
        // Boundary: the run's advance box over the line box, extended by any ink.
        var minX = run.OriginXMillimeters;
        var maxX = run.OriginXMillimeters + Math.Max(run.AdvanceMillimeters, 0);
        var minY = line.TopMillimeters;
        var maxY = line.TopMillimeters + line.HeightMillimeters;
        if (run.InkBoundsMillimeters is { } ink)
        {
            minX = Math.Min(minX, ink.X); maxX = Math.Max(maxX, ink.X + ink.Width);
            minY = Math.Min(minY, ink.Y); maxY = Math.Max(maxY, ink.Y + ink.Height);
        }
        var elementX = left + minX;
        var elementY = top + minY;
        var glyphs = run.Glyphs.Select(glyph => new OfdPositionedGlyph(glyph.GlyphId, glyph.ClusterStart,
            glyph.ClusterLength, glyph.XMillimeters - minX, glyph.YMillimeters - minY,
            glyph.AdvanceXMillimeters, glyph.AdvanceYMillimeters)).ToArray();
        var fill = run.Style.Paint.FillColor!;
        var element = new OfdTextElement
        {
            LayerType = "Body", XMillimeters = elementX, YMillimeters = elementY,
            WidthMillimeters = Math.Max(0.1, maxX - minX), HeightMillimeters = Math.Max(0.1, maxY - minY),
            FontName = run.Face.Name, FontResourceId = fontId, FontSizeMillimeters = run.FontSizeMillimeters,
            Text = run.Text, FillColor = new OfdColor(fill.Red, fill.Green, fill.Blue, fill.Alpha),
        };
        element.Runs.Add(new OfdTextRun
        {
            Text = run.Text, XMillimeters = glyphs[0].OriginXMillimeters, YMillimeters = glyphs[0].OriginYMillimeters,
            PositionedGlyphs = new OfdPositionedGlyphRun(run.Text, run.SourceStart, run.Face.FaceSha256, glyphs,
                OfdPositionedGlyphMapping.ClusterMap),
        });
        return element;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
