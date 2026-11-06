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
/// <para>Output profile: Unicode text including supplementary characters (CodePosition/CodeCount count
/// scalar values); fill and/or stroke paint in DeviceRGB with alpha, which carries synthetic bold;
/// synthetic oblique as a baseline shear CTM; highlight, underline and strikethrough as filled
/// PathObjects; TrueType or static CFF faces that permit embedding (OS/2.fsType). Fonts are embedded
/// as GID-preserving subsets by default. Anything else is rejected with
/// <see cref="TextLayoutError.UnsupportedOutput"/> before a package is created.</para>
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
                AppendLine(page.Elements, line, leftMillimeters, topMillimeters, face => resources[face.FaceSha256]);
            package.Pages.Add(page);
        }
        return package;
    }

    /// <summary>Validates one run against the output profile before any package object is created.</summary>
    internal static void Preflight(ParagraphLayoutResult result, ParagraphLayoutRun run)
    {
        PreflightRun(run);
        if (result.Text.Substring(run.SourceStart, run.SourceLength) != run.Text)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The run lost its source identity.", sourceStart: run.SourceStart, sourceLength: run.SourceLength);
    }

    /// <summary>Validates one run's style, glyph identity and cluster mapping against the output profile.</summary>
    internal static void PreflightRun(ParagraphLayoutRun run)
    {
        var style = run.Style;
        // The face satisfies the request directly, or the recorded synthetic styles make up the difference.
        var bold = style.FaceBold || (style.Synthetic & TextSyntheticStyle.Bold) != 0;
        var italic = style.FaceItalic || (style.Synthetic & TextSyntheticStyle.Oblique) != 0;
        if (bold != style.RequestedBold || italic != style.RequestedItalic ||
            (style.Synthetic & TextSyntheticStyle.Bold) != 0 && !(style.Paint.Stroke && style.Paint.StrokeColor!.Equals(style.Paint.FillColor)) ||
            (style.Synthetic & TextSyntheticStyle.Oblique) != 0 && style.ObliqueSkew != TextResolvedStyle.SyntheticObliqueSkew)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput, "The run's resolved style is inconsistent with its face.",
                sourceStart: run.SourceStart, sourceLength: run.SourceLength);
        if (run.Text.Length > 4096 || run.Glyphs.Count > 16384)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "A run exceeds the OFD TextCode budget.",
                sourceStart: run.SourceStart, sourceLength: run.SourceLength);
        if (run.ShapedRun.Request.Text != run.Text || run.ShapedRun.Request.SourceStart != run.SourceStart ||
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

    /// <summary>Appends one final line: highlights, then one TextObject per run, then underline and
    /// strikethrough rectangles. All coordinates are offset by (left, top).</summary>
    internal static void AppendLine(IList<OfdElement> elements, ParagraphLayoutLine line, double left, double top,
        Func<TextFontFace, string> fontId)
    {
        foreach (var run in line.Runs)
            foreach (var decoration in run.Decorations)
                if (decoration.Kind == ParagraphLayoutDecorationKind.Highlight) elements.Add(Rectangle(decoration, left, top));
        foreach (var run in line.Runs) elements.Add(CreateElement(run, line, left, top, fontId(run.Face)));
        foreach (var run in line.Runs)
            foreach (var decoration in run.Decorations)
                if (decoration.Kind != ParagraphLayoutDecorationKind.Highlight) elements.Add(Rectangle(decoration, left, top));
    }

    private static OfdPathElement Rectangle(ParagraphLayoutDecoration decoration, double left, double top)
    {
        var bounds = decoration.BoundsMillimeters;
        var color = decoration.Color;
        var width = Math.Max(bounds.Width, 0.001);
        var height = Math.Max(bounds.Height, 0.001);
        return new OfdPathElement
        {
            LayerType = "Body", XMillimeters = left + bounds.X, YMillimeters = top + bounds.Y,
            WidthMillimeters = width, HeightMillimeters = height, Fill = true, Stroke = false,
            FillColor = new OfdColor(color.Red, color.Green, color.Blue, color.Alpha),
            AbbreviatedData = string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "M 0 0 L {0:0.######} 0 L {0:0.######} {1:0.######} L 0 {1:0.######} C", width, height),
        };
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
        var paint = run.Style.Paint;
        var fill = paint.FillColor ?? paint.StrokeColor!;
        var element = new OfdTextElement
        {
            LayerType = "Body", XMillimeters = elementX, YMillimeters = elementY,
            WidthMillimeters = Math.Max(0.1, maxX - minX), HeightMillimeters = Math.Max(0.1, maxY - minY),
            FontName = run.Face.Name, FontResourceId = fontId, FontSizeMillimeters = run.FontSizeMillimeters,
            Text = run.Text, FillColor = new OfdColor(fill.Red, fill.Green, fill.Blue, fill.Alpha),
            ProducerFill = paint.Fill,
            ProducerStrokeColor = paint.StrokeColor is { } stroke ? new OfdColor(stroke.Red, stroke.Green, stroke.Blue, stroke.Alpha) : null,
            ProducerLineWidthMillimeters = paint.StrokeWidthMillimeters,
        };
        if (run.Style.ObliqueSkew != 0)
        {
            // Shear about the run's baseline in TextObject space: x' = x − skew × (y − baseline).
            var baseline = run.BaselineYMillimeters - minY;
            element.Transform = new[] { 1, 0, -run.Style.ObliqueSkew, 1, run.Style.ObliqueSkew * baseline, 0 };
        }
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
