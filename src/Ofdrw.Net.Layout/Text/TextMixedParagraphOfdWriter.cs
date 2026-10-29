using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Text;

/// <summary>Writes the bounded mixed paragraph's exact final glyph runs to one OFD page.</summary>
public static class TextMixedParagraphOfdWriter
{
    /// <summary>Creates one page with logical word/space elements and positioned original GIDs.</summary>
    public static OfdDocumentPackage CreateSinglePage(TextMixedParagraphResult result,
        double pageWidthMillimeters, double pageHeightMillimeters,
        double leftMillimeters, double topMillimeters)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        if (!Finite(pageWidthMillimeters) || !Finite(pageHeightMillimeters) ||
            !Finite(leftMillimeters) || !Finite(topMillimeters) ||
            pageWidthMillimeters <= 0 || pageHeightMillimeters <= 0 ||
            pageWidthMillimeters > 1_000_000 || pageHeightMillimeters > 1_000_000 ||
            leftMillimeters < 0 || topMillimeters < 0 ||
            leftMillimeters + result.WidthMillimeters > pageWidthMillimeters ||
            topMillimeters + result.HeightMillimeters > pageHeightMillimeters ||
            result.Request.Face.Bold || result.Request.Face.Italic)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The mixed paragraph does not fit the regular horizontal page profile.");
        var request = result.Request;
        var face = request.Face;
        // Reader copy/Find is currently verified only for one fixed-face,
        // unwrapped Hebrew word between two Latin words. Refuse broader query
        // results before creating any package elements or output stream.
        if (result.Lines.Count != 1)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "Mixed output with a line break has not passed the PDF reader profile.",
                sourceStart: result.Lines.Count > 1 ? result.Lines[1].SourceStart : 0,
                sourceLength: result.Lines.Count > 1 ? result.Lines[1].SourceLength : request.Text.Length);
        var profileRuns = result.Lines[0].Runs;
        var hebrewRuns = profileRuns.Where(run => OfdPositionedGlyphProfile.IsHebrewWord(run.Text)).ToArray();
        if (hebrewRuns.Length != 1)
        {
            var offending = hebrewRuns.Length > 1 ? hebrewRuns[1] : profileRuns.FirstOrDefault();
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "Mixed output requires one Hebrew word between two Latin words.",
                sourceStart: offending?.SourceStart ?? 0, sourceLength: offending?.Text.Length ?? request.Text.Length);
        }
        if (face.FaceSha256 != "d7441d42a24882db09640ea4bb768b17a106cdb31ee9749a52cb34a075b8fbcc" ||
            request.FontSizeMillimeters != 4 || profileRuns.Count != 5 ||
            !profileRuns[0].Text.All(value => value is >= 'A' and <= 'Z' or >= 'a' and <= 'z') ||
            profileRuns[0].Text.Length == 0 || profileRuns[1].Text != " " ||
            profileRuns[2] != hebrewRuns[0] || profileRuns[3].Text != " " ||
            !profileRuns[4].Text.All(value => value is >= 'A' and <= 'Z' or >= 'a' and <= 'z') ||
            profileRuns[4].Text.Length == 0)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The mixed output reader profile has not been verified for this word sequence or font.",
                sourceStart: hebrewRuns[0].SourceStart, sourceLength: hebrewRuns[0].Text.Length);
        try { XmlConvert.VerifyXmlChars(face.Name); }
        catch (XmlException error)
        {
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The selected font name cannot be written as OFD XML.", error);
        }
        var scale = request.FontSizeMillimeters / face.UnitsPerEm;
        var height = Math.Max(face.UnitsPerEm,
            face.Ascender - face.Descender + face.LineGap) * scale;
        var elements = new List<OfdTextElement>();
        var nextSource = 0;
        foreach (var line in result.Lines)
        {
            if (line.SourceStart != nextSource || line.Runs.Count == 0 ||
                line.VisibleWidthMillimeters > request.WidthMillimeters + 1e-8)
                throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                    "The mixed paragraph lines lost source or fit identity.");
            var runSource = line.SourceStart;
            foreach (var run in line.Runs)
            {
                var shaped = run.Shaped;
                var text = run.Text;
                if (run.SourceStart != runSource ||
                    run.SourceStart + text.Length > line.SourceStart + line.SourceLength ||
                    request.Text.Substring(run.SourceStart, text.Length) != text ||
                    !ReferenceEquals(shaped.Request.Face, face) ||
                    shaped.Request.SourceStart != run.SourceStart ||
                    shaped.Request.Text != text || shaped.YAdvance != 0 ||
                    text.Length is < 1 or > 8)
                    throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                        "A mixed paragraph run cannot be represented by one source TextCode.");
                var space = text == " ";
                var hebrew = OfdPositionedGlyphProfile.IsHebrewWord(text);
                var latin = !space && text.All(value => value is >= 'A' and <= 'Z' or >= 'a' and <= 'z');
                if (!space && !hebrew && !latin ||
                    space && (shaped.Request.Script != "Latn" || shaped.Request.Direction != TextDirection.LeftToRight) ||
                    hebrew && (shaped.Request.Script != "Hebr" || shaped.Request.Direction != TextDirection.RightToLeft) ||
                    latin && (shaped.Request.Script != "Latn" || shaped.Request.Direction != TextDirection.LeftToRight))
                    throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                        "The mixed paragraph has an unsupported script or direction run.");
                var glyphs = shaped.Glyphs;
                if (space)
                {
                    if (glyphs.Count != 1 || glyphs[0].GlyphId != face.NominalGlyph(' ') ||
                        glyphs[0].ClusterStart != run.SourceStart || glyphs[0].ClusterLength != 1 ||
                        glyphs[0].XOffset != 0 || glyphs[0].YOffset != 0)
                        throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                            "A paragraph separator needs one nominal unoffset space glyph.");
                }
                else if (hebrew)
                {
                    if (glyphs.Count != text.Length ||
                        glyphs.Select(value => value.GlyphId).Distinct().Count() != glyphs.Count)
                        throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                            "A Hebrew word needs one distinct drawn glyph per letter.");
                    for (var index = 0; index < glyphs.Count; index++)
                    {
                        var sourceIndex = text.Length - 1 - index;
                        if (glyphs[index].ClusterStart != run.SourceStart + sourceIndex ||
                            glyphs[index].ClusterLength != 1 ||
                            glyphs[index].GlyphId != face.NominalGlyph(text[sourceIndex]))
                            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                                "The Hebrew word has an unsupported source-to-visual glyph mapping.");
                    }
                }
                else
                {
                    if (glyphs.Count != text.Length &&
                        !(glyphs.Count == 1 && text.Length is 2 or 3))
                        throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                            "The Latin word needs one-to-one glyphs or one complete short ligature.");
                    for (var index = 0; index < glyphs.Count; index++)
                        if (glyphs[index].ClusterStart != run.SourceStart +
                                (glyphs.Count == 1 ? 0 : index) ||
                            glyphs[index].ClusterLength != (glyphs.Count == 1 ? text.Length : 1))
                            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                                "The Latin word has an unsupported source cluster.");
                }
                double penX = 0, penY = 0;
                double? inkLeft = null, inkRight = null, inkTop = null, inkBottom = null;
                var baseline = line.BaselineMillimeters - line.TopMillimeters;
                foreach (var glyph in glyphs)
                {
                    if (glyph.YAdvance != 0)
                        throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                            "Vertical glyph advances exceed the horizontal paragraph profile.");
                    if (glyph.Ink is not null)
                    {
                        var x = (penX + glyph.XOffset + glyph.Ink.X) * scale;
                        var right = x + glyph.Ink.Width * scale;
                        var bottom = baseline - (penY + glyph.YOffset + glyph.Ink.Y) * scale;
                        var top = bottom - glyph.Ink.Height * scale;
                        inkLeft = Math.Min(inkLeft ?? x, x);
                        inkRight = Math.Max(inkRight ?? right, right);
                        inkTop = Math.Min(inkTop ?? top, top);
                        inkBottom = Math.Max(inkBottom ?? bottom, bottom);
                    }
                    penX += glyph.XAdvance; penY += glyph.YAdvance;
                }
                var boundaryLeft = Math.Min(0, inkLeft ?? 0);
                var boundaryTop = Math.Min(0, inkTop ?? 0);
                var boundaryRight = Math.Max(shaped.XAdvance * scale, inkRight ?? 0);
                var boundaryBottom = Math.Max(height, inkBottom ?? 0);
                var elementX = leftMillimeters + run.XMillimeters + boundaryLeft;
                var elementY = topMillimeters + line.TopMillimeters + boundaryTop;
                if (!Finite(elementX) || !Finite(elementY) ||
                    elementX < -1e-8 || elementY < -1e-8 ||
                    elementX + boundaryRight - boundaryLeft > pageWidthMillimeters + 1e-8 ||
                    elementY + boundaryBottom - boundaryTop > pageHeightMillimeters + 1e-8)
                    throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                        "A mixed paragraph glyph would leave the destination page.");
                var element = new OfdTextElement
                {
                    LayerType = "Body", XMillimeters = elementX, YMillimeters = elementY,
                    WidthMillimeters = Math.Max(0.1, boundaryRight - boundaryLeft),
                    HeightMillimeters = Math.Max(0.1, boundaryBottom - boundaryTop),
                    FontName = face.Name, FontResourceId = "F1",
                    FontSizeMillimeters = request.FontSizeMillimeters, Text = text,
                };
                var output = new OfdTextRun
                {
                    Text = text, XMillimeters = -boundaryLeft,
                    YMillimeters = baseline - boundaryTop,
                };
                if (!space)
                {
                    var positioned = new OfdPositionedGlyph[glyphs.Count];
                    penX = 0; penY = 0;
                    for (var index = 0; index < glyphs.Count; index++)
                    {
                        var glyph = glyphs[index];
                        positioned[index] = new OfdPositionedGlyph(glyph.GlyphId,
                            glyph.ClusterStart, glyph.ClusterLength,
                            -boundaryLeft + (penX + glyph.XOffset) * scale,
                            baseline - boundaryTop - (penY + glyph.YOffset) * scale,
                            glyph.XAdvance * scale, glyph.YAdvance * scale);
                        penX += glyph.XAdvance; penY += glyph.YAdvance;
                    }
                    output.XMillimeters = positioned[0].OriginXMillimeters;
                    output.YMillimeters = positioned[0].OriginYMillimeters;
                    output.DeltaX = positioned.Length > 1 ? string.Join(" ", positioned.Skip(1)
                        .Select((glyph, index) => Invariant(glyph.OriginXMillimeters -
                            positioned[index].OriginXMillimeters))) : null;
                    output.DeltaY = Enumerable.Range(1, positioned.Length - 1).Any(index =>
                        positioned[index].OriginYMillimeters != positioned[index - 1].OriginYMillimeters)
                        ? string.Join(" ", positioned.Skip(1).Select((glyph, index) =>
                            Invariant(glyph.OriginYMillimeters - positioned[index].OriginYMillimeters))) : null;
                    output.PositionedGlyphs = new OfdPositionedGlyphRun(text, run.SourceStart,
                        face.FaceSha256, positioned);
                }
                element.Runs.Add(output);
                elements.Add(element);
                runSource += text.Length;
            }
            if (runSource != line.SourceStart + line.SourceLength)
                throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                    "The mixed paragraph line omitted a logical source character.");
            nextSource = runSource;
        }
        if (nextSource != request.Text.Length)
            throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                "The mixed paragraph output omitted the end of its logical source.");
        var package = new OfdDocumentPackage();
        package.Options.DefaultPageWidthMillimeters = pageWidthMillimeters;
        package.Options.DefaultPageHeightMillimeters = pageHeightMillimeters;
        package.Fonts.Add(new OfdFontResource { Id = "F1", FontName = face.Name,
            FamilyName = face.Name, Charset = "unicode", Data = face.GetFaceBytes(), CollectionFaceIndex = 0 });
        var page = new OfdPage { Id = "1", Index = 0,
            WidthMillimeters = pageWidthMillimeters, HeightMillimeters = pageHeightMillimeters };
        foreach (var element in elements) page.Elements.Add(element);
        package.Pages.Add(page);
        return package;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static string Invariant(double value) => value.ToString("0.###############", CultureInfo.InvariantCulture);
}
