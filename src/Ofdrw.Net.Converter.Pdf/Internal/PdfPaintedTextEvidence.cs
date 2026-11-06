using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Graphics.Operations;
using UglyToad.PdfPig.Graphics.Operations.PathConstruction;
using UglyToad.PdfPig.Graphics.Operations.SpecialGraphicsState;
using UglyToad.PdfPig.Graphics.Operations.TextState;
using UglyToad.PdfPig.Tokens;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// Glyph geometry is evidence of the visible page only when the glyph is
// actually painted. Invisible text layers (render mode 3, clipped away, zero
// alpha, hidden optional content) carry a producer's claim whose positions need
// not match the visible glyphs, so the visual-line profile does not consume
// them. PdfPig 0.1.15 does not expose per-letter clip or alpha, so this scans
// the page's own content operations and maps each show-text operation to
// Letter.TextSequence. Anything it cannot account for is reported as unknown.
internal sealed class PdfPaintedTextEvidence
{
    private const double ClipTolerance = 0.5;
    private static readonly NameToken FillAlphaKey = NameToken.Create("ca");
    private static readonly NameToken StrokeAlphaKey = NameToken.Create("CA");
    private static readonly NameToken SoftMaskKey = NameToken.Create("SMask");
    private static readonly NameToken ExtGStateKey = NameToken.Create("ExtGState");

    private sealed class State
    {
        internal double[] Ctm = { 1, 0, 0, 1, 0, 0 };
        internal List<PdfRectangle> Clips = new();
        internal bool ComplexClip;
        internal double? FillAlpha = 1, StrokeAlpha = 1;
        internal bool TextClipMode;
        internal State Copy() => new()
        {
            Ctm = (double[])Ctm.Clone(), Clips = new List<PdfRectangle>(Clips), ComplexClip = ComplexClip,
            FillAlpha = FillAlpha, StrokeAlpha = StrokeAlpha, TextClipMode = TextClipMode,
        };
    }

    private readonly List<State> _shows;
    private readonly string? _unverified;

    private PdfPaintedTextEvidence(List<State> shows, string? unverified) { _shows = shows; _unverified = unverified; }

    internal static PdfPaintedTextEvidence Scan(PdfDocument document, Page page, CancellationToken cancellationToken)
    {
        var shows = new List<State>();
        var stack = new Stack<State>();
        var state = new State();
        var pathRects = new List<PdfRectangle>();
        var pathComplex = false;
        var pendingClip = false;
        var hasXObject = false;
        string? unverified = null;
        foreach (var operation in page.Operations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            switch (operation.Operator)
            {
                case "q": stack.Push(state.Copy()); break;
                case "Q": if (stack.Count > 0) state = stack.Pop(); break;
                case "cm" when operation is ModifyCurrentTransformationMatrix cm:
                    state.Ctm = Multiply(cm.Value.Select(Convert.ToDouble).ToArray(), state.Ctm);
                    break;
                case "gs" when operation is SetGraphicsStateParametersFromDictionary gs:
                    ApplyExtGState(document, page, gs.Name, state);
                    break;
                case "re" when operation is AppendRectangle re:
                    var rect = Transform(state.Ctm, re.LowerLeftX, re.LowerLeftY, re.Width, re.Height);
                    if (rect is { } axisAligned) pathRects.Add(axisAligned); else pathComplex = true;
                    break;
                case "m": case "l": case "c": case "v": case "y": case "h":
                    pathComplex = true;
                    break;
                case "W": case "W*":
                    pendingClip = true;
                    break;
                case "n": case "S": case "s": case "f": case "F": case "f*": case "B": case "B*": case "b": case "b*":
                    if (pendingClip)
                    {
                        if (pathComplex || pathRects.Count != 1) state.ComplexClip = true;
                        else state.Clips.Add(pathRects[0]);
                    }
                    pendingClip = false; pathComplex = false; pathRects.Clear();
                    break;
                case "Tr" when operation is SetTextRenderingMode tr:
                    state.TextClipMode = tr.Mode >= TextRenderingMode.FillClip;
                    break;
                case "ET":
                    // Glyph outlines shown in a clip mode join the clipping path at ET.
                    if (state.TextClipMode) state.ComplexClip = true;
                    break;
                case "Tj": case "TJ": case "'": case "\"":
                    shows.Add(state.Copy());
                    break;
                case "Do":
                    hasXObject = true;
                    break;
            }
        }
        var maximum = page.Letters.Count == 0 ? 0 : page.Letters.Max(letter => letter.TextSequence);
        if (maximum != shows.Count)
            unverified = hasXObject
                ? "text is shown inside a form XObject, whose painting state the profile does not verify."
                : "its text operations could not be matched to the decoded glyphs.";
        return new PdfPaintedTextEvidence(shows, unverified);
    }

    // Null when the letter is directly painted and fully inside every clip;
    // otherwise the reason it is not visible-glyph evidence.
    internal string? Reject(Letter letter)
    {
        if (_unverified is not null) return _unverified;
        var fill = letter.RenderingMode is TextRenderingMode.Fill or TextRenderingMode.FillThenStroke;
        var stroke = letter.RenderingMode is TextRenderingMode.Stroke or TextRenderingMode.FillThenStroke;
        if (!fill && !stroke)
            return $"a glyph uses text rendering mode {letter.RenderingMode}, so it is not a painted glyph.";
        var index = letter.TextSequence - 1;
        if (index < 0 || index >= _shows.Count) return "a glyph has no matching text operation.";
        var state = _shows[index];
        if (state.ComplexClip) return "a glyph is shown under a non-rectangular clipping path.";
        double? alpha = fill && stroke ? Max(state.FillAlpha, state.StrokeAlpha) : fill ? state.FillAlpha : state.StrokeAlpha;
        if (alpha is null) return "a glyph uses graphics-state transparency the profile cannot resolve.";
        if (alpha <= 0) return "a glyph is fully transparent.";
        var box = letter.BoundingBox;
        foreach (var clip in state.Clips)
            if (box.Left < clip.Left - ClipTolerance || box.Right > clip.Right + ClipTolerance ||
                box.Bottom < clip.Bottom - ClipTolerance || box.Top > clip.Top + ClipTolerance)
                return "a glyph is clipped.";
        return null;
    }

    private static double? Max(double? left, double? right) =>
        left is null || right is null ? null : Math.Max(left.Value, right.Value);

    private static void ApplyExtGState(PdfDocument document, Page page, NameToken name, State state)
    {
        DictionaryToken? graphics = null;
        if (Resolve(document, page.Dictionary, NameToken.Resources) is DictionaryToken resources &&
            Resolve(document, resources, ExtGStateKey) is DictionaryToken states)
            graphics = Resolve(document, states, name) as DictionaryToken;
        if (graphics is null)
        {
            state.FillAlpha = state.StrokeAlpha = null;
            return;
        }
        if (graphics.ContainsKey(FillAlphaKey)) state.FillAlpha = Number(Resolve(document, graphics, FillAlphaKey));
        if (graphics.ContainsKey(StrokeAlphaKey)) state.StrokeAlpha = Number(Resolve(document, graphics, StrokeAlphaKey));
        if (graphics.ContainsKey(SoftMaskKey) &&
            !(Resolve(document, graphics, SoftMaskKey) is NameToken mask && mask.Data == "None"))
            state.FillAlpha = state.StrokeAlpha = null;
    }

    private static double? Number(IToken? token) => token is NumericToken number ? number.Double : null;

    private static IToken? Resolve(PdfDocument document, DictionaryToken dictionary, NameToken key)
    {
        if (!dictionary.TryGet(key, out var token)) return null;
        for (var depth = 0; depth < 8 && token is IndirectReferenceToken reference; depth++)
            token = document.Structure.GetObject(reference.Data).Data;
        return token;
    }

    private static double[] Multiply(double[] m, double[] n) => new[]
    {
        m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3],
        m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
        m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5],
    };

    private static PdfRectangle? Transform(double[] ctm, double x, double y, double width, double height)
    {
        if (ctm[1] != 0 || ctm[2] != 0) return null;
        var x0 = ctm[0] * x + ctm[4]; var x1 = ctm[0] * (x + width) + ctm[4];
        var y0 = ctm[3] * y + ctm[5]; var y1 = ctm[3] * (y + height) + ctm[5];
        return new PdfRectangle(Math.Min(x0, x1), Math.Min(y0, y1), Math.Max(x0, x1), Math.Max(y0, y1));
    }
}
