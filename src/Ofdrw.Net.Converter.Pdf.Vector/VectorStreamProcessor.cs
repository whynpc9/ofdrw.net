using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Graphics.SkiaSharp;
using Ofdrw.Net.Core.Models;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Geometry;
using UglyToad.PdfPig.Graphics;
using UglyToad.PdfPig.Graphics.Core;
using UglyToad.PdfPig.Graphics.Operations;
using UglyToad.PdfPig.Graphics.Operations.TextShowing;
using UglyToad.PdfPig.Graphics.Operations.SpecialGraphicsState;
using UglyToad.PdfPig.Parser;
using UglyToad.PdfPig.PdfFonts;
using UglyToad.PdfPig.Tokens;
using SkiaSharp;

namespace Ofdrw.Net.Converter.Pdf.Vector;

internal sealed class VectorStreamProcessor : BaseStreamProcessor<List<SkiaDrawEvent>>, IDisposable
{
    private readonly VectorPageContext _context; private readonly double _height; private readonly VectorPageFrame _frame;
    private readonly Dictionary<string, VectorFont> _fonts; private readonly PdfVectorToOfdOptions _limits;
    private readonly DictionaryToken _pageResources; private readonly long _remainingImageBytes; private long _imageBytes;
    private readonly CancellationToken _token; private readonly SKPath _path = new(); private readonly List<SkiaDrawEvent> _events = new();
    private readonly Action<string> _ensureFont, _ensureHorizontalFont;
    private readonly List<ClipShape> _activeClips = new();
    private readonly Stack<int> _clipDepths = new();
    private readonly List<ClipShape[]> _eventClips = new();
    private readonly List<StrokeShape?> _eventStrokes = new();
    private readonly List<OfdImageElement?> _paintOrder = new();
    private readonly List<ClipShape[]?> _imageClips = new();
    private FillingRule? _pendingClipRule;
    internal readonly HashSet<VectorFont> OutputFonts = new();
    private readonly StringBuilder _text = new(); private readonly List<float> _advances = new();
    private TextFloatPrecision? _textPrecision, _writerTextPrecision; private bool _writerTextFailed;
    private SKMatrix _candidateTextMatrix; private float _candidateTextSize;
    private double _lastTextPosition;
    private TransformationMatrix _textMatrix; private VectorFont? _textFont; private double _textSize;
    private bool _hasSegments, _subpathHasSegments, _hasClosedSingleton;
    private PathFloatPrecision? _precision;
    private SKPoint _currentCandidate, _startCandidate;
    private SKColor _textColor; private bool _inText; private int _commands, _characters; private PdfPoint? _subpathStart;

    internal VectorStreamProcessor(VectorPageContext context, IPageContentParser parser, VectorPageFrame frame,
        DictionaryToken pageResources, Dictionary<string, VectorFont> fonts, PdfVectorToOfdOptions limits, CancellationToken token, int number,
        long remainingImageBytes, Action<string> ensureFont, Action<string> ensureHorizontalFont)
        : base(number, context.Resources, context.Scanner, parser, context.Filters, new CropBox(new PdfRectangle(0, 0, frame.Width, frame.Height)),
            UserSpaceUnit.Default, new PageRotationDegrees(0), TransformationMatrix.Identity, context.Parsing)
    { _context = context; _height = frame.Height; _frame = frame; _pageResources = pageResources; _fonts = fonts; _limits = limits; _token = token;
        _remainingImageBytes = remainingImageBytes;
        _ensureFont = ensureFont; _ensureHorizontalFont = ensureHorizontalFont; }

    internal int PaintCount => _paintOrder.Count;

    public override List<SkiaDrawEvent> Process(int number, IReadOnlyList<IGraphicsStateOperation> operations)
    {
        if (!_frame.IsLegacyFrame) new ModifyCurrentTransformationMatrix(_frame.InitialTransform).Run(this);
        foreach (var operation in operations)
        {
            _token.ThrowIfCancellationRequested();
            if (_path.VerbCount != 0 && operation.Operator is "cm" or "q" or "Q") throw Unsupported("PATH_STATE_CHANGE");
            if (operation is ShowText show)
            {
                // Stock ShowText is nonvirtual and gets a font before checking
                // its bytes. An empty show only advances its sequence counter.
                if (show.Text is null ? show.Bytes.IsEmpty : show.Text.Length == 0) TextSequence++;
                else { _ensureFont(SelectedFontName()); operation.Run(this); }
            }
            else if (operation is InvokeNamedXObject invoke)
            {
                var image = _context.RequireImageXObject(_pageResources, invoke.Name);
                _context.Resources.LoadResourceDictionary(new DictionaryToken(new Dictionary<NameToken, IToken> {
                    [NameToken.Create("XObject")] = new DictionaryToken(new Dictionary<NameToken, IToken> { [invoke.Name] = image }) }));
                try { operation.Run(this); }
                finally { _context.Resources.UnloadResourceDictionary(); }
            }
            else operation.Run(this);
            if (operation.Operator is "Tj" or "TJ") FlushText();
        }
        if (_inText || StackSize != 1 || _path.VerbCount != 0) throw new InvalidDataException("PDF has an unterminated text, graphics state or path.");
        return _events;
    }
    public override void ShowPositionedText(IReadOnlyList<IToken> tokens)
    {
        var hasText = false; var hasNumbers = false;
        foreach (var item in tokens)
        {
            _token.ThrowIfCancellationRequested();
            if (item is NumericToken) hasNumbers = true;
            else if (item is HexToken hex) hasText |= hex.Bytes.Length != 0;
            else if (item is StringToken text) hasText |= text.GetBytes().Length != 0;
            else throw new InvalidDataException("PDF TJ requires strings or numbers.");
        }
        var state = GetCurrentState().FontState;
        if (hasText)
        {
            _ensureFont(SelectedFontName());
            base.ShowPositionedText(tokens);
            return;
        }
        TextSequence++;
        if (!hasNumbers) return;
        // Only direction metadata is needed for numeric-only horizontal TJ.
        // Preserve each stock adjustment and leave TextLineMatrix unchanged.
        _ensureHorizontalFont(SelectedFontName());
        foreach (var item in tokens)
        {
            _token.ThrowIfCancellationRequested();
            if (item is not NumericToken number) continue;
            var tx = -number.Data / 1000 * state.FontSize * (state.HorizontalScaling / 100.0);
            TextMatrices.TextMatrix = TransformationMatrix.GetTranslationMatrix(tx, 0).Multiply(TextMatrices.TextMatrix);
        }
    }
    private string SelectedFontName() => GetCurrentState().FontState.FontName?.Data
        ?? throw new InvalidDataException("PDF text show requires a selected font.");
    public override void PushState()
    { if (StackSize >= _limits.MaxStackDepth) throw new InvalidDataException("PDF graphics stack limit exceeded."); base.PushState(); _clipDepths.Push(_activeClips.Count); }
    public override void PopState()
    {
        if (StackSize <= 1) throw new InvalidDataException("PDF graphics stack underflow.");
        base.PopState(); var count = _clipDepths.Pop();
        if (_activeClips.Count > count) _activeClips.RemoveRange(count, _activeClips.Count - count);
    }
    public override void BeginText()
    { if (_inText) throw Unsupported("NESTED_TEXT"); _inText = true; base.BeginText(); }
    public override void EndText()
    { if (!_inText) throw Unsupported("TEXT_UNDERFLOW"); _inText = false; base.EndText(); }
    private NotSupportedException Unsupported(string reason) => _context.Unsupported(reason);
    private void RequireEvents(int required)
    {
        _token.ThrowIfCancellationRequested();
        if (required > _limits.MaxEventsPerPage - _paintOrder.Count) throw new InvalidDataException("PDF draw event limit exceeded.");
    }
    private void Add(SkiaDrawEvent value, StrokeShape? stroke = null)
    {
        RequireEvents(1);
        _events.Add(value); _eventClips.Add(_activeClips.ToArray()); _eventStrokes.Add(stroke);
        _paintOrder.Add(null); _imageClips.Add(null);
    }

    internal void ApplyAnnotations(OfdPage page, string ns)
    {
        if (page.Elements.Count != _eventClips.Count) throw new InvalidDataException("PDF vector event/object ordering mismatch.");
        var xmlBytes = 0L; var name = XNamespace.Get(ns);
        for (var index = 0; index < _eventClips.Count; index++)
        {
            _token.ThrowIfCancellationRequested();
            var element = page.Elements[index];
            if (_eventStrokes[index] is { } style)
            {
                if (element is not OfdPathElement path || !path.Stroke || path.SourceXml is null)
                    throw new InvalidDataException("PDF stroke event/object ordering mismatch.");
                var source = XElement.Parse(path.SourceXml);
                source.SetAttributeValue("Cap", style.Cap);
                source.SetAttributeValue("Join", style.Join);
                source.SetAttributeValue("MiterLimit", OfdNumericFormat.Plain(style.Miter));
                source.SetAttributeValue("DashPattern", style.Dashes.Length == 0 ? null : string.Join(" ", style.Dashes.Select(OfdNumericFormat.Plain)));
                source.SetAttributeValue("DashOffset", style.Dashes.Length == 0 ? null : OfdNumericFormat.Plain(style.Phase));
                path.SourceXml = source.ToString(SaveOptions.DisableFormatting);
            }
            element.ClippingXml = ClipXml(_eventClips[index], element, name, ref xmlBytes);
        }
        var native = page.Elements.ToArray(); var nativeIndex = 0;
        page.Elements.Clear();
        for (var index = 0; index < _paintOrder.Count; index++)
        {
            _token.ThrowIfCancellationRequested();
            var image = _paintOrder[index];
            if (image is null) page.Elements.Add(native[nativeIndex++]);
            else
            {
                image.ClippingXml = ClipXml(_imageClips[index]!, image, name, ref xmlBytes);
                page.Elements.Add(image);
            }
        }
        if (nativeIndex != native.Length) throw new InvalidDataException("PDF vector event/object ordering mismatch.");
    }

    private string? ClipXml(ClipShape[] clips, OfdElement element, XNamespace name, ref long xmlBytes)
    {
            if (clips.Length == 0) return null;
            var x = OfdNumericFormat.WriterValue(element.XMillimeters);
            var y = OfdNumericFormat.WriterValue(element.YMillimeters);
            var root = new XElement(name + "Clips");
            foreach (var clip in clips)
            {
                _token.ThrowIfCancellationRequested();
                var path = new XElement(name + "Path", new XAttribute("Stroke", "false"), new XAttribute("Fill", "true"),
                    new XAttribute("Rule", clip.EvenOdd ? "Even-Odd" : "NonZero"),
                    new XAttribute("CTM", "1 0 0 1 " + OfdNumericFormat.Plain(-x) + " " + OfdNumericFormat.Plain(-y)),
                    new XElement(name + "AbbreviatedData", clip.Data));
                root.Add(new XElement(name + "Clip", new XElement(name + "Area", path)));
            }
            var output = root.ToString(SaveOptions.DisableFormatting);
            if (output.Length > _limits.MaxContentBytesPerPage - xmlBytes)
                throw new InvalidDataException("PDF emitted clipping geometry limit exceeded.");
            xmlBytes += output.Length;
            return output;
    }

    private readonly struct ClipShape
    {
        internal readonly string Data; internal readonly bool EvenOdd;
        internal ClipShape(string data, bool evenOdd) { Data = data; EvenOdd = evenOdd; }
    }

    private readonly struct StrokeShape
    {
        internal readonly string Cap, Join; internal readonly double Miter, Phase; internal readonly double[] Dashes;
        internal StrokeShape(string cap, string join, double miter, double phase, double[] dashes)
        { Cap = cap; Join = join; Miter = miter; Phase = phase; Dashes = dashes; }
    }

    private StrokeShape CaptureStroke()
    {
        var state = GetCurrentState();
        if (!(state.LineWidth > 0) || double.IsNaN(state.LineWidth) || double.IsInfinity(state.LineWidth) ||
            !(state.MiterLimit >= 1 && state.MiterLimit <= 10) || state.MiterLimit != Math.Truncate(state.MiterLimit))
            throw Unsupported("STROKE_PROFILE: positive width and integral miter limit 1..10 required.");
        var cap = state.CapStyle switch
        {
            LineCapStyle.Butt => "Butt", LineCapStyle.Round => "Round", LineCapStyle.ProjectingSquare => "Square",
            _ => throw Unsupported("STROKE_PROFILE: unknown cap.")
        };
        var join = state.JoinStyle switch
        {
            LineJoinStyle.Miter => "Miter", LineJoinStyle.Round => "Round", LineJoinStyle.Bevel => "Bevel",
            _ => throw Unsupported("STROKE_PROFILE: unknown join.")
        };
        if (cap != "Butt" && state.MiterLimit != 10)
            throw Unsupported("STROKE_PROFILE: PDF export requires miter limit 10 with round/square cap.");
        var dash = state.LineDashPattern;
        var values = dash.Array?.ToArray() ?? Array.Empty<double>();
        if (values.Length > 32 || values.Any(value => !(value > 0) || double.IsNaN(value) || double.IsInfinity(value)) ||
            dash.Phase < 0 || (values.Length == 0 && dash.Phase != 0) ||
            values.Length != 0 && (double.IsInfinity(values.Sum()) || double.IsInfinity(dash.Phase / state.LineWidth)))
            throw Unsupported("STROKE_PROFILE: dash must be bounded, positive and representable by PDF export.");
        const double millimetersPerPoint = 25.4 / 72;
        return new StrokeShape(cap, join, state.MiterLimit, dash.Phase * millimetersPerPoint,
            values.Select(value => value * millimetersPerPoint).ToArray());
    }
    private void Command(int count = 1)
    {
        _token.ThrowIfCancellationRequested();
        if (count > _limits.MaxPathCommandsPerPage - _commands) throw new InvalidDataException("PDF path command limit exceeded.");
        _commands += count;
    }
    private static float Float(double value)
    {
        var result = (float)value;
        if (float.IsNaN(result) || float.IsInfinity(result)) throw new NotSupportedException("PDFV_NONFINITE_OR_FLOAT_RANGE");
        return result;
    }
    private SKMatrix Matrix(TransformationMatrix value, bool text = false)
    {
        var matrix = new SKMatrix(Float(value.A), Float(text ? -value.C : value.C), Float(value.E),
            Float(-value.B), Float(text ? value.D : -value.D), Float(_height - value.F), 0, 0, 1);
        // Use 04's exact serialized-coefficient contract at both adapter stages.
        const double units = 25.4 / 72;
        if (!OfdNumericFormat.Nonsingular(matrix.ScaleX, matrix.SkewY, matrix.SkewX, matrix.ScaleY) ||
            !OfdNumericFormat.Nonsingular(units * matrix.ScaleX, units * matrix.SkewY, units * matrix.SkewX, units * matrix.ScaleY))
            throw Unsupported("SINGULAR_SERIALIZED_MATRIX");
        return matrix;
    }
    private SKColor Color(bool stroke)
    {
        var state = GetCurrentState();
        var rgb = (stroke ? state.CurrentStrokingColor : state.CurrentNonStrokingColor)?.ToRGBValues() ?? (0d, 0d, 0d);
        byte Channel(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0 || value > 1) throw Unsupported("COLOR_RANGE");
            return (byte)Math.Round(value * 255);
        }
        return new SKColor(Channel(rgb.Item1), Channel(rgb.Item2), Channel(rgb.Item3));
    }
    private SKPaint Paint(bool stroke)
    {
        var state = GetCurrentState();
        if (stroke) _ = CaptureStroke();
        return new SKPaint { Color = Color(stroke), Style = stroke ? SKPaintStyle.Stroke : SKPaintStyle.Fill,
            StrokeWidth = Float(state.LineWidth), StrokeMiter = 10 };
    }
    private void PaintPath(bool fill, bool stroke, FillingRule rule, bool close)
    {
        if (close) ClosePath();
        if (_path.VerbCount == 0) { EndPath(); return; }
        // Only a real/potential paint needs capacity; move-only no-ops stay free.
        // Admit both fill/stroke events before a profile failure can choose fallback.
        if (_hasSegments || (fill && _hasClosedSingleton))
            RequireEvents((fill ? 1 : 0) + (stroke ? 1 : 0));
        // A closed singleton fill can paint a device pixel; do not silently elide it.
        if (fill && _hasClosedSingleton) throw Unsupported("DEGENERATE_POINT_FILL");
        if (!_hasSegments)
        {
            // Validate the supported paint profile before treating a source no-op as such.
            if (fill) { using var paint = Paint(false); }
            if (stroke)
            {
                using var paint = Paint(true);
                if (GetCurrentState().CapStyle != LineCapStyle.Butt)
                    throw Unsupported("STROKE_PROFILE: a round or square cap can paint an otherwise empty path.");
            }
            EndPath(); return;
        }
        if (fill) RememberImplicitClosure();
        var matrix = Matrix(CurrentTransformationMatrix); // Keep existing singular diagnostics.
        using var fillPaint = fill ? Paint(false) : null;
        using var strokePaint = stroke ? Paint(true) : null;
        if (!_precision!.Accept(fill, stroke, GetCurrentState().LineWidth, strokePaint?.StrokeWidth ?? 0))
            throw Unsupported("PATH_FLOAT_PRECISION: added path/control/CTM/stroke conversion error exceeds bounded profile.");
        // 19 passes this local float directly to OfdPen; pt-to-mm and PDF CTM
        // scale belong in the transform, not in the serialized scalar check.
        if (stroke && !OfdNumericFormat.WriterPositive(strokePaint!.StrokeWidth))
            throw Unsupported("STROKE_WRITER_PRECISION: local width must remain positive at OFD writer precision.");
        if (stroke && !_precision!.AcceptWrittenStroke(GetCurrentState().LineWidth, OfdNumericFormat.WriterValue(strokePaint!.StrokeWidth), matrix))
            throw Unsupported("STROKE_WRITER_PRECISION: total written stroke-map error exceeds bounded profile.");
        _path.FillType = rule == FillingRule.EvenOdd ? SKPathFillType.EvenOdd : SKPathFillType.Winding;
        if (fill) Add(SkiaDrawEvent.Path(_path, fillPaint!, matrix, _limits.MaxPathCommandsPerPage));
        if (stroke) Add(SkiaDrawEvent.Path(_path, strokePaint!, matrix, _limits.MaxPathCommandsPerPage), CaptureStroke());
        EndPath();
    }
    public override void RenderGlyph(IFont font, CurrentGraphicsState state, double fontSize, double pointSize, int code, string unicode,
        long offset, in TransformationMatrix rendering, in TransformationMatrix text, in TransformationMatrix ctm, CharacterBoundingBox bounds)
    {
        _token.ThrowIfCancellationRequested();
        if (_characters >= _limits.MaxTextCharactersPerPage) throw new InvalidDataException("PDF original text character limit exceeded.");
        RequireEvents(1);
        if (!_inText || !font.TryGetUnicode(code, out var original) || original != unicode || string.IsNullOrEmpty(original) || original.Length != 1 ||
            char.IsSurrogate(original[0]) || char.IsControl(original[0]) || font.IsVertical)
            throw Unsupported("UNICODE_MAPPING: one original BMP character per glyph required.");
        var category = CharUnicodeInfo.GetUnicodeCategory(original[0]);
        if (category is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark)
            throw Unsupported("COMBINING_SHAPING");
        if (!(fontSize > 0) || state.FontState.HorizontalScaling != 100 || state.FontState.Rise != 0) throw Unsupported("TEXT_STATE");
        if (!_fonts.TryGetValue(state.FontState.FontName.Data, out var binding)) throw Unsupported("FONT_RESOURCE");
        var glyph = binding.Binding.GetGlyph(original[0]);
        if (glyph == 0 || glyph != binding.GlyphForCid(code)) throw Unsupported("SOURCE_GLYPH_CMAP_MISMATCH: font=" + state.FontState.FontName.Data + " code=" + code);
        _characters++;
        var matrix = text.Multiply(ctm); var color = Color(false);
        var first = _text.Length == 0; var advance = 0f;
        if (first)
        {
            _textMatrix = matrix; _textFont = binding; _textSize = fontSize; _textColor = color; _lastTextPosition = 0;
            _candidateTextMatrix = Matrix(matrix, true); _candidateTextSize = (float)fontSize;
            _textPrecision = new TextFloatPrecision(_candidateTextMatrix, _candidateTextSize, _height);
            _writerTextPrecision = OfdNumericFormat.WriterPositive(_candidateTextSize)
                ? TextFloatPrecision.ForWriter(_candidateTextMatrix, _candidateTextSize, _height) : null;
            _writerTextFailed = false;
        }
        else
        {
            if (_textFont != binding || _textSize != fontSize || _textColor != color || matrix.A != _textMatrix.A || matrix.B != _textMatrix.B || matrix.C != _textMatrix.C || matrix.D != _textMatrix.D)
                throw Unsupported("MIXED_TEXT_RUN");
            var determinant = matrix.A * matrix.D - matrix.B * matrix.C;
            var dx = matrix.E - _textMatrix.E; var dy = matrix.F - _textMatrix.F;
            var x = (matrix.D * dx - matrix.C * dy) / determinant;
            var y = (-matrix.B * dx + matrix.A * dy) / determinant;
            if (double.IsNaN(x) || double.IsInfinity(x) || Math.Abs(y) > 0.000001) throw Unsupported("TEXT_BASELINE");
            advance = Float(x - _lastTextPosition); _lastTextPosition = x;
        }
        if (!_textPrecision!.Glyph(text, ctm, fontSize, bounds.GlyphBounds, advance))
            throw Unsupported("TEXT_FLOAT_PRECISION: added matrix/size/glyph/accumulated advance conversion error exceeds bounded profile.");
        // Latch only this new certificate: a later glyph must still encounter
        // its existing character/event hard gates before a run can fall back.
        _writerTextFailed = _writerTextFailed || (_writerTextPrecision is not null &&
            !_writerTextPrecision.Glyph(text, ctm, fontSize, bounds.GlyphBounds, advance));
        if (!first) _advances.Add(advance);
        _text.Append(original);
    }
    private void FlushText()
    {
        if (_text.Length == 0) return;
        var original = _text.ToString();
        // Empty runs return above; nonempty runs must fit before any run profile.
        RequireEvents(1);
        if (string.IsNullOrWhiteSpace(original) || StringInfo.ParseCombiningCharacters(original).Length != _advances.Count + 1) throw Unsupported("WHITESPACE_ONLY_OR_SHAPED_RUN");
        _token.ThrowIfCancellationRequested();
        if (!OfdNumericFormat.WriterPositive(_candidateTextSize))
            throw Unsupported("TEXT_WRITER_PRECISION: local font size must remain positive at OFD writer precision.");
        foreach (var character in original)
            if (!XmlConvert.IsXmlChar(character))
                throw Unsupported("TEXT_XML_CHARACTER: original text cannot be represented in OFD XML 1.0.");
        if (_writerTextFailed)
            throw Unsupported("TEXT_WRITER_PRECISION: total written glyph/map/baseline/advance error exceeds bounded profile.");
        using var paint = new SKPaint { Color = _textColor, Style = SKPaintStyle.Fill };
        Add(_textFont!.Binding.SnapshotText(original, new SKPoint(0, 0), _candidateTextSize, paint, _advances,
            _candidateTextMatrix, _limits.MaxTextCharactersPerPage, _token));
        OutputFonts.Add(_textFont);
        _text.Clear(); _advances.Clear(); _textFont = null; _textPrecision = _writerTextPrecision = null; _writerTextFailed = false;
    }
    public override void BeginSubpath() { }
    public override PdfPoint? CloseSubpath() { ClosePath(); return _subpathStart; }
    public override void StrokePath(bool close) => PaintPath(false, true, FillingRule.NonZeroWinding, close);
    public override void FillPath(FillingRule rule, bool close) => PaintPath(true, false, rule, close);
    public override void FillStrokePath(FillingRule rule, bool close) => PaintPath(true, true, rule, close);
    private PathFloatPrecision.Point SourcePoint(PdfPoint point, SKPoint candidate) => new(point.X, point.Y, candidate);
    private void RememberImplicitClosure()
    {
        if (_subpathHasSegments && _subpathStart.HasValue)
            _precision!.ImplicitClosure(SourcePoint(CurrentPosition, _currentCandidate), SourcePoint(_subpathStart.Value, _startCandidate));
    }
    public override void MoveTo(double x, double y)
    {
        Command(); RememberImplicitClosure();
        _precision ??= new PathFloatPrecision(CurrentTransformationMatrix, _height);
        _currentCandidate = _startCandidate = new SKPoint(Float(x), Float(y));
        _path.MoveTo(_currentCandidate); _subpathStart = CurrentPosition = new PdfPoint(x, y); _subpathHasSegments = false;
    }
    public override void LineTo(double x, double y)
    {
        if (_subpathStart is null) throw Unsupported("PATH_WITHOUT_MOVE"); Command();
        var candidate = new SKPoint(Float(x), Float(y));
        _precision!.Primitive(SourcePoint(CurrentPosition, _currentCandidate), new PathFloatPrecision.Point(x, y, candidate));
        _path.LineTo(candidate); _currentCandidate = candidate; _hasSegments = _subpathHasSegments = true; CurrentPosition = new PdfPoint(x, y);
    }
    public override void BezierCurveTo(double x2, double y2, double x3, double y3) => BezierCurveTo(CurrentPosition.X, CurrentPosition.Y, x2, y2, x3, y3);
    public override void BezierCurveTo(double x1, double y1, double x2, double y2, double x3, double y3)
    {
        if (_subpathStart is null) throw Unsupported("PATH_WITHOUT_MOVE"); Command();
        var first = new SKPoint(Float(x1), Float(y1)); var second = new SKPoint(Float(x2), Float(y2)); var end = new SKPoint(Float(x3), Float(y3));
        _precision!.Primitive(SourcePoint(CurrentPosition, _currentCandidate), new PathFloatPrecision.Point(x1, y1, first),
            new PathFloatPrecision.Point(x2, y2, second), new PathFloatPrecision.Point(x3, y3, end));
        _path.CubicTo(first, second, end); _currentCandidate = end; _hasSegments = _subpathHasSegments = true; CurrentPosition = new PdfPoint(x3, y3);
    }
    public override void Rectangle(double x, double y, double width, double height)
    {
        // Preserve signed re direction, including holes; SKRect normalization would change winding.
        MoveTo(x, y); LineTo(x + width, y); LineTo(x + width, y + height); LineTo(x, y + height); ClosePath();
        // Include addition error before float casting, e.g. x=2^53, width=1.
        var bx = PathFloatPrecision.Binary.From(x); var by = PathFloatPrecision.Binary.From(y);
        var right = bx + PathFloatPrecision.Binary.From(width); var top = by + PathFloatPrecision.Binary.From(height);
        _precision!.Primitive(new PathFloatPrecision.Point(bx, by, new SKPoint(Float(x), Float(y))),
            new PathFloatPrecision.Point(right, by, new SKPoint(Float(x + width), Float(y))),
            new PathFloatPrecision.Point(right, top, new SKPoint(Float(x + width), Float(y + height))),
            new PathFloatPrecision.Point(bx, top, new SKPoint(Float(x), Float(y + height))));
    }
    public override void EndPath()
    {
        if (_pendingClipRule is { } rule)
        {
            if (_activeClips.Count >= _limits.MaxStackDepth) throw new InvalidDataException("PDF clip stack limit exceeded.");
            if (!_hasSegments || _hasClosedSingleton) throw Unsupported("CLIP_PROFILE: drawable nondegenerate path required.");
            RememberImplicitClosure();
            if (!_precision!.Accept(true, false, 0, 0)) throw Unsupported("CLIP_FLOAT_PRECISION");
            _activeClips.Add(CaptureClip(rule)); _pendingClipRule = null;
        }
        _path.Reset(); _subpathStart = null; _precision = null; _hasSegments = _subpathHasSegments = _hasClosedSingleton = false;
    }

    private ClipShape CaptureClip(FillingRule rule)
    {
        const double units = 25.4 / 72;
        var matrix = Matrix(CurrentTransformationMatrix);
        using var iterator = _path.CreateRawIterator(); var points = new SKPoint[4]; var data = new StringBuilder();
        void Add(string command, params SKPoint[] values)
        {
            if (data.Length != 0) data.Append(' ');
            data.Append(command);
            foreach (var value in values)
            {
                var page = matrix.MapPoint(value);
                data.Append(' ').Append(OfdNumericFormat.Plain(page.X * units));
                data.Append(' ').Append(OfdNumericFormat.Plain(page.Y * units));
            }
            if (data.Length > _limits.MaxContentBytesPerPage) throw new InvalidDataException("PDF clip geometry limit exceeded.");
        }
        for (var verb = iterator.Next(points); verb != SKPathVerb.Done; verb = iterator.Next(points))
        {
            _token.ThrowIfCancellationRequested();
            switch (verb)
            {
                case SKPathVerb.Move: Add("M", points[0]); break;
                case SKPathVerb.Line: Add("L", points[1]); break;
                case SKPathVerb.Cubic: Add("B", points[1], points[2], points[3]); break;
                case SKPathVerb.Close: Add("C"); break;
                default: throw Unsupported("CLIP_PROFILE: path verb outside bounded profile.");
            }
        }
        return new ClipShape(data.ToString(), rule == FillingRule.EvenOdd);
    }
    public override void ClosePath()
    {
        if (_path.VerbCount == 0) return; Command();
        if (_subpathHasSegments && _subpathStart.HasValue)
            _precision!.Primitive(SourcePoint(CurrentPosition, _currentCandidate), SourcePoint(_subpathStart.Value, _startCandidate));
        _path.Close(); if (!_subpathHasSegments) _hasClosedSingleton = true;
        if (_subpathStart.HasValue) { CurrentPosition = _subpathStart.Value; _currentCandidate = _startCandidate; }
    }
    public override void ModifyClippingIntersect(FillingRule rule)
    {
        if (_pendingClipRule is not null) throw Unsupported("CLIP_PROFILE: repeated clip before path completion.");
        _pendingClipRule = rule;
    }
    protected override void ClipToRectangle(PdfRectangle bounds, FillingRule rule) => throw Unsupported("CLIP");
    protected override void RenderXObjectImage(XObjectContentRecord image)
    {
        RequireEvents(1);
        var remaining = _remainingImageBytes - _imageBytes;
        if (remaining <= 0) throw new InvalidDataException("PDF native image byte limit exceeded.");
        var element = _context.DecodeImageXObject(image.Stream, _limits, remaining, _token);
        var original = CurrentTransformationMatrix;
        var matrix = Matrix(original);
        const double units = 25.4 / 72;
        var points = new[] { new SKPoint(0, 0), new SKPoint(0, 1), new SKPoint(1, 0), new SKPoint(1, 1) }.Select(matrix.MapPoint).ToArray();
        var left = points.Min(point => (double)point.X) * units;
        var right = points.Max(point => (double)point.X) * units;
        var top = points.Min(point => (double)point.Y) * units;
        var bottom = points.Max(point => (double)point.Y) * units;
        if (!OfdNumericFormat.WriterPositive(right - left) || !OfdNumericFormat.WriterPositive(bottom - top))
            throw Unsupported("IMAGE_GEOMETRY: mapped image has no writable area.");
        element.XMillimeters = left; element.YMillimeters = top;
        element.WidthMillimeters = right - left; element.HeightMillimeters = bottom - top;
        element.Transform = new[] { (double)matrix.ScaleX * units, (double)matrix.SkewY * units,
            -(double)matrix.SkewX * units, -(double)matrix.ScaleY * units,
            ((double)matrix.SkewX + matrix.TransX) * units - left,
            ((double)matrix.ScaleY + matrix.TransY) * units - top };
        foreach (var u in new[] { 0d, 1d }) foreach (var v in new[] { 0d, 1d })
        {
            // PDF image samples run top-to-bottom; compare all four actual
            // emitted corners with the original double PDF CTM.
            var expectedX = (original.A * u + original.C * (1 - v) + original.E) * units;
            var expectedY = (_height - original.B * u - original.D * (1 - v) - original.F) * units;
            var writtenX = OfdNumericFormat.WriterValue(left) + element.Transform[0] * u + element.Transform[2] * v + element.Transform[4];
            var writtenY = OfdNumericFormat.WriterValue(top) + element.Transform[1] * u + element.Transform[3] * v + element.Transform[5];
            if (Math.Abs(expectedX - writtenX) > 0.01 * units || Math.Abs(expectedY - writtenY) > 0.01 * units)
                throw Unsupported("IMAGE_WRITER_PRECISION: mapped image corners exceed bounded profile.");
        }
        _imageBytes = checked(_imageBytes + element.Data.Length);
        _paintOrder.Add(element); _imageClips.Add(_activeClips.ToArray());
    }
    protected override void RenderInlineImage(InlineImage image) => throw Unsupported("INLINE_IMAGE");
    public override void BeginMarkedContent(NameToken name, NameToken? propertyDictionaryName, DictionaryToken? properties) => throw Unsupported("MARKED_CONTENT");
    public override void EndMarkedContent() => throw Unsupported("MARKED_CONTENT");
    public override void PaintShading(NameToken name) => throw Unsupported("SHADING");
    public void Dispose() => _path.Dispose();
}
