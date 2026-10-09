using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using Ofdrw.Net.Graphics.SkiaSharp;
using Ofdrw.Net.Core.Models;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Geometry;
using UglyToad.PdfPig.Graphics;
using UglyToad.PdfPig.Graphics.Core;
using UglyToad.PdfPig.Graphics.Operations;
using UglyToad.PdfPig.Parser;
using UglyToad.PdfPig.PdfFonts;
using UglyToad.PdfPig.Tokens;
using SkiaSharp;

namespace Ofdrw.Net.Converter.Pdf.Vector;

internal sealed class VectorStreamProcessor : BaseStreamProcessor<List<SkiaDrawEvent>>, IDisposable
{
    private readonly VectorPageContext _context; private readonly double _height;
    private readonly Dictionary<string, VectorFont> _fonts; private readonly PdfVectorToOfdOptions _limits;
    private readonly CancellationToken _token; private readonly SKPath _path = new(); private readonly List<SkiaDrawEvent> _events = new();
    private readonly StringBuilder _text = new(); private readonly List<double> _positions = new();
    private TransformationMatrix _textMatrix; private VectorFont? _textFont; private double _textSize;
    private bool _hasSegments, _subpathHasSegments, _hasClosedSingleton;
    private PathFloatPrecision? _precision;
    private SKPoint _currentCandidate, _startCandidate;
    private SKColor _textColor; private bool _inText; private int _commands, _characters; private PdfPoint? _subpathStart;

    internal VectorStreamProcessor(VectorPageContext context, IPageContentParser parser, double width, double height,
        Dictionary<string, VectorFont> fonts, PdfVectorToOfdOptions limits, CancellationToken token, int number)
        : base(number, context.Resources, context.Scanner, parser, context.Filters, new CropBox(new PdfRectangle(0, 0, width, height)),
            UserSpaceUnit.Default, new PageRotationDegrees(0), TransformationMatrix.Identity, context.Parsing)
    { _context = context; _height = height; _fonts = fonts; _limits = limits; _token = token; }

    public override List<SkiaDrawEvent> Process(int number, IReadOnlyList<IGraphicsStateOperation> operations)
    {
        foreach (var operation in operations)
        {
            _token.ThrowIfCancellationRequested();
            if (_path.VerbCount != 0 && operation.Operator is "cm" or "q" or "Q") throw Unsupported("PATH_STATE_CHANGE");
            operation.Run(this);
            if (operation.Operator is "Tj" or "TJ") FlushText();
        }
        if (_inText || StackSize != 1 || _path.VerbCount != 0) throw new InvalidDataException("PDF has an unterminated text, graphics state or path.");
        return _events;
    }
    public override void PushState()
    { if (StackSize >= _limits.MaxStackDepth) throw new InvalidDataException("PDF graphics stack limit exceeded."); base.PushState(); }
    public override void PopState()
    { if (StackSize <= 1) throw new InvalidDataException("PDF graphics stack underflow."); base.PopState(); }
    public override void BeginText()
    { if (_inText) throw Unsupported("NESTED_TEXT"); _inText = true; base.BeginText(); }
    public override void EndText()
    { if (!_inText) throw Unsupported("TEXT_UNDERFLOW"); _inText = false; base.EndText(); }
    private NotSupportedException Unsupported(string reason) => _context.Unsupported(reason);
    private void Add(SkiaDrawEvent value)
    { if (_events.Count >= _limits.MaxEventsPerPage) throw new InvalidDataException("PDF draw event limit exceeded."); _events.Add(value); }
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
        if (stroke && (state.LineWidth <= 0 || state.CapStyle != LineCapStyle.Butt || state.JoinStyle != LineJoinStyle.Miter || state.MiterLimit != 10))
            throw Unsupported("STROKE_PROFILE: positive width, butt cap, miter join and miter limit 10 required.");
        return new SKPaint { Color = Color(stroke), Style = stroke ? SKPaintStyle.Stroke : SKPaintStyle.Fill,
            StrokeWidth = Float(state.LineWidth), StrokeMiter = 10 };
    }
    private void PaintPath(bool fill, bool stroke, FillingRule rule, bool close)
    {
        if (close) ClosePath();
        if (_path.VerbCount == 0) { EndPath(); return; }
        // A closed singleton fill can paint a device pixel; do not silently elide it.
        if (fill && _hasClosedSingleton) throw Unsupported("DEGENERATE_POINT_FILL");
        if (!_hasSegments)
        {
            // Validate the supported paint profile before treating a source no-op as such.
            if (fill) { using var paint = Paint(false); }
            if (stroke) { using var paint = Paint(true); }
            EndPath(); return;
        }
        if (fill) RememberImplicitClosure();
        var matrix = Matrix(CurrentTransformationMatrix); // Keep existing singular diagnostics.
        using var fillPaint = fill ? Paint(false) : null;
        using var strokePaint = stroke ? Paint(true) : null;
        if (!_precision!.Accept(fill, stroke, GetCurrentState().LineWidth, strokePaint?.StrokeWidth ?? 0))
            throw Unsupported("PATH_FLOAT_PRECISION: added path/control/CTM/stroke conversion error exceeds bounded profile.");
        _path.FillType = rule == FillingRule.EvenOdd ? SKPathFillType.EvenOdd : SKPathFillType.Winding;
        if (fill) Add(SkiaDrawEvent.Path(_path, fillPaint!, matrix, _limits.MaxPathCommandsPerPage));
        if (stroke) Add(SkiaDrawEvent.Path(_path, strokePaint!, matrix, _limits.MaxPathCommandsPerPage));
        EndPath();
    }
    public override void RenderGlyph(IFont font, CurrentGraphicsState state, double fontSize, double pointSize, int code, string unicode,
        long offset, in TransformationMatrix rendering, in TransformationMatrix text, in TransformationMatrix ctm, CharacterBoundingBox bounds)
    {
        _token.ThrowIfCancellationRequested();
        if (!_inText || !font.TryGetUnicode(code, out var original) || original != unicode || string.IsNullOrEmpty(original) || original.Length != 1 ||
            char.IsSurrogate(original[0]) || char.IsControl(original[0]) || font.IsVertical)
            throw Unsupported("UNICODE_MAPPING: one original BMP character per glyph required.");
        var category = CharUnicodeInfo.GetUnicodeCategory(original[0]);
        if (category is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark)
            throw Unsupported("COMBINING_SHAPING");
        if (!(fontSize > 0) || state.FontState.HorizontalScaling != 100 || state.FontState.Rise != 0) throw Unsupported("TEXT_STATE");
        if (!_fonts.TryGetValue(state.FontState.FontName.Data, out var binding)) throw Unsupported("FONT_RESOURCE");
        var glyph = binding.Face.GetGlyph(original[0]);
        if (glyph == 0 || glyph != code) throw Unsupported("SOURCE_GLYPH_CMAP_MISMATCH: font=" + state.FontState.FontName.Data + " code=" + code);
        if (_characters >= _limits.MaxTextCharactersPerPage) throw new InvalidDataException("PDF original text character limit exceeded.");
        _characters++;
        var matrix = text.Multiply(ctm); var color = Color(false);
        if (_text.Length == 0)
        {
            _textMatrix = matrix; _textFont = binding; _textSize = fontSize; _textColor = color; _positions.Add(0);
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
            _positions.Add(x);
        }
        _text.Append(original);
    }
    private void FlushText()
    {
        if (_text.Length == 0) return;
        var original = _text.ToString();
        if (string.IsNullOrWhiteSpace(original) || StringInfo.ParseCombiningCharacters(original).Length != _positions.Count) throw Unsupported("WHITESPACE_ONLY_OR_SHAPED_RUN");
        var advances = new float[_positions.Count - 1];
        for (var index = 0; index < advances.Length; index++) advances[index] = Float(_positions[index + 1] - _positions[index]);
        using var font = new SKFont(_textFont!.Face, Float(_textSize)); using var paint = new SKPaint { Color = _textColor, Style = SKPaintStyle.Fill };
        Add(SkiaDrawEvent.Text(original, new SKPoint(0, 0), font, paint, _textFont.Resource.Id!, advances,
            Matrix(_textMatrix, true), _limits.MaxTextCharactersPerPage, _limits.MaxFontBytes));
        _text.Clear(); _positions.Clear(); _textFont = null;
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
    public override void EndPath() { _path.Reset(); _subpathStart = null; _precision = null; _hasSegments = _subpathHasSegments = _hasClosedSingleton = false; }
    public override void ClosePath()
    {
        if (_path.VerbCount == 0) return; Command();
        if (_subpathHasSegments && _subpathStart.HasValue)
            _precision!.Primitive(SourcePoint(CurrentPosition, _currentCandidate), SourcePoint(_subpathStart.Value, _startCandidate));
        _path.Close(); if (!_subpathHasSegments) _hasClosedSingleton = true;
        if (_subpathStart.HasValue) { CurrentPosition = _subpathStart.Value; _currentCandidate = _startCandidate; }
    }
    public override void ModifyClippingIntersect(FillingRule rule) => throw Unsupported("CLIP");
    protected override void ClipToRectangle(PdfRectangle bounds, FillingRule rule) => throw Unsupported("CLIP");
    protected override void RenderXObjectImage(XObjectContentRecord image) => throw Unsupported("IMAGE");
    protected override void RenderInlineImage(InlineImage image) => throw Unsupported("INLINE_IMAGE");
    public override void BeginMarkedContent(NameToken name, NameToken? propertyDictionaryName, DictionaryToken? properties) => throw Unsupported("MARKED_CONTENT");
    public override void EndMarkedContent() => throw Unsupported("MARKED_CONTENT");
    public override void PaintShading(NameToken name) => throw Unsupported("SHADING");
    public void Dispose() => _path.Dispose();
}
