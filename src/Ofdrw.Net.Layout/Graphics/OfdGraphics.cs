using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Graphics;

/// <summary>Appends native OFD path/text objects to an existing page. Coordinates and font sizes use millimeters from the physical page top-left; positive Y points down.</summary>
public sealed class OfdGraphics
{
    private readonly XNamespace Ns;
    private readonly OfdDocumentPackage _package;
    private readonly OfdPage _page;
    private readonly OfdGraphicsOptions _limits;
    private readonly Stack<(OfdMatrix Matrix, List<(PathSnapshot Path, OfdMatrix Matrix)> Clips)> _states = new();
    private List<(PathSnapshot Path, OfdMatrix Matrix)> _clips = new();
    private long _textCharacters, _geometryCharacters;
    /// <summary>Creates a context for a page already in package.Pages. Does not own the package, fonts or output stream.</summary>
    public OfdGraphics(OfdDocumentPackage package, OfdPage page, OfdGraphicsOptions? options = null)
    {
        _package = package ?? throw new ArgumentNullException(nameof(package));
        Ns = XNamespace.Get(package.Options.Namespace);
        _page = page ?? throw new ArgumentNullException(nameof(page));
        if (!package.Pages.Contains(page)) throw new ArgumentException("Page must belong to the destination package.", nameof(page));
        GraphicsValidation.Finite(page.XMillimeters, page.YMillimeters, page.XMillimeters + page.WidthMillimeters, page.YMillimeters + page.HeightMillimeters);
        GraphicsValidation.Positive(page.WidthMillimeters); GraphicsValidation.Positive(page.HeightMillimeters);
        options ??= new OfdGraphicsOptions();
        if (options.MaxPageElements <= 0 || options.MaxPathCommands <= 0 || options.MaxTextCharacters <= 0 ||
            options.MaxGeometryCharacters <= 0 || options.MaxSavedStates <= 0 || options.MaxClipRegions <= 0)
            throw new ArgumentException("Graphics limits must be positive.", nameof(options));
        _limits = new OfdGraphicsOptions { MaxPageElements = options.MaxPageElements, MaxPathCommands = options.MaxPathCommands,
            MaxTextCharacters = options.MaxTextCharacters, MaxGeometryCharacters = options.MaxGeometryCharacters,
            MaxSavedStates = options.MaxSavedStates, MaxClipRegions = options.MaxClipRegions };
    }
    /// <summary>Current immutable user-to-page matrix.</summary>
    public OfdMatrix Transform { get; private set; } = OfdMatrix.Identity;
    /// <summary>Replaces the user-to-page matrix. Matrices singular in memory or at OFD writer precision, and derived overflow, fail before state changes.</summary>
    public void SetTransform(OfdMatrix matrix)
    {
        ValidateTransform(matrix);
        Transform = matrix;
    }
    private static void ValidateTransform(OfdMatrix matrix)
    {
        if (matrix is null) throw new ArgumentNullException(nameof(matrix));
        if (!OfdNumericFormat.Nonsingular(matrix.A, matrix.B, matrix.C, matrix.D))
            throw new ArgumentException("Graphics transform must have nonsingular serialized coefficients.", nameof(matrix));
    }
    /// <summary>Appends in user space: current = current * matrix; the appended matrix acts first.</summary>
    public void MultiplyTransform(OfdMatrix matrix)
    {
        // Reject a singular operand before floating multiplication can perturb
        // its coefficients into an accidentally nonsingular serialized result.
        ValidateTransform(matrix);
        SetTransform(Transform.Multiply(matrix));
    }
    /// <summary>Appends a user-space translation.</summary>
    public void Translate(double x, double y) => MultiplyTransform(OfdMatrix.Translation(x, y));
    /// <summary>Appends a user-space scale.</summary>
    public void Scale(double x, double y) => MultiplyTransform(OfdMatrix.Scale(x, y));
    /// <summary>Appends a clockwise rotation in degrees.</summary>
    public void Rotate(double degrees) => MultiplyTransform(OfdMatrix.Rotation(degrees));
    /// <summary>Saves the transform and frozen clips. Previously emitted objects are unaffected.</summary>
    public void Save()
    {
        if (_states.Count >= _limits.MaxSavedStates) throw new InvalidOperationException("Graphics state budget exceeded.");
        _states.Push((Transform, new List<(PathSnapshot, OfdMatrix)>(_clips)));
    }
    /// <summary>Restores and consumes the most recently saved state; an empty stack fails.</summary>
    public void Restore()
    {
        if (_states.Count == 0) throw new InvalidOperationException("There is no saved graphics state.");
        var state = _states.Pop(); Transform = state.Matrix; _clips = state.Clips;
    }
    /// <summary>Intersects the current clip with a snapshot transformed to page space now. Later transforms do not move it.</summary>
    public void IntersectClip(OfdGraphicsPath path, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (path is null) throw new ArgumentNullException(nameof(path));
        if (_clips.Count >= _limits.MaxClipRegions) throw new InvalidOperationException("Graphics clip budget exceeded.");
        var snapshot = path.Snapshot(_limits.MaxPathCommands, cancellationToken);
        foreach (var point in snapshot.Points) { cancellationToken.ThrowIfCancellationRequested(); PageTransform.TransformPoint(point.X, point.Y); }
        var length = snapshot.Data.Length;
        CheckGeometry(length);
        cancellationToken.ThrowIfCancellationRequested();
        _clips.Add((snapshot, PageTransform)); _geometryCharacters += length;
    }
    /// <summary>Clears current clips, without changing any saved state.</summary>
    public void ResetClip() => _clips = new();
    /// <summary>Draws a line with a solid pen.</summary>
    public void DrawLine(OfdPen pen, double x1, double y1, double x2, double y2, CancellationToken cancellationToken = default)
        => DrawPath(new OfdGraphicsPath().MoveTo(x1, y1).LineTo(x2, y2), pen ?? throw new ArgumentNullException(nameof(pen)), null, cancellationToken);
    /// <summary>Strokes a positive-size rectangle.</summary>
    public void DrawRectangle(OfdPen pen, double x, double y, double width, double height, CancellationToken cancellationToken = default)
        => DrawPath(new OfdGraphicsPath().AddRectangle(x, y, width, height), pen ?? throw new ArgumentNullException(nameof(pen)), null, cancellationToken);
    /// <summary>Fills a positive-size rectangle.</summary>
    public void FillRectangle(OfdBrush brush, double x, double y, double width, double height, CancellationToken cancellationToken = default)
        => DrawPath(new OfdGraphicsPath().AddRectangle(x, y, width, height), null, brush ?? throw new ArgumentNullException(nameof(brush)), cancellationToken);
    /// <summary>Draws a snapshot, optionally filled then stroked. At least one style is required. A failed draw appends nothing.</summary>
    public void DrawPath(OfdGraphicsPath path, OfdPen? pen = null, OfdBrush? brush = null, CancellationToken cancellationToken = default)
    {
        CheckPage(cancellationToken);
        if (path is null) throw new ArgumentNullException(nameof(path));
        if (pen is null && brush is null) throw new ArgumentException("A pen or brush is required.");
        var snapshot = path.Snapshot(_limits.MaxPathCommands, cancellationToken);
        var points = snapshot.Points.Select(point => PageTransform.TransformPoint(point.X, point.Y)).ToArray();
        // Control polygon bounds cover Bezier curves. Miter limit 10 bounds stroke overhang under the full affine transform.
        var halfWidth = (pen?.WidthMillimeters ?? 0) * 5;
        var dx = halfWidth * (Math.Abs(Transform.A) + Math.Abs(Transform.C));
        var dy = halfWidth * (Math.Abs(Transform.B) + Math.Abs(Transform.D));
        var x = points.Min(point => point.X) - dx; var y = points.Min(point => point.Y) - dy;
        var width = Math.Max(0.001, points.Max(point => point.X) + dx - x);
        var height = Math.Max(0.001, points.Max(point => point.Y) + dy - y);
        GraphicsValidation.Finite(x, y, width, height);
        var data = snapshot.Data;
        var element = new OfdPathElement { XMillimeters = x, YMillimeters = y, WidthMillimeters = width, HeightMillimeters = height,
            Transform = Rebase(PageTransform, x, y).ToArray(), AbbreviatedData = data,
            Stroke = pen is not null, Fill = brush is not null, LineWidthMillimeters = pen?.WidthMillimeters ?? 0.353,
            StrokeColor = pen?.Color ?? OfdColor.Black, FillColor = brush?.Color,
            SourceXml = new XElement(Ns + "PathObject", new XAttribute("Rule", Rule(snapshot.Rule)),
                new XAttribute("Cap", "Butt"), new XAttribute("Join", "Miter"), new XAttribute("MiterLimit", "10")).ToString(SaveOptions.DisableFormatting) };
        Commit(element, data.Length, 0, cancellationToken);
    }
    /// <summary>Fills a path snapshot according to its fill rule.</summary>
    public void FillPath(OfdBrush brush, OfdGraphicsPath path, CancellationToken cancellationToken = default)
        => DrawPath(path, null, brush ?? throw new ArgumentNullException(nameof(brush)), cancellationToken);
    /// <summary>Draws one baseline text run. No measurement or wrapping is performed. Optional advances must contain exactly grapheme-count minus one finite millimeter values.</summary>
    /// <remarks>Text uses the physical page as its Boundary viewport; baseline and glyph CTM remain exact. Off-page ink is clipped by the page. Font lookup/embedding is owned by the existing package writer/exporters.</remarks>
    public void DrawString(string text, OfdFont font, OfdBrush brush, double x, double baselineY,
        IReadOnlyList<double>? advances = null, CancellationToken cancellationToken = default)
    {
        CheckPage(cancellationToken);
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (font is null) throw new ArgumentNullException(nameof(font));
        if (brush is null) throw new ArgumentNullException(nameof(brush));
        GraphicsValidation.Finite(x, baselineY); PageTransform.TransformPoint(x, baselineY);
        if (text.Length == 0 || text.IndexOfAny(new[] { '\r', '\n', '\t' }) >= 0) throw new ArgumentException("Text must be a nonempty single baseline run.", nameof(text));
        if (text.Length > _limits.MaxTextCharacters - _textCharacters) throw new InvalidOperationException("Graphics text budget exceeded.");
        XmlConvert.VerifyXmlChars(text);
        var name = font.Name;
        if (font.ResourceId is not null)
        {
            var matches = _package.Fonts.Where(value => value.Id == font.ResourceId).ToArray();
            if (matches.Length != 1 || string.IsNullOrWhiteSpace(matches[0].FontName)) throw new ArgumentException("Font resource ID must identify exactly one named destination resource.", nameof(font));
            name = matches[0].FontName;
        }
        string? deltas = null;
        if (advances is not null)
        {
            var glyphs = OfdTextGeometry.Glyphs(text);
            if (advances.Count != glyphs.Count - 1) throw new ArgumentException("Advances must match grapheme count minus one.", nameof(advances));
            var values = new List<string>(advances.Count); var cursor = x;
            foreach (var advance in advances)
            {
                cancellationToken.ThrowIfCancellationRequested(); GraphicsValidation.Finite(advance);
                cursor += advance; PageTransform.TransformPoint(cursor, baselineY); values.Add(GraphicsValidation.Number(advance));
            }
            deltas = string.Join(" ", values);
        }
        var localBaseline = GraphicsValidation.WriterValue(font.SizeMillimeters);
        var element = new OfdTextElement { Text = text, FontName = name, FontResourceId = font.ResourceId,
            FontSizeMillimeters = font.SizeMillimeters, Weight = font.Weight, Italic = font.Italic, FillColor = brush.Color,
            XMillimeters = _page.XMillimeters, YMillimeters = _page.YMillimeters,
            WidthMillimeters = _page.WidthMillimeters, HeightMillimeters = _page.HeightMillimeters,
            Transform = Rebase(PageTransform.Multiply(OfdMatrix.Translation(x, baselineY - localBaseline)),
                _page.XMillimeters, _page.YMillimeters).ToArray() };
        // Normalize the run's baseline so the existing writer can compose its
        // generated name-only italic factor about this anchor, under any user CTM.
        element.Runs.Add(new OfdTextRun { Text = text, XMillimeters = 0, YMillimeters = localBaseline, DeltaX = deltas });
        Commit(element, deltas?.Length ?? 0, text.Length, cancellationToken);
    }
    private OfdMatrix PageTransform => OfdMatrix.Translation(_page.XMillimeters, _page.YMillimeters).Multiply(Transform);
    private void CheckPage(CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (!_package.Pages.Contains(_page)) throw new InvalidOperationException("Target page was removed from the package.");
        GraphicsValidation.Finite(_page.XMillimeters, _page.YMillimeters, _page.XMillimeters + _page.WidthMillimeters, _page.YMillimeters + _page.HeightMillimeters);
        GraphicsValidation.Positive(_page.WidthMillimeters); GraphicsValidation.Positive(_page.HeightMillimeters);
        if (_page.Elements.Count >= _limits.MaxPageElements) throw new InvalidOperationException("Graphics page element budget exceeded.");
    }
    private void CheckGeometry(long length)
    { if (length > _limits.MaxGeometryCharacters - _geometryCharacters) throw new InvalidOperationException("Graphics geometry budget exceeded."); }
    private void Commit(OfdElement element, long geometryLength, long textLength, CancellationToken token)
    {
        if (_clips.Count > 0)
        {
            var xml = new XElement(Ns + "Clips");
            foreach (var clip in _clips)
            {
                token.ThrowIfCancellationRequested();
                xml.Add(new XElement(Ns + "Clip", new XElement(Ns + "Area", new XElement(Ns + "Path",
                    new XAttribute("CTM", string.Join(" ", Rebase(clip.Matrix, element.XMillimeters, element.YMillimeters).ToArray().Select(GraphicsValidation.Number))),
                    new XAttribute("Rule", Rule(clip.Path.Rule)), new XElement(Ns + "AbbreviatedData", clip.Path.Data)))));
            }
            element.ClippingXml = xml.ToString(SaveOptions.DisableFormatting); geometryLength += element.ClippingXml.Length;
        }
        CheckGeometry(geometryLength); CheckPage(token);
        _page.Elements.Add(element); _geometryCharacters += geometryLength; _textCharacters += textLength;
    }
    private static OfdMatrix Rebase(OfdMatrix matrix, double x, double y) => new(matrix.A, matrix.B, matrix.C, matrix.D, matrix.E - x, matrix.F - y);
    private static string Rule(OfdFillRule rule) => rule == OfdFillRule.EvenOdd ? "Even-Odd" : "NonZero";
}
