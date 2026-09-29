using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Internal.Flow;

namespace Ofdrw.Net.Layout;

/// <summary>A reusable description of flowing OFD content. Each render creates an independent package.</summary>
public sealed class FlowDocument
{
    /// <summary>Page, margin, font and resource limits for this document.</summary>
    public FlowDocumentOptions Options { get; } = new();

    /// <summary>Ordered flow blocks. Paragraph is the supported block in this release.</summary>
    public IList<FlowBlock> Blocks { get; } = new List<FlowBlock>();

    /// <summary>Measures and paginates the blocks into native OFD text objects.</summary>
    public OfdDocumentPackage Render(CancellationToken cancellationToken = default) =>
        new FlowDocumentRenderer(Options, Blocks, cancellationToken).Render();
}

/// <summary>Base of the built-in flow blocks.</summary>
public abstract class FlowBlock
{
    internal FlowBlock() { }
}

/// <summary>A paragraph composed of independently styled spans.</summary>
public sealed class Paragraph : FlowBlock
{
    /// <summary>Creates an empty paragraph.</summary>
    public Paragraph() { }

    /// <summary>Creates a paragraph with one unstyled span.</summary>
    public Paragraph(string text) => Spans.Add(new Span(text));

    /// <summary>Inline text in reading order.</summary>
    public IList<Span> Spans { get; } = new List<Span>();
    public ParagraphAlignment Alignment { get; set; }
    public double SpaceBeforeMillimeters { get; set; }
    public double SpaceAfterMillimeters { get; set; }
    public double LeftIndentMillimeters { get; set; }
    public double RightIndentMillimeters { get; set; }
    public double FirstLineIndentMillimeters { get; set; }
    public double? MinimumLineHeightMillimeters { get; set; }
    public bool PageBreakBefore { get; set; }
}

/// <summary>Supported horizontal paragraph alignment.</summary>
public enum ParagraphAlignment { Left, Center, Right }

/// <summary>Text and styling local to this span; omitted family, size and color use document defaults.</summary>
public sealed class Span
{
    public Span(string text) => Text = text ?? throw new ArgumentNullException(nameof(text));
    public string Text { get; set; }
    public string? FontFamily { get; set; }
    public double? FontSizeMillimeters { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public OfdColor? Color { get; set; }
}

/// <summary>Geometry in millimeters and per-render resource limits.</summary>
public sealed class FlowDocumentOptions
{
    public double PageWidthMillimeters { get; set; } = 210;
    public double PageHeightMillimeters { get; set; } = 297;
    public double MarginLeftMillimeters { get; set; } = 25.4;
    public double MarginRightMillimeters { get; set; } = 25.4;
    public double MarginTopMillimeters { get; set; } = 25.4;
    public double MarginBottomMillimeters { get; set; } = 25.4;
    public string DefaultFontFamily { get; set; } = "SimSun";
    public double DefaultFontSizeMillimeters { get; set; } = 10.5 * 25.4 / 72;
    public OfdColor DefaultColor { get; set; } = OfdColor.Black;
    public int MaxPageCount { get; set; } = 10_000;
    public int MaxCharacters { get; set; } = 1_000_000;
    public int MaxTextElements { get; set; } = 1_000_000;
}

internal sealed class FlowDocumentRenderer : IFlowFontMetrics
{
    private readonly FlowDocumentOptions _options;
    private readonly IList<FlowBlock> _blocks;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<string, OfdFontResource> _fonts = new(StringComparer.Ordinal);
    private readonly OfdDocumentPackage _package = new();
    private OfdPage? _page;
    private double _y;
    private bool _hasBodyContent;
    private int _textElements;
    private int _pendingPageBreaks;
    private double _pendingSpaceAfter;

    internal FlowDocumentRenderer(FlowDocumentOptions options, IList<FlowBlock> blocks, CancellationToken cancellationToken)
    { _options = options; _blocks = blocks; _cancellationToken = cancellationToken; }

    internal OfdDocumentPackage Render()
    {
        ValidateOptions();
        _package.Options.DefaultPageWidthMillimeters = _options.PageWidthMillimeters;
        _package.Options.DefaultPageHeightMillimeters = _options.PageHeightMillimeters;
        var paragraphs = _blocks.ToArray();
        var characterCount = 0L;
        foreach (var block in paragraphs)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (block is not Paragraph paragraph) throw new NotSupportedException("Unsupported flow block.");
            var spans = paragraph.Spans.ToArray();
            foreach (var span in spans)
            {
                if (span is null || span.Text is null) throw new ArgumentException("A paragraph contains a null span or text.");
                characterCount += span.Text.Length;
                if (characterCount > _options.MaxCharacters)
                    throw new InvalidOperationException("Flow text exceeds MaxCharacters.");
            }
            RenderParagraph(paragraph, spans);
        }
        if (_page is null) StartPage();
        return _package;
    }

    private void RenderParagraph(Paragraph paragraph, IReadOnlyList<Span> spans)
    {
        ValidateLength(paragraph.SpaceBeforeMillimeters, nameof(paragraph.SpaceBeforeMillimeters));
        ValidateLength(paragraph.SpaceAfterMillimeters, nameof(paragraph.SpaceAfterMillimeters));
        ValidateLength(paragraph.LeftIndentMillimeters, nameof(paragraph.LeftIndentMillimeters));
        ValidateLength(paragraph.RightIndentMillimeters, nameof(paragraph.RightIndentMillimeters));
        ValidateLength(paragraph.FirstLineIndentMillimeters, nameof(paragraph.FirstLineIndentMillimeters));
        if (paragraph.MinimumLineHeightMillimeters is { } lineHeight)
            ValidateLength(lineHeight, nameof(paragraph.MinimumLineHeightMillimeters));
        if (!Enum.IsDefined(typeof(ParagraphAlignment), paragraph.Alignment))
            throw new ArgumentOutOfRangeException(nameof(paragraph.Alignment));
        var width = _options.PageWidthMillimeters - _options.MarginLeftMillimeters -
            _options.MarginRightMillimeters - paragraph.LeftIndentMillimeters - paragraph.RightIndentMillimeters;
        if (width <= 0 || paragraph.FirstLineIndentMillimeters >= width)
            throw new ArgumentException("Paragraph indents leave no usable line width.");
        if (_page is null) StartPage();
        var hadPendingBreaks = _pendingPageBreaks > 0;
        ApplyPendingPageBreaks();
        if (paragraph.PageBreakBefore && (_hasBodyContent || hadPendingBreaks)) StartPage();
        var inlines = new List<FlowInline>();
        foreach (var span in spans)
        {
            var size = span.FontSizeMillimeters ?? _options.DefaultFontSizeMillimeters;
            if (double.IsNaN(size) || double.IsInfinity(size) || size <= 0)
                throw new ArgumentOutOfRangeException(nameof(span.FontSizeMillimeters));
            var family = span.FontFamily ?? _options.DefaultFontFamily;
            if (string.IsNullOrWhiteSpace(family)) throw new ArgumentException("Font family cannot be empty.");
            inlines.Add(new FlowInline
            {
                Text = span.Text,
                Style = new FlowTextStyle { FontFamily = family.Trim(), FontSizeMillimeters = size,
                    Bold = span.Bold, Italic = span.Italic, Source = span.Color ?? _options.DefaultColor }
            });
        }
        var fallback = new FlowTextStyle { FontFamily = _options.DefaultFontFamily,
            FontSizeMillimeters = _options.DefaultFontSizeMillimeters, Source = _options.DefaultColor };
        var format = new FlowParagraphFormat
        {
            Alignment = (FlowAlignment)paragraph.Alignment,
            FirstLineIndentMillimeters = paragraph.FirstLineIndentMillimeters,
            MinimumLineHeightMillimeters = paragraph.MinimumLineHeightMillimeters ?? 0
        };
        var firstLine = true;
        foreach (var line in FlowParagraphLayout.Layout(inlines, format, width, this, fallback, _cancellationToken))
        {
            if (line.PageBreak) { _pendingPageBreaks++; continue; }
            ApplyPendingPageBreaks();
            if (firstLine)
            {
                PlaceFirstLine(line.Height, _pendingSpaceAfter + paragraph.SpaceBeforeMillimeters);
                _pendingSpaceAfter = 0;
                firstLine = false;
            }
            else Place(line.Height);
            DrawLine(line, _options.MarginLeftMillimeters + paragraph.LeftIndentMillimeters, width);
            _y += line.Height;
            _hasBodyContent = true;
        }
        // Carry the gap to the next block, without materializing a trailing empty page.
        _pendingSpaceAfter += paragraph.SpaceAfterMillimeters;
    }

    private void DrawLine(FlowLine line, double left, double width)
    {
        var x = FlowParagraphLayout.Align(left + line.Indent, width - line.Indent,
            line.Glyphs.Sum(g => g.Width), line.Alignment);
        for (var i = 0; i < line.Glyphs.Count;)
        {
            var style = line.Glyphs[i].Style;
            var start = i++;
            while (i < line.Glyphs.Count && ReferenceEquals(line.Glyphs[i].Style, style)) i++;
            if (++_textElements > _options.MaxTextElements)
                throw new InvalidOperationException("Flow text exceeds MaxTextElements.");
            var group = line.Glyphs.GetRange(start, i - start);
            var font = GetFont(style);
            var text = string.Concat(group.Select(g => g.Text));
            var element = new OfdTextElement
            {
                LayerType = "Body", XMillimeters = x, YMillimeters = _y,
                WidthMillimeters = Math.Max(group.Sum(g => g.Width), 0.1), HeightMillimeters = line.Height,
                FontName = font.FontName, FontResourceId = font.Id,
                FontSizeMillimeters = style.FontSizeMillimeters,
                Weight = style.Bold ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight,
                Italic = style.Italic, FillColor = (OfdColor)style.Source!, Text = text
            };
            element.Runs.Add(new OfdTextRun { Text = text, YMillimeters = line.BaselineMillimeters,
                DeltaX = group.Count > 1 ? string.Join(" ", group.Take(group.Count - 1)
                    .Select(g => g.Width.ToString("0.######", CultureInfo.InvariantCulture))) : null });
            _page!.Elements.Add(element);
            x += group.Sum(g => g.Width);
        }
    }

    private OfdFontResource GetFont(FlowTextStyle style)
    {
        var key = style.FontFamily + "|" + style.Bold + "|" + style.Italic;
        if (_fonts.TryGetValue(key, out var resource)) return resource;
        resource = new OfdFontResource { Id = "F" + (_fonts.Count + 1), FontName = style.FontFamily,
            FamilyName = style.FontFamily, Charset = "unicode", Bold = style.Bold, Italic = style.Italic };
        _fonts.Add(key, resource);
        _package.Fonts.Add(resource);
        return resource;
    }

    private double ContentTop => _options.MarginTopMillimeters;
    private double ContentBottom => _options.PageHeightMillimeters - _options.MarginBottomMillimeters;

    private void Place(double height)
    {
        if (_page is null) StartPage();
        if (FlowPagination.NeedsNewPage(_y, height, ContentTop, ContentBottom)) StartPage();
        if (FlowPagination.NeedsNewPage(_y, height, ContentTop, ContentBottom))
            throw new InvalidOperationException("Flow line cannot fit on a fresh page.");
    }

    private void PlaceFirstLine(double height, double gap)
    {
        var usableHeight = ContentBottom - ContentTop;
        if (height > usableHeight) throw new InvalidOperationException("A flow line is taller than the usable page area.");
        if (_y + gap + height > ContentBottom + 0.000001d && (_hasBodyContent || _y > ContentTop))
            StartPage();
        // Exceptionally large paragraph spacing yields to the first complete line.
        _y += Math.Min(gap, usableHeight - height);
        Place(height);
    }

    private void StartPage()
    {
        if (_package.Pages.Count >= _options.MaxPageCount)
            throw new InvalidOperationException("Flow document exceeds MaxPageCount.");
        _page = new OfdPage { Index = _package.Pages.Count,
            WidthMillimeters = _options.PageWidthMillimeters, HeightMillimeters = _options.PageHeightMillimeters };
        _package.Pages.Add(_page);
        _y = ContentTop;
        _hasBodyContent = false;
    }

    private void ApplyPendingPageBreaks()
    {
        while (_pendingPageBreaks > 0)
        {
            StartPage();
            _pendingPageBreaks--;
        }
    }

    private void ValidateOptions()
    {
        ValidatePositive(_options.PageWidthMillimeters, nameof(_options.PageWidthMillimeters));
        ValidatePositive(_options.PageHeightMillimeters, nameof(_options.PageHeightMillimeters));
        ValidatePositive(_options.DefaultFontSizeMillimeters, nameof(_options.DefaultFontSizeMillimeters));
        ValidateLength(_options.MarginLeftMillimeters, nameof(_options.MarginLeftMillimeters));
        ValidateLength(_options.MarginRightMillimeters, nameof(_options.MarginRightMillimeters));
        ValidateLength(_options.MarginTopMillimeters, nameof(_options.MarginTopMillimeters));
        ValidateLength(_options.MarginBottomMillimeters, nameof(_options.MarginBottomMillimeters));
        if (_options.PageWidthMillimeters <= _options.MarginLeftMillimeters + _options.MarginRightMillimeters ||
            _options.PageHeightMillimeters <= _options.MarginTopMillimeters + _options.MarginBottomMillimeters)
            throw new ArgumentException("Margins leave no usable page area.");
        if (string.IsNullOrWhiteSpace(_options.DefaultFontFamily)) throw new ArgumentException("Default font family is required.");
        if (_options.DefaultColor is null) throw new ArgumentException("Default color is required.");
        if (_options.MaxPageCount < 1 || _options.MaxCharacters < 1 || _options.MaxTextElements < 1)
            throw new ArgumentOutOfRangeException(nameof(_options.MaxPageCount));
    }

    private static void ValidatePositive(double value, string name)
    {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) throw new ArgumentOutOfRangeException(name);
    }
    private static void ValidateLength(double value, string name)
    {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) throw new ArgumentOutOfRangeException(name);
    }

    public double AdvanceMillimeters(string grapheme, FlowTextStyle style) =>
        FlowLatinMetrics.AdvanceMillimeters(grapheme, style);
}
