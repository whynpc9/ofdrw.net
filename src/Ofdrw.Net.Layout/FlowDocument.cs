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

    /// <summary>Ordered Paragraph and Table blocks. Do not mutate the description during rendering.</summary>
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
    /// <summary>Horizontal alignment of each laid-out paragraph line.</summary>
    public ParagraphAlignment Alignment { get; set; }
    /// <summary>Vertical space preceding the paragraph, in millimeters.</summary>
    public double SpaceBeforeMillimeters { get; set; }
    /// <summary>Vertical space following the paragraph, in millimeters.</summary>
    public double SpaceAfterMillimeters { get; set; }
    /// <summary>Left paragraph indent measured inside the content width.</summary>
    public double LeftIndentMillimeters { get; set; }
    /// <summary>Right paragraph indent measured inside the content width.</summary>
    public double RightIndentMillimeters { get; set; }
    /// <summary>Additional first-line indent relative to the paragraph left edge.</summary>
    public double FirstLineIndentMillimeters { get; set; }
    /// <summary>Optional minimum line height in millimeters; null uses measured metrics.</summary>
    public double? MinimumLineHeightMillimeters { get; set; }
    /// <summary>Whether to begin this block on a new page.</summary>
    public bool PageBreakBefore { get; set; }
}

/// <summary>Supported horizontal paragraph alignment.</summary>
public enum ParagraphAlignment
{
    /// <summary>Align visible line content to the left edge.</summary>
    Left,
    /// <summary>Center visible line content.</summary>
    Center,
    /// <summary>Align visible line content to the right edge.</summary>
    Right,
}

/// <summary>Text and styling local to this span; omitted family, size and color use document defaults.</summary>
public sealed class Span
{
    /// <summary>Creates a span from non-null logical Unicode text.</summary>
    public Span(string text) => Text = text ?? throw new ArgumentNullException(nameof(text));
    /// <summary>Logical Unicode text emitted into native OFD text objects.</summary>
    public string Text { get; set; }
    /// <summary>Optional span font-family override; null uses the document default.</summary>
    public string? FontFamily { get; set; }
    /// <summary>Optional span font size in millimeters; null uses the document default.</summary>
    public double? FontSizeMillimeters { get; set; }
    /// <summary>Whether this span requests bold styling.</summary>
    public bool Bold { get; set; }
    /// <summary>Whether this span requests italic styling.</summary>
    public bool Italic { get; set; }
    /// <summary>Optional span text color; null uses the document default.</summary>
    public OfdColor? Color { get; set; }
}

/// <summary>Geometry in millimeters and per-render resource limits.</summary>
public sealed class FlowDocumentOptions
{
    /// <summary>Output page width in millimeters.</summary>
    public double PageWidthMillimeters { get; set; } = 210;
    /// <summary>Output page height in millimeters.</summary>
    public double PageHeightMillimeters { get; set; } = 297;
    /// <summary>Left page margin in millimeters.</summary>
    public double MarginLeftMillimeters { get; set; } = 25.4;
    /// <summary>Right page margin in millimeters.</summary>
    public double MarginRightMillimeters { get; set; } = 25.4;
    /// <summary>Top page margin in millimeters.</summary>
    public double MarginTopMillimeters { get; set; } = 25.4;
    /// <summary>Bottom page margin in millimeters.</summary>
    public double MarginBottomMillimeters { get; set; } = 25.4;
    /// <summary>Font family used when a span provides no override.</summary>
    public string DefaultFontFamily { get; set; } = "SimSun";
    /// <summary>Default native text font size in millimeters.</summary>
    public double DefaultFontSizeMillimeters { get; set; } = 10.5 * 25.4 / 72;
    /// <summary>Text color used when a span provides no override.</summary>
    public OfdColor DefaultColor { get; set; } = OfdColor.Black;
    /// <summary>Maximum generated pages; exceeding the limit rejects rendering.</summary>
    public int MaxPageCount { get; set; } = 10_000;
    /// <summary>Maximum logical input characters admitted for rendering.</summary>
    public int MaxCharacters { get; set; } = 1_000_000;
    /// <summary>Maximum generated native text elements.</summary>
    public int MaxTextElements { get; set; } = 1_000_000;
    /// <summary>Maximum table cells across all blocks, including empty cells.</summary>
    public int MaxTableCells { get; set; } = 100_000;
}

internal sealed partial class FlowDocumentRenderer : IFlowFontMetrics
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
    private readonly List<double> _pendingLineBreakHeights = new();

    internal FlowDocumentRenderer(FlowDocumentOptions options, IList<FlowBlock> blocks, CancellationToken cancellationToken)
    { _options = options; _blocks = blocks; _cancellationToken = cancellationToken; }

    internal OfdDocumentPackage Render()
    {
        ValidateOptions();
        _package.Options.DefaultPageWidthMillimeters = _options.PageWidthMillimeters;
        _package.Options.DefaultPageHeightMillimeters = _options.PageHeightMillimeters;
        var blocks = _blocks.ToArray();
        var characterCount = 0L;
        var cellCount = 0L;
        foreach (var block in blocks)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            IEnumerable<Paragraph> paragraphs;
            if (block is Paragraph p) paragraphs = new[] { p };
            else if (block is Table table)
            {
                var list = new List<Paragraph>();
                foreach (var row in table.Rows)
                {
                    if (row is null) throw new ArgumentException("A table contains a null row.");
                    foreach (var cell in row.Cells)
                    {
                        _cancellationToken.ThrowIfCancellationRequested();
                        if (++cellCount > _options.MaxTableCells) throw new InvalidOperationException("Flow tables exceed MaxTableCells.");
                        if (cell is null) throw new ArgumentException("A row contains a null cell.");
                        list.AddRange(cell.Paragraphs);
                    }
                }
                paragraphs = list;
            }
            else throw new NotSupportedException("Unsupported flow block.");
            foreach (var paragraph in paragraphs)
            {
                if (paragraph is null) throw new ArgumentException("A cell contains a null paragraph.");
                foreach (var span in paragraph.Spans)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    if (span is null || span.Text is null) throw new ArgumentException("A paragraph contains a null span or text.");
                    characterCount += span.Text.Length;
                    if (characterCount > _options.MaxCharacters) throw new InvalidOperationException("Flow text exceeds MaxCharacters.");
                }
            }
            if (block is Paragraph paragraphBlock) RenderParagraph(paragraphBlock, paragraphBlock.Spans.ToArray());
            else RenderTable((Table)block);
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
        if (paragraph.PageBreakBefore && (_hasBodyContent || hadPendingBreaks || _pendingLineBreakHeights.Count > 0))
        {
            StartPage();
            _pendingSpaceAfter = 0;
            _pendingLineBreakHeights.Clear();
        }
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
            if (line.TerminalNewline) { _pendingLineBreakHeights.Add(line.Height); continue; }
            ApplyPendingPageBreaks();
            if (firstLine)
            {
                ApplyPendingLineBreaks();
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

    private void DrawLine(FlowLine line, double left, double width, double? top = null)
    {
        var x = FlowParagraphLayout.Align(left + line.Indent, width - line.Indent,
            line.AlignmentWidth, line.Alignment);
        for (var i = 0; i < line.Glyphs.Count;)
        {
            var style = line.Glyphs[i].Style;
            var start = i++;
            while (i < line.Glyphs.Count && ReferenceEquals(line.Glyphs[i].Style, style)) i++;
            if (++_textElements > _options.MaxTextElements)
                throw new InvalidOperationException("Flow text exceeds MaxTextElements.");
            var group = line.Glyphs.GetRange(start, i - start);
            var groupWidth = group.Sum(g => g.Width);
            var elementWidth = Math.Max(groupWidth, 0.1);
            var font = GetFont(style);
            var text = string.Concat(group.Select(g => g.Text));
            var element = new OfdTextElement
            {
                LayerType = "Body", XMillimeters = groupWidth == 0 ? Math.Min(x, left + width - elementWidth) : x,
                YMillimeters = top ?? _y, WidthMillimeters = elementWidth, HeightMillimeters = line.Height,
                FontName = font.FontName, FontResourceId = font.Id,
                FontSizeMillimeters = style.FontSizeMillimeters,
                Weight = style.Bold ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight,
                Italic = style.Italic, FillColor = (OfdColor)style.Source!, Text = text
            };
            element.Runs.Add(new OfdTextRun { Text = text, YMillimeters = line.BaselineMillimeters,
                DeltaX = group.Count > 1 ? string.Join(" ", group.Take(group.Count - 1)
                    .Select(g => g.Width.ToString("0.######", CultureInfo.InvariantCulture))) : null });
            _page!.Elements.Add(element);
            x += groupWidth;
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

    private void ApplyPendingLineBreaks()
    {
        foreach (var height in _pendingLineBreakHeights)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            Place(height);
            _y += height;
            _hasBodyContent = true;
        }
        _pendingLineBreakHeights.Clear();
    }

    private void ApplyPendingPageBreaks()
    {
        if (_pendingPageBreaks > 0) { _pendingSpaceAfter = 0; _pendingLineBreakHeights.Clear(); }
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
        if (_options.MaxPageCount < 1 || _options.MaxCharacters < 1 || _options.MaxTextElements < 1 || _options.MaxTableCells < 1)
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
