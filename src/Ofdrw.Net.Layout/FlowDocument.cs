using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Layout.Internal.Flow;

namespace Ofdrw.Net.Layout;

/// <summary>A reusable description of flowing OFD content. Each render creates an independent package.</summary>
public sealed class FlowDocument
{
    /// <summary>Page, margin, font and resource limits for this document.</summary>
    public FlowDocumentOptions Options { get; } = new();

    /// <summary>Ordered paragraph, table, named-area, image and canvas blocks. Do not mutate the description during rendering.</summary>
    public IList<FlowBlock> Blocks { get; } = new List<FlowBlock>();

    /// <summary>Called once per final page, in page order, after all layout succeeds. Drawn objects are behind flow content. Throwing fails Render.</summary>
    public Action<OfdGraphics, FlowBoundary>? OnPage { get; set; }

    /// <summary>Called once per final flow fragment, in document order, after all layout succeeds. Drawn objects follow that fragment. Throwing fails Render.</summary>
    public Action<OfdGraphics, FlowElementPlacement>? OnElement { get; set; }

    /// <summary>Measures and paginates the blocks into native OFD content and named blank regions.</summary>
    public OfdDocumentPackage Render(CancellationToken cancellationToken = default) =>
        new FlowDocumentRenderer(Options, Blocks, OnPage, OnElement, cancellationToken).Render();
}

/// <summary>Base of the built-in flow blocks.</summary>
public abstract class FlowBlock
{
    internal FlowBlock() { }
}

/// <summary>A named, full-content-width blank region reserved during flow pagination.</summary>
public sealed class AreaHolderBlock : FlowBlock
{
    /// <summary>Creates an unfilled region with a fixed height in millimeters.</summary>
    public AreaHolderBlock(string name, double heightMillimeters)
    { Name = name; HeightMillimeters = heightMillimeters; }

    /// <summary>Case-sensitive name used for later lookup after the OFD is reopened.</summary>
    public string Name { get; }
    /// <summary>Non-splittable reserved height in millimeters.</summary>
    public double HeightMillimeters { get; }
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
    /// <summary>Optional numeric font weight from 100 to 900 in steps of 100. Conflicts with Bold unless it is 700.</summary>
    public int? FontWeight { get; set; }
    /// <summary>Additional nonnegative millimeters between adjacent graphemes of this span on the same line.</summary>
    public double LetterSpacingMillimeters { get; set; }
    /// <summary>Draw a line below this span's final text fragments.</summary>
    public bool Underline { get; set; }
    /// <summary>Optional baseline-to-underline offset in millimeters; null uses 10% of the font size.</summary>
    public double? UnderlineOffsetMillimeters { get; set; }
    /// <summary>Optional underline stroke width in millimeters; null uses 6% of the font size.</summary>
    public double? UnderlineWidthMillimeters { get; set; }
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
    private readonly Action<OfdGraphics, FlowBoundary>? _onPage;
    private readonly Action<OfdGraphics, FlowElementPlacement>? _onElement;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<string, OfdFontResource> _fonts = new(StringComparer.Ordinal);
    private readonly OfdDocumentPackage _package = new();
    private readonly List<OfdNamedAreaIndex.Entry> _areas = new();
    private readonly HashSet<string> _areaNames = new(StringComparer.Ordinal);
    private OfdPage? _page;
    private double _y;
    private bool _hasBodyContent;
    private int _textElements;
    private int _pendingPageBreaks;
    private double _pendingSpaceAfter;
    private readonly List<double> _pendingLineBreakHeights = new();
    private readonly List<(OfdPage Page, int InsertIndex, FlowElementPlacement Placement, Action<OfdGraphics>? Draw)> _placements = new();

    internal FlowDocumentRenderer(FlowDocumentOptions options, IList<FlowBlock> blocks,
        Action<OfdGraphics, FlowBoundary>? onPage, Action<OfdGraphics, FlowElementPlacement>? onElement,
        CancellationToken cancellationToken)
    { _options = options; _blocks = blocks; _onPage = onPage; _onElement = onElement; _cancellationToken = cancellationToken; }

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
            else if (block is AreaHolderBlock or FlowFixedBlock)
                paragraphs = Array.Empty<Paragraph>();
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
            else if (block is AreaHolderBlock holder) RenderArea(holder);
            else if (block is FlowFixedBlock fixedBlock) RenderFixedBlock(fixedBlock);
            else RenderTable((Table)block);
        }
        if (_page is null) StartPage();
        if (_areas.Count > 0)
        {
            for (var i = 0; i < _package.Pages.Count; i++)
                _package.Pages[i].Id = (i + 1).ToString(CultureInfo.InvariantCulture);
            OfdNamedAreaIndex.Save(_package, _areas);
        }
        CompleteCallbacks();
        return _package;
    }

    private void RenderFixedBlock(FlowFixedBlock block)
    {
        GraphicsValidation.Positive(block.WidthMillimeters);
        GraphicsValidation.Positive(block.HeightMillimeters);
        ValidateLength(block.SpaceBeforeMillimeters, nameof(block.SpaceBeforeMillimeters));
        ValidateLength(block.SpaceAfterMillimeters, nameof(block.SpaceAfterMillimeters));
        if (!Enum.IsDefined(typeof(ParagraphAlignment), block.Alignment))
            throw new ArgumentOutOfRangeException(nameof(block.Alignment));
        var bodyWidth = _options.PageWidthMillimeters - _options.MarginLeftMillimeters - _options.MarginRightMillimeters;
        if (block.WidthMillimeters > bodyWidth + 0.000001d)
            throw new InvalidOperationException("A fixed flow block is wider than the usable page area.");
        byte[]? imageBytes = null;
        var imageAlpha = 255;
        if (block is ImageBlock image)
        {
            if (image.Alpha is < 0 or > 255) throw new ArgumentOutOfRangeException(nameof(image.Alpha));
            imageAlpha = image.Alpha;
            imageBytes = image.SnapshotImage();
            var limits = new OfdGraphicsOptions();
            if (imageBytes.Length == 0 || imageBytes.Length > limits.MaxEncodedImageBytes)
                throw new ArgumentException("Image block exceeds the graphics image byte budget or is empty.");
            OfdGraphicsImage.Validate(imageBytes, limits.MaxImagePixels, _cancellationToken);
        }
        if (_page is null) StartPage();
        var hadPendingBreaks = _pendingPageBreaks > 0;
        ApplyPendingPageBreaks();
        if (block.PageBreakBefore && (_hasBodyContent || hadPendingBreaks || _pendingLineBreakHeights.Count > 0))
        {
            StartPage();
            _pendingSpaceAfter = 0;
            _pendingLineBreakHeights.Clear();
        }
        ApplyPendingLineBreaks();
        PlaceFirstLine(block.HeightMillimeters, _pendingSpaceAfter + block.SpaceBeforeMillimeters);
        _pendingSpaceAfter = 0;
        var x = FlowParagraphLayout.Align(_options.MarginLeftMillimeters, bodyWidth, block.WidthMillimeters,
            (FlowAlignment)block.Alignment);
        var boundary = new FlowBoundary(_page!.Index, x, _y, block.WidthMillimeters, block.HeightMillimeters);
        Action<OfdGraphics>? draw = block switch
        {
            ImageBlock => graphics => graphics.DrawImage(imageBytes!, boundary.XMillimeters, boundary.YMillimeters,
                boundary.WidthMillimeters, boundary.HeightMillimeters, imageAlpha, _cancellationToken),
            CanvasBlock canvas => graphics =>
            {
                graphics.Translate(boundary.XMillimeters, boundary.YMillimeters);
                graphics.IntersectClip(new OfdGraphicsPath().AddRectangle(0, 0,
                    boundary.WidthMillimeters, boundary.HeightMillimeters), _cancellationToken);
                canvas.Draw(graphics, boundary);
            },
            _ => throw new NotSupportedException("Unsupported fixed flow block.")
        };
        RecordPlacement(block, boundary, draw);
        _y += block.HeightMillimeters;
        _hasBodyContent = true;
        _pendingSpaceAfter = block.SpaceAfterMillimeters;
    }

    private void RecordPlacement(FlowBlock block, FlowBoundary boundary, Action<OfdGraphics>? draw = null)
    {
        if (draw is null && _onElement is null) return;
        _placements.Add((_page!, _page!.Elements.Count, new FlowElementPlacement(block, boundary), draw));
    }

    private void CompleteCallbacks()
    {
        _cancellationToken.ThrowIfCancellationRequested();
        var insertedByPage = new Dictionary<OfdPage, int>();
        foreach (var page in _package.Pages)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_onPage is null) continue;
            var before = page.Elements.Count;
            _onPage(new OfdGraphics(_package, page), new FlowBoundary(page.Index,
                page.XMillimeters, page.YMillimeters, page.WidthMillimeters, page.HeightMillimeters));
            var background = page.Elements.Skip(before).ToArray();
            if (background.Length == 0) continue;
            page.Elements.RemoveRange(before, background.Length);
            page.Elements.InsertRange(0, background);
            insertedByPage[page] = background.Length;
        }
        foreach (var pending in _placements)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (pending.Draw is null && _onElement is null) continue;
            var graphics = new OfdGraphics(_package, pending.Page);
            var before = pending.Page.Elements.Count;
            pending.Draw?.Invoke(graphics);
            if (_onElement is not null)
                _onElement(new OfdGraphics(_package, pending.Page), pending.Placement);
            var added = pending.Page.Elements.Skip(before).ToArray();
            if (added.Length == 0) continue;
            pending.Page.Elements.RemoveRange(before, added.Length);
            insertedByPage.TryGetValue(pending.Page, out var inserted);
            pending.Page.Elements.InsertRange(pending.InsertIndex + inserted, added);
            insertedByPage[pending.Page] = inserted + added.Length;
        }
    }

    private void RenderArea(AreaHolderBlock holder)
    {
        OfdNamedAreaIndex.ValidateName(holder.Name);
        ValidatePositive(holder.HeightMillimeters, nameof(holder.HeightMillimeters));
        if (!_areaNames.Add(holder.Name)) throw new ArgumentException("Duplicate named area.", nameof(holder));
        if (_areas.Count >= OfdNamedAreaIndex.MaxAreas) throw new InvalidOperationException("Named area limit exceeded.");
        if (_page is null) StartPage();
        ApplyPendingPageBreaks();
        ApplyPendingLineBreaks();
        PlaceFirstLine(holder.HeightMillimeters, _pendingSpaceAfter);
        _pendingSpaceAfter = 0;
        _areas.Add(new OfdNamedAreaIndex.Entry(holder.Name, _page!, _y,
            _options.MarginLeftMillimeters,
            _options.PageWidthMillimeters - _options.MarginLeftMillimeters - _options.MarginRightMillimeters,
            holder.HeightMillimeters));
        if (_onElement is not null)
            RecordPlacement(holder, new FlowBoundary(_page!.Index, _options.MarginLeftMillimeters, _y,
                _options.PageWidthMillimeters - _options.MarginLeftMillimeters - _options.MarginRightMillimeters,
                holder.HeightMillimeters));
        _y += holder.HeightMillimeters;
        _hasBodyContent = true;
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
            inlines.Add(new FlowInline { Text = span.Text, Style = CreateStyle(span) });
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
            if (_onElement is not null)
                RecordPlacement(paragraph, new FlowBoundary(_page!.Index,
                    FlowParagraphLayout.Align(_options.MarginLeftMillimeters + paragraph.LeftIndentMillimeters + line.Indent,
                        width - line.Indent, line.AlignmentWidth, line.Alignment),
                    _y, line.AlignmentWidth, line.Height));
            _y += line.Height;
            _hasBodyContent = true;
        }
        // Carry the gap to the next block, without materializing a trailing empty page.
        _pendingSpaceAfter += paragraph.SpaceAfterMillimeters;
    }

    private FlowTextStyle CreateStyle(Span span)
    {
        var size = span.FontSizeMillimeters ?? _options.DefaultFontSizeMillimeters;
        ValidatePositive(size, nameof(span.FontSizeMillimeters));
        var family = span.FontFamily ?? _options.DefaultFontFamily;
        if (string.IsNullOrWhiteSpace(family)) throw new ArgumentException("Font family cannot be empty.");
        var weight = span.FontWeight ?? (span.Bold ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight);
        if (weight < 100 || weight > 900 || weight % 100 != 0 || span.Bold && weight != OfdTextElement.BoldWeight)
            throw new ArgumentOutOfRangeException(nameof(span.FontWeight));
        ValidateLength(span.LetterSpacingMillimeters, nameof(span.LetterSpacingMillimeters));
        var underlineOffset = span.UnderlineOffsetMillimeters ?? size * 0.1d;
        var underlineWidth = span.UnderlineWidthMillimeters ?? size * 0.06d;
        if (span.Underline)
        {
            ValidatePositive(underlineOffset, nameof(span.UnderlineOffsetMillimeters));
            GraphicsValidation.Positive(underlineWidth);
            if (underlineOffset + underlineWidth / 2 > size * 0.3d)
                throw new ArgumentOutOfRangeException(nameof(span.UnderlineOffsetMillimeters),
                    "Underline must fit below the baseline within the normal line height.");
        }
        else if (span.UnderlineOffsetMillimeters.HasValue || span.UnderlineWidthMillimeters.HasValue)
            throw new ArgumentException("Underline geometry requires Underline to be enabled.");
        return new FlowTextStyle { FontFamily = family.Trim(), FontSizeMillimeters = size,
            Bold = weight >= 600, Italic = span.Italic, Source = span.Color ?? _options.DefaultColor,
            Weight = weight, LetterSpacingMillimeters = span.LetterSpacingMillimeters,
            Underline = span.Underline, UnderlineOffsetMillimeters = underlineOffset,
            UnderlineWidthMillimeters = underlineWidth };
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
                Weight = style.Weight,
                Italic = style.Italic, FillColor = (OfdColor)style.Source!, Text = text
            };
            element.Runs.Add(new OfdTextRun { Text = text, YMillimeters = line.BaselineMillimeters,
                DeltaX = group.Count > 1 ? string.Join(" ", group.Take(group.Count - 1)
                    .Select(g => g.Width.ToString("0.######", CultureInfo.InvariantCulture))) : null });
            _page!.Elements.Add(element);
            if (style.Underline && groupWidth > 0)
            {
                var underlineY = line.BaselineMillimeters + style.UnderlineOffsetMillimeters;
                _page.Elements.Add(new OfdPathElement
                {
                    LayerType = "Body", XMillimeters = x, YMillimeters = top ?? _y,
                    WidthMillimeters = groupWidth, HeightMillimeters = line.Height,
                    Stroke = true, Fill = false, LineWidthMillimeters = style.UnderlineWidthMillimeters,
                    StrokeColor = (OfdColor)style.Source!,
                    AbbreviatedData = FormattableString.Invariant($"M 0 {underlineY:0.######} L {groupWidth:0.######} {underlineY:0.######}")
                });
            }
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
