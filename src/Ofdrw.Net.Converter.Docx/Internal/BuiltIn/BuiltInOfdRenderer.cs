using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Pdf;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

/// <summary>
/// Lays out a <see cref="BuiltInDocumentModel"/> directly as structured OFD page
/// objects (text + table borders) without PDF rasterization.
/// </summary>
internal sealed class BuiltInOfdRenderer
{
    private const double PointsPerInch = 72d;
    private const double MillimetersPerInch = 25.4d;
    private const double DefaultFontSizePoints = 10.5d;
    private const double MinCellPaddingMillimeters = 0.8d;
    private readonly DocxConversionOptions _options;
    private readonly IList<DocxConversionDiagnostic> _diagnostics;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<string, double> _advances = new();
    private readonly List<PageDecoration> _decorations = new();
    private int _lastPageNumber;
    private readonly Dictionary<string, XFont> _fonts = new();
    private readonly Dictionary<string, (byte[] Data, string FileName)> _fontFiles = new();

    internal BuiltInOfdRenderer(
        DocxConversionOptions options,
        IList<DocxConversionDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        _options = options;
        _diagnostics = diagnostics;
        _cancellationToken = cancellationToken;
    }

    internal OfdDocumentPackage Render(BuiltInDocumentModel model)
    {
        _cancellationToken.ThrowIfCancellationRequested();

        RegisterConfiguredFonts();
        var package = new OfdDocumentPackage
        {
            Options = new OfdDocumentOptions
            {
                DocType = "OFD-H",
                DocumentId = "Doc_0",
                Metadata = new OfdMetadata
                {
                    Title = "DOCX document",
                    Creator = "Ofdrw.Net BuiltIn OFD renderer",
                    CreationDate = DateTimeOffset.UtcNow,
                    ModificationDate = DateTimeOffset.UtcNow
                }
            }
        };
        package.CustomTags["source-text-origin"] = "DOCX/OpenXML";
        package.CustomTags["source-text-kind"] = "machine-readable";
        package.CustomTags["docx-ofd-mode"] = "Native";

        var sections = model.Sections.Count == 0
            ? new List<BuiltInSectionModel> { new BuiltInSectionModel() }
            : model.Sections;

        LayoutState? lastState = null;
        foreach (var section in sections)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            lastState = RenderSection(package, section);
        }
        if (lastState is not null)
        {
            foreach (var note in model.SupplementalText)
            {
                RenderParagraph(package, lastState, note.CreateHeading(), lastState.ContentLeft, lastState.ContentWidth);
                foreach (var block in note.Blocks) RenderBlock(package, lastState, block);
            }
        }
        if (model.SupplementalText.Count > 0)
            package.CustomTags["source-text-supplemental-layout"] = "labeled-after-body";

        if (package.Pages.Count == 0)
        {
            package.Pages.Add(CreatePage(package, new BuiltInSectionModel()));
        }

        RenderDecorations(package);
        return package;
    }

    private LayoutState RenderSection(OfdDocumentPackage package, BuiltInSectionModel section)
    {
        var state = new LayoutState(section, section.PageNumberStart ?? _lastPageNumber + 1);
        EnsurePage(package, state);
        foreach (var block in section.Blocks)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            RenderBlock(package, state, block);
        }
        return state;
    }

    private void RenderBlock(OfdDocumentPackage package, LayoutState state, BuiltInBlockModel block)
    {
        switch (block)
        {
            case BuiltInParagraphModel paragraph:
                RenderParagraph(package, state, paragraph, state.ContentLeft, state.ContentWidth);
                break;
            case BuiltInTableModel table:
                RenderTable(package, state, table);
                state.HasBodyContent = true;
                break;
        }
    }

    private void RenderParagraph(
        OfdDocumentPackage package,
        LayoutState state,
        BuiltInParagraphModel paragraph,
        double left,
        double availableWidth)
    {
        if (paragraph.Format.PageBreakBefore && state.HasBodyContent)
        {
            StartNewPage(package, state);
        }

        var spaceBefore = PointsToMillimeters(paragraph.Format.SpaceBeforePoints ?? 0);
        var spaceAfter = PointsToMillimeters(paragraph.Format.SpaceAfterPoints ?? 0);
        EnsureVerticalSpace(package, state, spaceBefore);
        state.Y += spaceBefore;

        var leftIndent = PointsToMillimeters(paragraph.Format.LeftIndentPoints ?? 0);
        var rightIndent = PointsToMillimeters(paragraph.Format.RightIndentPoints ?? 0);
        var firstIndent = PointsToMillimeters(paragraph.Format.FirstLineIndentPoints ?? 0);
        var width = Math.Max(availableWidth - leftIndent - rightIndent, 5d);
        foreach (var line in LayoutParagraph(paragraph, width, firstIndent, state.PageNumber))
        {
            if (line.PageBreak)
            {
                StartNewPage(package, state);
                continue;
            }
            EnsureVerticalSpace(package, state, line.Height);
            DrawLine(package, state.Page!, line, left + leftIndent, state.Y, width);
            state.Y += line.Height;
            state.HasBodyContent = true;
        }

        EnsureVerticalSpace(package, state, spaceAfter);
        state.Y += spaceAfter;
    }

    private void RenderTable(
        OfdDocumentPackage package,
        LayoutState state,
        BuiltInTableModel table)
    {
        var columnWidths = ResolveColumnWidths(table, state.ContentWidth);

        foreach (var row in table.Rows)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var cellLayouts = MeasureRow(row, columnWidths);
            var rowHeight = cellLayouts.Count == 0
                ? PointsToMillimeters(DefaultFontSizePoints) * 1.5d
                : cellLayouts.Max(cell => cell.Height);

            EnsureVerticalSpace(package, state, rowHeight);
            var rowTop = state.Y;
            var x = state.ContentLeft;
            var columnIndex = 0;

            foreach (var cellLayout in cellLayouts)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                var cellWidth = 0d;
                for (var span = 0; span < cellLayout.ColumnSpan && columnIndex + span < columnWidths.Count; span++)
                {
                    cellWidth += columnWidths[columnIndex + span];
                }

                var cell = row.Cells[cellLayouts.IndexOf(cellLayout)];
                DrawCell(state.Page!, table, cell, table.Rows.IndexOf(row), columnIndex, columnWidths.Count,
                    x, rowTop, cellWidth, rowHeight);
                var textLeft = x + MinCellPaddingMillimeters;
                var textWidth = Math.Max(cellWidth - MinCellPaddingMillimeters * 2, 1d);
                var localY = AlignBlockVertically(rowTop, rowHeight,
                    cellLayout.Lines.Sum(line => line.Height), cellLayout.VerticalAlignment);
                foreach (var line in cellLayout.Lines)
                {
                    DrawLine(package, state.Page!, line, textLeft, localY, textWidth);
                    localY += line.Height;
                }

                x += cellWidth;
                columnIndex += cellLayout.ColumnSpan;
            }

            state.Y = rowTop + rowHeight;
        }
    }

    private List<CellLayout> MeasureRow(BuiltInTableRowModel row, IReadOnlyList<double> columnWidths)
    {
        var layouts = new List<CellLayout>();
        var columnIndex = 0;
        foreach (var cell in row.Cells)
        {
            if (columnIndex >= columnWidths.Count)
            {
                break;
            }

            var span = Math.Max(1, cell.ColumnSpan);
            var cellWidth = 0d;
            for (var i = 0; i < span && columnIndex + i < columnWidths.Count; i++)
            {
                cellWidth += columnWidths[columnIndex + i];
            }

            var textWidth = Math.Max(cellWidth - (MinCellPaddingMillimeters * 2), 1d);
            var lines = new List<StyledLine>();
            foreach (var paragraph in cell.Paragraphs)
            {
                lines.AddRange(LayoutParagraph(paragraph, textWidth, 0).Where(line => !line.PageBreak));
            }

            var height = lines.Sum(line => line.Height) + (MinCellPaddingMillimeters * 2);
            layouts.Add(new CellLayout(span, height, lines, cell.VerticalAlignment));
            columnIndex += span;
        }

        return layouts;
    }

    private DocxFontCatalog _configuredFonts = null!;

    private void RegisterConfiguredFonts()
    {
        _configuredFonts = new DocxFontCatalog(_options, _diagnostics, _cancellationToken);
    }

    private string FontKey(BuiltInTextFormat format) =>
        NormalizeFontFamily(format.FontFamily) + (format.Bold ? "|bold" : "|regular") + (format.Italic ? "|italic" : "");

    private XFont GetFont(BuiltInTextFormat format)
    {
        var key = FontKey(format);
        if (!_fonts.TryGetValue(key, out var font))
        {
            PdfFontRegistry.EnsureInstalled();
            font = new XFont(_configuredFonts.Resolve(NormalizeFontFamily(format.FontFamily), format.Bold, format.Italic), 1000,
                (format.Bold ? XFontStyle.Bold : XFontStyle.Regular) |
                (format.Italic ? XFontStyle.Italic : XFontStyle.Regular));
            _fonts[key] = font;
        }
        return font;
    }

    private double Advance(string text, BuiltInTextFormat format)
    {
        var key = FontKey(format) + "\n" + text;
        if (!_advances.TryGetValue(key, out var advance))
        {
            using var measure = XGraphics.CreateMeasureContext(new XSize(1000, 1000), XGraphicsUnit.Point, XPageDirection.Downwards);
            advance = measure.MeasureString(text == "\t" ? "    " : text, GetFont(format)).Width / 1000d;
            _advances[key] = advance;
        }
        return advance * PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints);
    }

    private List<StyledLine> LayoutParagraph(BuiltInParagraphModel paragraph, double width, double firstIndent, int pageNumber = 1, int totalPages = 1, int sectionPages = 1)
    {
        var glyphs = new List<StyledGlyph>();
        var fallback = paragraph.Inlines.OfType<BuiltInTextModel>().FirstOrDefault()?.Format ?? new BuiltInTextFormat();
        void Add(string value, BuiltInTextFormat format)
        {
            foreach (var glyph in EnumerateTextElements(value.Replace("\r\n", "\n").Replace('\r', '\n')))
                glyphs.Add(new StyledGlyph(glyph, format, glyph == "\n" || glyph == "\f" ? 0 : Advance(glyph, format)));
        }
        foreach (var inline in paragraph.Inlines)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            switch (inline)
            {
                case BuiltInTextModel text: Add(text.Text, text.Format); break;
                case BuiltInTabModel: Add("\t", fallback); break;
                case BuiltInBreakModel br: Add(br.IsPageBreak ? "\f" : "\n", fallback); break;
                case BuiltInPageNumberModel field:
                    var value = field.Kind == BuiltInPageFieldKind.TotalPages ? totalPages :
                        field.Kind == BuiltInPageFieldKind.SectionPages ? sectionPages : pageNumber;
                    Add(value.ToString(CultureInfo.InvariantCulture), field.Format);
                    break;
                case BuiltInImageModel image:
                    var imageWidth = PointsToMillimeters(image.WidthPoints ?? 72);
                    var imageHeight = PointsToMillimeters(image.HeightPoints ?? 72);
                    if (imageWidth > width) { imageHeight *= width / imageWidth; imageWidth = width; }
                    glyphs.Add(new StyledGlyph(string.Empty, fallback, imageWidth) { Image = image, ImageHeight = imageHeight });
                    break;
            }
        }
        var result = new List<StyledLine>();
        var current = new List<StyledGlyph>();
        var currentWidth = 0d;
        var indent = firstIndent;
        void Flush()
        {
            var size = current.Count == 0 ? PointsToMillimeters(fallback.FontSizePoints ?? DefaultFontSizePoints)
                : current.Max(g => PointsToMillimeters(g.Format.FontSizePoints ?? DefaultFontSizePoints));
            var imageHeight = current.Where(glyph => glyph.Image is not null).Select(glyph => glyph.ImageHeight).DefaultIfEmpty(0).Max();
            result.Add(new StyledLine(current, Math.Max(imageHeight, Math.Max(size * 1.3, PointsToMillimeters(paragraph.Format.LineSpacingPoints ?? 0))),
                indent, paragraph.Format.Alignment));
            current = new List<StyledGlyph>();
            currentWidth = 0;
            indent = 0;
        }
        for (var i = 0; i < glyphs.Count; i++)
        {
            var glyph = glyphs[i];
            if (glyph.Text == "\f" || glyph.Text == "\n")
            {
                if (current.Count > 0 || glyph.Text == "\n") Flush();
                if (glyph.Text == "\f") result.Add(new StyledLine(new List<StyledGlyph>(), 0, 0, paragraph.Format.Alignment) { PageBreak = true });
                continue;
            }
            // Keep Latin words intact when they fit on a fresh line, even across run boundaries.
            bool Word(StyledGlyph g) => g.Text.Length == 1 && g.Text[0] < 128 && char.IsLetterOrDigit(g.Text[0]);
            if (Word(glyph) && (i == 0 || !Word(glyphs[i - 1])))
            {
                var wordWidth = 0d;
                for (var j = i; j < glyphs.Count && Word(glyphs[j]); j++) wordWidth += glyphs[j].Width;
                if (current.Count > 0 && wordWidth <= width && currentWidth + wordWidth > width - indent) Flush();
            }
            if (current.Count > 0 && currentWidth + glyph.Width > width - indent) Flush();
            current.Add(glyph);
            currentWidth += glyph.Width;
        }
        if (current.Count > 0 || result.Count == 0) Flush();
        return result;
    }

    private void DrawLine(OfdDocumentPackage package, OfdPage page, StyledLine line, double left, double top, double width)
    {
        var x = AlignHorizontally(left + line.Indent, width - line.Indent, line.Glyphs.Sum(g => g.Width), line.Alignment);
        var baseline = line.Glyphs.Count == 0 ? 0 : line.Glyphs.Max(g => PointsToMillimeters(g.Format.FontSizePoints ?? DefaultFontSizePoints));
        for (var i = 0; i < line.Glyphs.Count;)
        {
            var start = i;
            if (line.Glyphs[i].Image is BuiltInImageModel image)
            {
                var glyph = line.Glyphs[i++];
                page.Elements.Add(new OfdImageElement
                {
                    XMillimeters = x, YMillimeters = top, WidthMillimeters = glyph.Width, HeightMillimeters = glyph.ImageHeight,
                    Data = image.Data, FileName = image.Name, MediaType = image.MediaType
                });
                x += glyph.Width;
                continue;
            }
            var format = line.Glyphs[i].Format;
            while (i < line.Glyphs.Count && line.Glyphs[i].Image is null && ReferenceEquals(line.Glyphs[i].Format, format)) i++;
            var group = line.Glyphs.GetRange(start, i - start);
            var fontKey = FontKey(format);
            if (!package.Fonts.Any(f => f.FontName == fontKey))
            {
                var resolver = GlobalFontSettings.FontResolver;
                var face = resolver.ResolveTypeface(_configuredFonts.Resolve(NormalizeFontFamily(format.FontFamily), format.Bold, format.Italic), format.Bold, format.Italic);
                if (!_fontFiles.TryGetValue(face.FaceName, out var file))
                {
                    var data = PdfFontRegistry.GetOriginalFont(face.FaceName);
                    using var hash = SHA256.Create();
                    var digest = BitConverter.ToString(hash.ComputeHash(data)).Replace("-", string.Empty).ToLowerInvariant();
                    file = (data, "native-font-" + digest + ".ttf");
                    _fontFiles[face.FaceName] = file;
                }
                package.Fonts.Add(new OfdFontResource
                {
                    FontName = fontKey, FamilyName = GetFont(format).Name, Charset = "unicode",
                    Bold = format.Bold, Italic = format.Italic,
                    FileName = file.FileName,
                    Data = file.Data
                });
            }
            var text = new OfdTextElement
            {
                LayerType = "Body", XMillimeters = x, YMillimeters = top,
                WidthMillimeters = Math.Max(group.Sum(g => g.Width), 0.1), HeightMillimeters = line.Height,
                FontName = fontKey, FontSizeMillimeters = PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints),
                FillColor = ParseColor(format.ColorHex), Text = string.Concat(group.Select(g => g.Text))
            };
            text.Runs.Add(new OfdTextRun { Text = text.Text, YMillimeters = baseline,
                DeltaX = group.Count > 1 ? string.Join(" ", group.Take(group.Count - 1).Select(g => g.Width.ToString("0.######", CultureInfo.InvariantCulture))) : null });
            page.Elements.Add(text);
            x += group.Sum(g => g.Width);
        }
    }

    private sealed class StyledGlyph
    {
        internal StyledGlyph(string text, BuiltInTextFormat format, double width) { Text = text; Format = format; Width = width; }
        internal string Text { get; }
        internal BuiltInTextFormat Format { get; }
        internal double Width { get; }
        internal BuiltInImageModel? Image { get; set; }
        internal double ImageHeight { get; set; }
    }

    private sealed class StyledLine
    {
        internal StyledLine(List<StyledGlyph> glyphs, double height, double indent, BuiltInParagraphAlignment alignment)
        { Glyphs = glyphs; Height = height; Indent = indent; Alignment = alignment; }
        internal List<StyledGlyph> Glyphs { get; }
        internal double Height { get; }
        internal double Indent { get; }
        internal BuiltInParagraphAlignment Alignment { get; }
        internal bool PageBreak { get; set; }
    }

    private static void DrawCell(OfdPage page, BuiltInTableModel table, BuiltInTableCellModel cell,
        int row, int column, int columnCount, double x, double y, double width, double height)
    {
        if (!string.IsNullOrWhiteSpace(cell.ShadingHex) && cell.ShadingHex != "auto")
        {
            page.Elements.Add(new OfdPathElement { LayerType = "Body", XMillimeters = x, YMillimeters = y,
                WidthMillimeters = width, HeightMillimeters = height, Fill = true, Stroke = false,
                FillColor = ParseColor(cell.ShadingHex),
                AbbreviatedData = FormattableString.Invariant($"M 0 0 L {width} 0 L {width} {height} L 0 {height} C") });
        }
        void Border(string side, string fallback, double x1, double y1, double x2, double y2)
        {
            if (!cell.Borders.TryGetValue(side, out var border)) table.Borders.TryGetValue(fallback, out border);
            if (border is null || !border.Visible) return;
            page.Elements.Add(new OfdPathElement { LayerType = "Body", XMillimeters = x, YMillimeters = y,
                WidthMillimeters = width, HeightMillimeters = height, Fill = false, Stroke = true,
                StrokeColor = ParseColor(border.ColorHex), LineWidthMillimeters = PointsToMillimeters(border.WidthPoints),
                AbbreviatedData = FormattableString.Invariant($"M {x1} {y1} L {x2} {y2}") });
        }
        Border("top", row == 0 ? "top" : "insideH", 0, 0, width, 0);
        Border("bottom", row == table.Rows.Count - 1 ? "bottom" : "insideH", 0, height, width, height);
        Border("left", column == 0 ? "left" : "insideV", 0, 0, 0, height);
        Border("right", column + cell.ColumnSpan >= columnCount ? "right" : "insideV", width, 0, width, height);
    }

    private static double AlignHorizontally(
        double left,
        double availableWidth,
        double contentWidth,
        BuiltInParagraphAlignment alignment)
    {
        if (contentWidth >= availableWidth)
        {
            return left;
        }

        return alignment switch
        {
            BuiltInParagraphAlignment.Center => left + ((availableWidth - contentWidth) / 2d),
            BuiltInParagraphAlignment.Right => left + (availableWidth - contentWidth),
            _ => left
        };
    }

    private static double AlignBlockVertically(
        double rowTop,
        double rowHeight,
        double contentHeight,
        BuiltInVerticalAlignment alignment)
    {
        var paddedTop = rowTop + MinCellPaddingMillimeters;
        var available = Math.Max(rowHeight - (MinCellPaddingMillimeters * 2), 0d);
        if (contentHeight >= available)
        {
            return paddedTop;
        }

        return alignment switch
        {
            BuiltInVerticalAlignment.Center => paddedTop + ((available - contentHeight) / 2d),
            BuiltInVerticalAlignment.Bottom => paddedTop + (available - contentHeight),
            _ => paddedTop
        };
    }

    private static List<string> EnumerateTextElements(string text)
    {
        var glyphs = new List<string>();
        if (string.IsNullOrEmpty(text))
        {
            return glyphs;
        }

        var enumerator = StringInfo.GetTextElementEnumerator(text);
        while (enumerator.MoveNext())
        {
            glyphs.Add(enumerator.GetTextElement());
        }

        return glyphs;
    }

    private static IReadOnlyList<double> ResolveColumnWidths(BuiltInTableModel table, double contentWidth)
    {
        if (table.ColumnWidthsPoints.Count > 0)
        {
            var widths = table.ColumnWidthsPoints
                .Select(PointsToMillimeters)
                .Select(width => Math.Max(width, 5d))
                .ToList();
            var total = widths.Sum();
            if (total <= 0)
            {
                return EqualWidths(Math.Max(1, widths.Count), contentWidth);
            }

            if (Math.Abs(total - contentWidth) > 0.5d)
            {
                var scale = contentWidth / total;
                for (var i = 0; i < widths.Count; i++)
                {
                    widths[i] *= scale;
                }
            }

            return widths;
        }

        var columnCount = table.Rows
            .Select(row => row.Cells.Sum(cell => Math.Max(1, cell.ColumnSpan)))
            .DefaultIfEmpty(1)
            .Max();
        return EqualWidths(Math.Max(1, columnCount), contentWidth);
    }

    private static IReadOnlyList<double> EqualWidths(int count, double contentWidth)
    {
        var width = contentWidth / count;
        return Enumerable.Repeat(width, count).ToList();
    }

    private string NormalizeFontFamily(string? family)
    {
        var resolved = string.IsNullOrWhiteSpace(family)
            ? ResolveFallbackFont()
            : family!.Trim();

        // Map common DOCX East-Asia names to fonts OFD viewers resolve locally.
        if (resolved.Equals("宋体", StringComparison.OrdinalIgnoreCase) ||
            resolved.Equals("SimSun", StringComparison.OrdinalIgnoreCase) ||
            resolved.Equals("NSimSun", StringComparison.OrdinalIgnoreCase))
        {
            resolved = "SimSun";
        }
        else if (resolved.Equals("黑体", StringComparison.OrdinalIgnoreCase) ||
                 resolved.Equals("SimHei", StringComparison.OrdinalIgnoreCase))
        {
            resolved = "SimHei";
        }
        else if (resolved.Equals("微软雅黑", StringComparison.OrdinalIgnoreCase) ||
                 resolved.Equals("Microsoft YaHei", StringComparison.OrdinalIgnoreCase) ||
                 resolved.Equals("Microsoft YaHei UI", StringComparison.OrdinalIgnoreCase))
        {
            resolved = "Microsoft YaHei";
        }
        else if (resolved.Equals("楷体", StringComparison.OrdinalIgnoreCase) ||
                 resolved.Equals("KaiTi", StringComparison.OrdinalIgnoreCase))
        {
            resolved = "KaiTi";
        }
        else if (resolved.Equals("仿宋", StringComparison.OrdinalIgnoreCase) ||
                 resolved.Equals("FangSong", StringComparison.OrdinalIgnoreCase))
        {
            resolved = "FangSong";
        }

        return resolved;
    }

    private string ResolveFallbackFont()
    {
        return NormalizeFontFamily(_configuredFonts.ResolveFallbackFamily(_options.FontFallbackFamilies));
    }

    private static OfdColor ParseColor(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
        {
            return OfdColor.Black;
        }

        var trimmed = hex!.Trim();
        var value = trimmed.StartsWith("#", StringComparison.Ordinal) ? trimmed.Substring(1) : trimmed;
        if (value.Length == 6 &&
            int.TryParse(value.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r) &&
            int.TryParse(value.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g) &&
            int.TryParse(value.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
        {
            return new OfdColor(r, g, b);
        }

        return OfdColor.Black;
    }

    private void EnsureVerticalSpace(OfdDocumentPackage package, LayoutState state, double needed)
    {
        if (state.Page is null || state.Y + needed > state.ContentBottom)
        {
            StartNewPage(package, state);
        }
    }

    private void StartNewPage(OfdDocumentPackage package, LayoutState state)
    {
        if (package.Pages.Count >= _options.MaxPageCount)
            throw new InvalidDataException("DOCX exceeds the configured rendered page limit.");
        state.Page = CreatePage(package, state.Section);
        state.PageNumber = state.FirstPageNumber + state.SectionPageIndex;
        _lastPageNumber = state.PageNumber;
        var headers = state.Section.GetHeaders(state.SectionPageIndex, state.PageNumber);
        var footers = state.Section.GetFooters(state.SectionPageIndex, state.PageNumber);
        var header = LayoutDecoration(headers, state.ContentWidth, state.PageNumber, _options.MaxPageCount, _options.MaxPageCount);
        var footer = LayoutDecoration(footers, state.ContentWidth, state.PageNumber, _options.MaxPageCount, _options.MaxPageCount);
        var headerBottom = PointsToMillimeters(state.Section.HeaderDistancePoints) + header.Height;
        var footerTop = state.Page.HeightMillimeters - PointsToMillimeters(state.Section.FooterDistancePoints) - footer.Height;
        state.ContentTop = Math.Max(PointsToMillimeters(state.Section.MarginTopPoints), headers.Count > 0 ? headerBottom + 1 : 0);
        state.ContentBottom = Math.Min(state.Page.HeightMillimeters - PointsToMillimeters(state.Section.MarginBottomPoints), footers.Count > 0 ? footerTop - 1 : state.Page.HeightMillimeters);
        if (state.ContentBottom <= state.ContentTop)
            throw new InvalidDataException("DOCX headers and footers leave no usable body area.");
        _decorations.Add(new PageDecoration(state.Page, state.Section, state.SectionPageIndex, state.PageNumber));
        state.SectionPageIndex++;
        state.Y = state.ContentTop;
        state.HasBodyContent = false;
    }

    private DecorationLayout LayoutDecoration(IEnumerable<BuiltInParagraphModel> paragraphs, double width, int pageNumber, int totalPages, int sectionPages)
    {
        var result = new DecorationLayout();
        foreach (var paragraph in paragraphs)
        {
            result.Height += PointsToMillimeters(paragraph.Format.SpaceBeforePoints ?? 0);
            var left = PointsToMillimeters(paragraph.Format.LeftIndentPoints ?? 0);
            var lineWidth = Math.Max(1, width - left - PointsToMillimeters(paragraph.Format.RightIndentPoints ?? 0));
            foreach (var line in LayoutParagraph(paragraph, lineWidth, PointsToMillimeters(paragraph.Format.FirstLineIndentPoints ?? 0), pageNumber, totalPages, sectionPages))
            {
                if (line.PageBreak) continue;
                result.Lines.Add((line, left, result.Height, lineWidth));
                result.Height += line.Height;
            }
            result.Height += PointsToMillimeters(paragraph.Format.SpaceAfterPoints ?? 0);
        }
        return result;
    }

    private void RenderDecorations(OfdDocumentPackage package)
    {
        var sectionCounts = _decorations.GroupBy(value => value.Section).ToDictionary(group => group.Key, group => group.Count());
        foreach (var decoration in _decorations)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var section = decoration.Section;
            var page = decoration.Page;
            var width = PointsToMillimeters(section.PageWidthPoints - section.MarginLeftPoints - section.MarginRightPoints);
            var left = PointsToMillimeters(section.MarginLeftPoints);
            var header = LayoutDecoration(section.GetHeaders(decoration.SectionPageIndex, decoration.PageNumber), width,
                decoration.PageNumber, package.Pages.Count, sectionCounts[section]);
            var footer = LayoutDecoration(section.GetFooters(decoration.SectionPageIndex, decoration.PageNumber), width,
                decoration.PageNumber, package.Pages.Count, sectionCounts[section]);
            foreach (var item in header.Lines)
                DrawLine(package, page, item.Line, left + item.Left, PointsToMillimeters(section.HeaderDistancePoints) + item.Top, item.Width);
            foreach (var item in footer.Lines)
                DrawLine(package, page, item.Line, left + item.Left,
                    page.HeightMillimeters - PointsToMillimeters(section.FooterDistancePoints) - footer.Height + item.Top, item.Width);
        }
    }

    private sealed class PageDecoration
    {
        internal PageDecoration(OfdPage page, BuiltInSectionModel section, int sectionPageIndex, int pageNumber)
        { Page = page; Section = section; SectionPageIndex = sectionPageIndex; PageNumber = pageNumber; }
        internal OfdPage Page { get; }
        internal BuiltInSectionModel Section { get; }
        internal int SectionPageIndex { get; }
        internal int PageNumber { get; }
    }

    private sealed class DecorationLayout
    {
        internal double Height { get; set; }
        internal List<(StyledLine Line, double Left, double Top, double Width)> Lines { get; } = new();
    }

    private void EnsurePage(OfdDocumentPackage package, LayoutState state)
    {
        if (state.Page is null)
        {
            StartNewPage(package, state);
        }
    }

    private static OfdPage CreatePage(OfdDocumentPackage package, BuiltInSectionModel section)
    {
        var page = new OfdPage
        {
            Index = package.Pages.Count,
            WidthMillimeters = PointsToMillimeters(section.PageWidthPoints),
            HeightMillimeters = PointsToMillimeters(section.PageHeightPoints)
        };
        package.Pages.Add(page);
        return page;
    }

    private static double PointsToMillimeters(double points)
    {
        return points * MillimetersPerInch / PointsPerInch;
    }

    private sealed class LayoutState
    {
        internal LayoutState(BuiltInSectionModel section, int firstPageNumber)
        {
            Section = section;
            FirstPageNumber = firstPageNumber;
            ContentLeft = PointsToMillimeters(section.MarginLeftPoints);
            ContentTop = PointsToMillimeters(section.MarginTopPoints);
            ContentBottom = PointsToMillimeters(section.PageHeightPoints - section.MarginBottomPoints);
            ContentWidth = PointsToMillimeters(
                section.PageWidthPoints - section.MarginLeftPoints - section.MarginRightPoints);
            Y = ContentTop;
        }

        internal BuiltInSectionModel Section { get; }

        internal OfdPage? Page { get; set; }
        internal int FirstPageNumber { get; }
        internal int PageNumber { get; set; }
        internal int SectionPageIndex { get; set; }
        internal bool HasBodyContent { get; set; }

        internal double ContentLeft { get; }

        internal double ContentTop { get; set; }

        internal double ContentBottom { get; set; }

        internal double ContentWidth { get; }

        internal double Y { get; set; }
    }

    private sealed class CellLayout
    {
        internal CellLayout(
            int columnSpan,
            double height,
            List<StyledLine> lines,
            BuiltInVerticalAlignment verticalAlignment)
        {
            ColumnSpan = columnSpan;
            Height = height;
            Lines = lines;
            VerticalAlignment = verticalAlignment;
        }

        internal int ColumnSpan { get; }

        internal double Height { get; }

        internal List<StyledLine> Lines { get; }

        internal BuiltInVerticalAlignment VerticalAlignment { get; }
    }

}
