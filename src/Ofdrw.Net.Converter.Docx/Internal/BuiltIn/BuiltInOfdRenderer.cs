using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Models;

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
    private const double BorderWidthMillimeters = 0.2d;
    /// <summary>Half-width (ASCII) advance relative to font size — matches ofdrw/OFD convention.</summary>
    private const double AsciiWidthFactor = 0.5d;
    /// <summary>Full-width (CJK) advance relative to font size.</summary>
    private const double CjkWidthFactor = 1.0d;

    private readonly DocxConversionOptions _options;
    private readonly IList<DocxConversionDiagnostic> _diagnostics;
    private readonly CancellationToken _cancellationToken;

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

        foreach (var section in sections)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            RenderSection(package, section);
        }

        if (package.Pages.Count == 0)
        {
            package.Pages.Add(CreatePage(package, new BuiltInSectionModel()));
        }

        EnsureFontResources(package);
        return package;
    }

    private void RenderSection(OfdDocumentPackage package, BuiltInSectionModel section)
    {
        var state = new LayoutState(section);
        EnsurePage(package, state);

        foreach (var header in section.Headers)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            RenderParagraph(package, state, header, state.ContentLeft, state.ContentWidth);
        }

        foreach (var block in section.Blocks)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            switch (block)
            {
                case BuiltInParagraphModel paragraph:
                    RenderParagraph(package, state, paragraph, state.ContentLeft, state.ContentWidth);
                    break;
                case BuiltInTableModel table:
                    RenderTable(package, state, table);
                    break;
            }
        }

        foreach (var footer in section.Footers)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            RenderParagraph(package, state, footer, state.ContentLeft, state.ContentWidth);
        }
    }

    private void RenderParagraph(
        OfdDocumentPackage package,
        LayoutState state,
        BuiltInParagraphModel paragraph,
        double left,
        double availableWidth)
    {
        if (paragraph.Format.PageBreakBefore)
        {
            StartNewPage(package, state);
        }

        var spaceBefore = PointsToMillimeters(paragraph.Format.SpaceBeforePoints ?? 0);
        var spaceAfter = PointsToMillimeters(paragraph.Format.SpaceAfterPoints ?? 0);
        EnsureVerticalSpace(package, state, spaceBefore);
        state.Y += spaceBefore;

        var marker = paragraph.Format.ListMarker;
        var text = BuildParagraphText(paragraph);
        if (!string.IsNullOrEmpty(marker))
        {
            text = marker + text;
        }

        var fontSizePoints = ResolveParagraphFontSize(paragraph);
        var fontSizeMm = PointsToMillimeters(fontSizePoints);
        var lineHeight = Math.Max(
            PointsToMillimeters(paragraph.Format.LineSpacingPoints ?? (fontSizePoints * 1.3d)),
            fontSizeMm * 1.2d);
        var fontName = ResolveParagraphFont(paragraph);
        var color = ParseColor(ResolveParagraphColor(paragraph));

        var leftIndent = PointsToMillimeters(paragraph.Format.LeftIndentPoints ?? 0);
        var rightIndent = PointsToMillimeters(paragraph.Format.RightIndentPoints ?? 0);
        var firstLineIndent = PointsToMillimeters(paragraph.Format.FirstLineIndentPoints ?? 0);
        var contentLeft = left + leftIndent;
        var contentWidth = Math.Max(availableWidth - leftIndent - rightIndent, 5d);

        if (ContainsPageBreak(paragraph))
        {
            // Emit text before the break, then start a new page for remainder.
            var parts = SplitByPageBreak(paragraph);
            foreach (var (partText, isBreak) in parts)
            {
                if (!string.IsNullOrEmpty(partText))
                {
                    EmitWrappedText(
                        package,
                        state,
                        partText,
                        contentLeft,
                        contentWidth,
                        firstLineIndent,
                        fontName,
                        fontSizeMm,
                        lineHeight,
                        color,
                        paragraph.Format.Alignment);
                    firstLineIndent = 0;
                }

                if (isBreak)
                {
                    StartNewPage(package, state);
                }
            }
        }
        else if (!string.IsNullOrEmpty(text) || paragraph.Inlines.OfType<BuiltInImageModel>().Any())
        {
            if (string.IsNullOrEmpty(text))
            {
                text = "[image]";
            }

            EmitWrappedText(
                package,
                state,
                text,
                contentLeft,
                contentWidth,
                firstLineIndent,
                fontName,
                fontSizeMm,
                lineHeight,
                color,
                paragraph.Format.Alignment);
        }
        else
        {
            EnsureVerticalSpace(package, state, lineHeight);
            state.Y += lineHeight;
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

                if (table.HasBorders)
                {
                    AddRectangle(state.Page!, x, rowTop, cellWidth, rowHeight);
                }

                var textLeft = x + MinCellPaddingMillimeters;
                var textWidth = Math.Max(cellWidth - (MinCellPaddingMillimeters * 2), 1d);
                var contentHeight = cellLayout.Lines.Sum(line => line.LineHeight);
                var localY = AlignBlockVertically(
                    rowTop,
                    rowHeight,
                    contentHeight,
                    cellLayout.VerticalAlignment);

                foreach (var line in cellLayout.Lines)
                {
                    var lineWidth = MeasureLineWidth(line.Text, line.FontSizeMillimeters);
                    var lineX = AlignHorizontally(textLeft, textWidth, lineWidth, line.Alignment);
                    state.Page!.Elements.Add(CreateTextElement(
                        line.Text,
                        lineX,
                        localY,
                        Math.Max(lineWidth, 1d),
                        line.LineHeight,
                        line.FontName,
                        line.FontSizeMillimeters,
                        line.Color));
                    localY += line.LineHeight;
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
            var lines = new List<TextLine>();
            foreach (var paragraph in cell.Paragraphs)
            {
                var text = BuildParagraphText(paragraph);
                if (string.IsNullOrEmpty(text))
                {
                    continue;
                }

                var fontSizePoints = ResolveParagraphFontSize(paragraph);
                var fontSizeMm = PointsToMillimeters(fontSizePoints);
                var lineHeight = fontSizeMm * 1.25d;
                var fontName = ResolveParagraphFont(paragraph);
                var color = ParseColor(ResolveParagraphColor(paragraph));
                var alignment = paragraph.Format.Alignment;
                foreach (var wrapped in WrapText(text, textWidth, fontSizeMm))
                {
                    lines.Add(new TextLine(wrapped, fontName, fontSizeMm, lineHeight, color, alignment));
                }
            }

            if (lines.Count == 0)
            {
                var fallbackSize = PointsToMillimeters(DefaultFontSizePoints);
                lines.Add(new TextLine(
                    string.Empty,
                    ResolveFallbackFont(),
                    fallbackSize,
                    fallbackSize * 1.25d,
                    OfdColor.Black,
                    BuiltInParagraphAlignment.Left));
            }

            var height = lines.Sum(line => line.LineHeight) + (MinCellPaddingMillimeters * 2);
            layouts.Add(new CellLayout(span, height, lines, cell.VerticalAlignment));
            columnIndex += span;
        }

        return layouts;
    }

    private void EmitWrappedText(
        OfdDocumentPackage package,
        LayoutState state,
        string text,
        double left,
        double availableWidth,
        double firstLineIndent,
        string fontName,
        double fontSizeMm,
        double lineHeight,
        OfdColor color,
        BuiltInParagraphAlignment alignment)
    {
        var first = true;
        foreach (var line in WrapText(text, Math.Max(availableWidth - (first ? firstLineIndent : 0), 5d), fontSizeMm))
        {
            EnsureVerticalSpace(package, state, lineHeight);
            var indent = first ? firstLineIndent : 0;
            var width = Math.Max(availableWidth - indent, 1d);
            var lineWidth = MeasureLineWidth(line, fontSizeMm);
            var x = AlignHorizontally(left + indent, width, lineWidth, alignment);
            state.Page!.Elements.Add(CreateTextElement(
                line,
                x,
                state.Y,
                Math.Max(lineWidth, 1d),
                lineHeight,
                fontName,
                fontSizeMm,
                color));
            state.Y += lineHeight;
            first = false;
        }
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

    private static double MeasureLineWidth(string text, double fontSizeMm)
    {
        var width = 0d;
        foreach (var glyph in EnumerateTextElements(text ?? string.Empty))
        {
            width += EstimateTextWidth(glyph, fontSizeMm);
        }

        return width;
    }

    private static OfdTextElement CreateTextElement(
        string text,
        double x,
        double y,
        double width,
        double height,
        string fontName,
        double fontSizeMm,
        OfdColor color)
    {
        var element = new OfdTextElement
        {
            LayerType = "Body",
            XMillimeters = x,
            YMillimeters = y,
            WidthMillimeters = width,
            HeightMillimeters = height,
            FontName = fontName,
            FontSizeMillimeters = fontSizeMm,
            FillColor = color,
            Text = text ?? string.Empty
        };

        var glyphs = EnumerateTextElements(element.Text);
        if (glyphs.Count == 0)
        {
            return element;
        }

        string? deltaX = null;
        if (glyphs.Count > 1)
        {
            var deltas = new string[glyphs.Count - 1];
            for (var i = 0; i < deltas.Length; i++)
            {
                deltas[i] = EstimateTextWidth(glyphs[i], fontSizeMm)
                    .ToString("0.###", CultureInfo.InvariantCulture);
            }

            deltaX = string.Join(" ", deltas);
        }

        element.Runs.Add(new OfdTextRun
        {
            Text = element.Text,
            XMillimeters = 0,
            YMillimeters = Math.Max(fontSizeMm, 1d),
            DeltaX = deltaX
        });
        return element;
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

    private static IEnumerable<string> WrapText(string text, double availableWidth, double fontSizeMm)
    {
        var normalized = (text ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace("\r", "\n");

        foreach (var paragraphLine in normalized.Split('\n'))
        {
            if (paragraphLine.Length == 0)
            {
                yield return string.Empty;
                continue;
            }

            var line = new StringBuilder();
            var width = 0d;
            var enumerator = StringInfo.GetTextElementEnumerator(paragraphLine);
            while (enumerator.MoveNext())
            {
                var element = enumerator.GetTextElement();
                var elementWidth = EstimateTextWidth(element, fontSizeMm);
                if (line.Length > 0 && width + elementWidth > availableWidth)
                {
                    yield return line.ToString();
                    line.Clear();
                    width = 0;
                }

                line.Append(element);
                width += elementWidth;
            }

            yield return line.ToString();
        }
    }

    private static double EstimateTextWidth(string textElement, double fontSize)
    {
        if (string.IsNullOrEmpty(textElement))
        {
            return fontSize * AsciiWidthFactor;
        }

        if (textElement == "\t")
        {
            return fontSize * 2d;
        }

        // OFD/ofdrw convention: printable ASCII is half-em; CJK and other
        // full-width glyphs are one em. Emitting matching DeltaX keeps viewer
        // advances aligned with wrap decisions (avoids Latin monospace fallback).
        if (textElement.Length == 1)
        {
            var ch = textElement[0];
            if (ch <= 0x7f)
            {
                return fontSize * AsciiWidthFactor;
            }

            // Halfwidth / fullwidth forms
            if (ch is >= '\uFF61' and <= '\uFF9F')
            {
                return fontSize * AsciiWidthFactor;
            }
        }

        return fontSize * CjkWidthFactor;
    }

    private static string BuildParagraphText(BuiltInParagraphModel paragraph)
    {
        var builder = new StringBuilder();
        foreach (var inline in paragraph.Inlines)
        {
            switch (inline)
            {
                case BuiltInTextModel text:
                    builder.Append(text.Text);
                    break;
                case BuiltInTabModel:
                    builder.Append('\t');
                    break;
                case BuiltInBreakModel { IsPageBreak: false }:
                    builder.Append('\n');
                    break;
                case BuiltInImageModel:
                    builder.Append("[image]");
                    break;
                case BuiltInPageNumberModel:
                    builder.Append('1');
                    break;
            }
        }

        return builder.ToString();
    }

    private static bool ContainsPageBreak(BuiltInParagraphModel paragraph)
    {
        return paragraph.Inlines.OfType<BuiltInBreakModel>().Any(breakModel => breakModel.IsPageBreak);
    }

    private static IEnumerable<(string Text, bool IsBreak)> SplitByPageBreak(BuiltInParagraphModel paragraph)
    {
        var buffer = new StringBuilder();
        foreach (var inline in paragraph.Inlines)
        {
            switch (inline)
            {
                case BuiltInTextModel text:
                    buffer.Append(text.Text);
                    break;
                case BuiltInTabModel:
                    buffer.Append('\t');
                    break;
                case BuiltInBreakModel { IsPageBreak: true }:
                    yield return (buffer.ToString(), true);
                    buffer.Clear();
                    break;
                case BuiltInBreakModel:
                    buffer.Append('\n');
                    break;
                case BuiltInImageModel:
                    buffer.Append("[image]");
                    break;
                case BuiltInPageNumberModel:
                    buffer.Append('1');
                    break;
            }
        }

        yield return (buffer.ToString(), false);
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

    private string ResolveParagraphFont(BuiltInParagraphModel paragraph)
    {
        string? family = null;
        var bold = false;
        foreach (var text in paragraph.Inlines.OfType<BuiltInTextModel>())
        {
            bold |= text.Format.Bold;
            if (family is null && !string.IsNullOrWhiteSpace(text.Format.FontFamily))
            {
                family = text.Format.FontFamily;
            }
        }

        return NormalizeFontFamily(family, bold);
    }

    private string NormalizeFontFamily(string? family, bool bold)
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

        if (bold)
        {
            // Prefer a real bold CJK face so title/header weight is visible.
            if (resolved.Equals("SimSun", StringComparison.OrdinalIgnoreCase) ||
                resolved.Equals("宋体", StringComparison.OrdinalIgnoreCase))
            {
                return "SimHei";
            }

            if (resolved.Equals("Microsoft YaHei", StringComparison.OrdinalIgnoreCase))
            {
                return "Microsoft YaHei";
            }
        }

        return resolved;
    }

    private static double ResolveParagraphFontSize(BuiltInParagraphModel paragraph)
    {
        foreach (var text in paragraph.Inlines.OfType<BuiltInTextModel>())
        {
            if (text.Format.FontSizePoints is double size && size > 0)
            {
                return size;
            }
        }

        return DefaultFontSizePoints;
    }

    private static string? ResolveParagraphColor(BuiltInParagraphModel paragraph)
    {
        foreach (var text in paragraph.Inlines.OfType<BuiltInTextModel>())
        {
            if (!string.IsNullOrWhiteSpace(text.Format.ColorHex))
            {
                return text.Format.ColorHex;
            }
        }

        return null;
    }

    private string ResolveFallbackFont()
    {
        // Prefer locally common CJK faces for OFD viewers; Noto is often absent on Windows
        // and causes Latin glyphs to fall back to a mismatched monospace face.
        foreach (var preferred in new[] { "SimSun", "Microsoft YaHei", "宋体", "微软雅黑" })
        {
            foreach (var family in _options.FontFallbackFamilies)
            {
                if (!string.IsNullOrWhiteSpace(family) &&
                    family.Equals(preferred, StringComparison.OrdinalIgnoreCase))
                {
                    return NormalizeFontFamily(family, bold: false);
                }
            }
        }

        foreach (var family in _options.FontFallbackFamilies)
        {
            if (!string.IsNullOrWhiteSpace(family))
            {
                return NormalizeFontFamily(family, bold: false);
            }
        }

        return "SimSun";
    }

    private static void EnsureFontResources(OfdDocumentPackage package)
    {
        var declared = new HashSet<string>(
            package.Fonts.Select(font => font.FontName),
            StringComparer.OrdinalIgnoreCase);
        foreach (var fontName in package.Pages
                     .SelectMany(page => page.Elements)
                     .OfType<OfdTextElement>()
                     .Select(text => text.FontName)
                     .Where(name => !string.IsNullOrWhiteSpace(name))
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (!declared.Add(fontName))
            {
                continue;
            }

            package.Fonts.Add(new OfdFontResource
            {
                FontName = fontName,
                FamilyName = fontName,
                Charset = "unicode"
            });
        }
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

    private static void AddRectangle(OfdPage page, double x, double y, double width, double height)
    {
        page.Elements.Add(new OfdPathElement
        {
            LayerType = "Body",
            XMillimeters = x,
            YMillimeters = y,
            WidthMillimeters = width,
            HeightMillimeters = height,
            LineWidthMillimeters = BorderWidthMillimeters,
            Stroke = true,
            Fill = false,
            StrokeColor = OfdColor.Black,
            AbbreviatedData = string.Format(
                CultureInfo.InvariantCulture,
                "M {0:0.###} {1:0.###} L {2:0.###} {1:0.###} L {2:0.###} {3:0.###} L {0:0.###} {3:0.###} C",
                0d,
                0d,
                width,
                height)
        });
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
        state.Page = CreatePage(package, state.Section);
        state.Y = state.ContentTop;
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
        internal LayoutState(BuiltInSectionModel section)
        {
            Section = section;
            ContentLeft = PointsToMillimeters(section.MarginLeftPoints);
            ContentTop = PointsToMillimeters(section.MarginTopPoints);
            ContentBottom = PointsToMillimeters(section.PageHeightPoints - section.MarginBottomPoints);
            ContentWidth = PointsToMillimeters(
                section.PageWidthPoints - section.MarginLeftPoints - section.MarginRightPoints);
            Y = ContentTop;
        }

        internal BuiltInSectionModel Section { get; }

        internal OfdPage? Page { get; set; }

        internal double ContentLeft { get; }

        internal double ContentTop { get; }

        internal double ContentBottom { get; }

        internal double ContentWidth { get; }

        internal double Y { get; set; }
    }

    private sealed class CellLayout
    {
        internal CellLayout(
            int columnSpan,
            double height,
            List<TextLine> lines,
            BuiltInVerticalAlignment verticalAlignment)
        {
            ColumnSpan = columnSpan;
            Height = height;
            Lines = lines;
            VerticalAlignment = verticalAlignment;
        }

        internal int ColumnSpan { get; }

        internal double Height { get; }

        internal List<TextLine> Lines { get; }

        internal BuiltInVerticalAlignment VerticalAlignment { get; }
    }

    private sealed class TextLine
    {
        internal TextLine(
            string text,
            string fontName,
            double fontSizeMillimeters,
            double lineHeight,
            OfdColor color,
            BuiltInParagraphAlignment alignment)
        {
            Text = text;
            FontName = fontName;
            FontSizeMillimeters = fontSizeMillimeters;
            LineHeight = lineHeight;
            Color = color;
            Alignment = alignment;
        }

        internal string Text { get; }

        internal string FontName { get; }

        internal double FontSizeMillimeters { get; }

        internal double LineHeight { get; }

        internal OfdColor Color { get; }

        internal BuiltInParagraphAlignment Alignment { get; }
    }
}
