using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Internal.Flow;

namespace Ofdrw.Net.Layout;

internal sealed partial class FlowDocumentRenderer
{
    private void RenderTable(Table table)
    {
        ValidateLength(table.SpaceBeforeMillimeters, nameof(table.SpaceBeforeMillimeters));
        ValidateLength(table.SpaceAfterMillimeters, nameof(table.SpaceAfterMillimeters));
        ValidateLength(table.BorderWidthMillimeters, nameof(table.BorderWidthMillimeters));
        if (table.BorderColor is null) throw new ArgumentException("Table border color is required.");
        if (!Enum.IsDefined(typeof(ParagraphAlignment), table.Alignment)) throw new ArgumentOutOfRangeException(nameof(table.Alignment));
        var rows = table.Rows.ToArray();
        if (rows.Length == 0 || rows[0].Cells.Count == 0) throw new ArgumentException("A table must contain a nonempty row.");
        var bodyWidth = _options.PageWidthMillimeters - _options.MarginLeftMillimeters - _options.MarginRightMillimeters;
        var columns = table.ColumnWidthsMillimeters.ToArray();
        if (columns.Length == 0)
        {
            var count = 0L;
            foreach (var cell in rows[0].Cells)
            {
                if (cell.ColumnSpan < 1) throw new ArgumentOutOfRangeException(nameof(cell.ColumnSpan));
                count += cell.ColumnSpan;
            }
            if (count > _options.MaxTableCells) throw new InvalidOperationException("Table grid exceeds MaxTableCells.");
            columns = Enumerable.Repeat(bodyWidth / count, (int)count).ToArray();
        }
        if (columns.Length > _options.MaxTableCells) throw new InvalidOperationException("Table grid exceeds MaxTableCells.");
        foreach (var width in columns) ValidatePositive(width, nameof(table.ColumnWidthsMillimeters));
        var tableWidth = columns.Sum();
        if (tableWidth > bodyWidth + 0.000001d || double.IsInfinity(tableWidth))
            throw new ArgumentException("Table columns exceed the usable page width.");
        if (table.BorderWidthMillimeters >= columns.Min()) throw new ArgumentException("Table border is too wide for the grid.");
        var left = FlowParagraphLayout.Align(_options.MarginLeftMillimeters, bodyWidth, tableWidth, (FlowAlignment)table.Alignment);
        if (_page is null) StartPage();
        var hadPendingBreaks = _pendingPageBreaks > 0;
        ApplyPendingPageBreaks();
        if (table.PageBreakBefore && (_hasBodyContent || hadPendingBreaks || _pendingLineBreakHeights.Count > 0))
        {
            StartPage();
            _pendingSpaceAfter = 0;
            _pendingLineBreakHeights.Clear();
        }
        ApplyPendingLineBreaks();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var row = rows[rowIndex];
            ValidateLength(row.MinimumHeightMillimeters, nameof(row.MinimumHeightMillimeters));
            var cells = row.Cells.ToArray();
            var specs = new List<FlowCellSpec>();
            foreach (var cell in cells)
            {
                if (cell.RowSpan < 1) throw new ArgumentOutOfRangeException(nameof(cell.RowSpan));
                if (cell.RowSpan > 1) throw new NotSupportedException("RowSpan greater than 1 is unsupported.");
                if (!Enum.IsDefined(typeof(CellVerticalAlignment), cell.VerticalAlignment)) throw new ArgumentOutOfRangeException(nameof(cell.VerticalAlignment));
                specs.Add(new FlowCellSpec { ColumnSpan = cell.ColumnSpan, Padding = cell.PaddingMillimeters,
                    Measure = width => MeasureCell(cell, width) });
            }
            var measured = FlowTableLayout.MeasureRow(columns, specs, _cancellationToken);
            var height = Math.Max(row.MinimumHeightMillimeters, measured.Max(cell => cell.Height));
            if (height <= 0) throw new ArgumentException("A table row must have positive height.");
            if (height < table.BorderWidthMillimeters) throw new ArgumentException("A table row is thinner than its border stroke.");
            if (height > ContentBottom - ContentTop + 0.000001d)
                throw new InvalidOperationException(FormattableString.Invariant($"Table row {rowIndex + 1} height {height} mm exceeds usable page height {ContentBottom - ContentTop} mm."));
            if (rowIndex == 0)
            {
                PlaceFirstLine(height, _pendingSpaceAfter + table.SpaceBeforeMillimeters);
                _pendingSpaceAfter = 0;
            }
            else Place(height);
            var rowTop = _y;
            var x = left;
            // Fills precede text and edges, including the adjacent row's already-painted shared edge.
            for (var i = 0; i < cells.Length; i++)
            {
                if (cells[i].BackgroundColor is { } fill)
                    _page!.Elements.Add(new OfdPathElement { LayerType = "Body", XMillimeters = x, YMillimeters = rowTop,
                        WidthMillimeters = measured[i].Width, HeightMillimeters = height, Fill = true, Stroke = false,
                        FillColor = fill, AbbreviatedData = FormattableString.Invariant($"M 0 0 L {measured[i].Width} 0 L {measured[i].Width} {height} L 0 {height} C") });
                x += measured[i].Width;
            }
            x = left;
            for (var i = 0; i < cells.Length; i++)
            {
                var cell = measured[i];
                var top = rowTop + FlowTableLayout.VerticalOffset(height, cell, (int)cells[i].VerticalAlignment);
                foreach (var line in cell.Content.Lines)
                    DrawLine(line.Line, x + cell.Padding + line.Left, line.Width, top + line.Top);
                x += cell.Width;
            }
            DrawGridRow(table, measured, left, rowTop, tableWidth, height, rowIndex == 0 || Math.Abs(rowTop - ContentTop) < 0.000001d);
            _y += height;
            _hasBodyContent = true;
        }
        _pendingSpaceAfter = table.SpaceAfterMillimeters;
    }

    private FlowCellContent MeasureCell(Cell cell, double width)
    {
        var content = new FlowCellContent();
        foreach (var paragraph in cell.Paragraphs)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (paragraph.PageBreakBefore || paragraph.Spans.Any(span => span.Text.Contains('\f')))
                throw new NotSupportedException("Explicit page breaks inside a table cell are unsupported.");
            ValidateLength(paragraph.LeftIndentMillimeters, nameof(paragraph.LeftIndentMillimeters));
            ValidateLength(paragraph.RightIndentMillimeters, nameof(paragraph.RightIndentMillimeters));
            ValidateLength(paragraph.FirstLineIndentMillimeters, nameof(paragraph.FirstLineIndentMillimeters));
            ValidateLength(paragraph.SpaceBeforeMillimeters, nameof(paragraph.SpaceBeforeMillimeters));
            ValidateLength(paragraph.SpaceAfterMillimeters, nameof(paragraph.SpaceAfterMillimeters));
            if (paragraph.MinimumLineHeightMillimeters is { } min) ValidateLength(min, nameof(paragraph.MinimumLineHeightMillimeters));
            if (!Enum.IsDefined(typeof(ParagraphAlignment), paragraph.Alignment)) throw new ArgumentOutOfRangeException(nameof(paragraph.Alignment));
            var textWidth = width - paragraph.LeftIndentMillimeters - paragraph.RightIndentMillimeters;
            if (textWidth <= 0 || paragraph.FirstLineIndentMillimeters >= textWidth) throw new ArgumentException("Cell paragraph indents leave no usable width.");
            var inlines = new List<FlowInline>();
            foreach (var span in paragraph.Spans)
            {
                var size = span.FontSizeMillimeters ?? _options.DefaultFontSizeMillimeters;
                ValidatePositive(size, nameof(span.FontSizeMillimeters));
                var family = span.FontFamily ?? _options.DefaultFontFamily;
                if (string.IsNullOrWhiteSpace(family)) throw new ArgumentException("Font family cannot be empty.");
                inlines.Add(new FlowInline { Text = span.Text, Style = new FlowTextStyle { FontFamily = family.Trim(),
                    FontSizeMillimeters = size, Bold = span.Bold, Italic = span.Italic, Source = span.Color ?? _options.DefaultColor } });
            }
            var fallback = new FlowTextStyle { FontFamily = _options.DefaultFontFamily,
                FontSizeMillimeters = _options.DefaultFontSizeMillimeters, Source = _options.DefaultColor };
            content.Height += paragraph.SpaceBeforeMillimeters;
            foreach (var line in FlowParagraphLayout.Layout(inlines, new FlowParagraphFormat { Alignment = (FlowAlignment)paragraph.Alignment,
                FirstLineIndentMillimeters = paragraph.FirstLineIndentMillimeters, MinimumLineHeightMillimeters = paragraph.MinimumLineHeightMillimeters ?? 0 },
                textWidth, this, fallback, _cancellationToken))
            {
                content.Lines.Add(new FlowCellLine(line, content.Height, paragraph.LeftIndentMillimeters, textWidth));
                content.Height += line.Height;
            }
            content.Height += paragraph.SpaceAfterMillimeters;
        }
        return content;
    }

    private void DrawGridRow(Table table, IReadOnlyList<FlowMeasuredCell> cells, double x, double y, double width, double height, bool drawTop)
    {
        var stroke = table.BorderWidthMillimeters;
        if (stroke == 0) return;
        var inset = stroke / 2;
        void Edge(double x1, double y1, double x2, double y2) =>
            _page!.Elements.Add(new OfdPathElement { LayerType = "Body", XMillimeters = x, YMillimeters = y,
                WidthMillimeters = width, HeightMillimeters = height, Fill = false, Stroke = true,
                StrokeColor = table.BorderColor, LineWidthMillimeters = stroke,
                AbbreviatedData = FormattableString.Invariant($"M {x1} {y1} L {x2} {y2}") });
        if (drawTop) Edge(inset, inset, width - inset, inset);
        // The bottom stroke is inset. The next row's fill starts outside it, so it remains visible without re-stroking.
        Edge(inset, height - inset, width - inset, height - inset);
        Edge(inset, 0, inset, height);
        var cursor = 0d;
        for (var i = 0; i < cells.Count; i++)
        {
            cursor += cells[i].Width;
            var edgeX = i == cells.Count - 1 ? width - inset : cursor;
            Edge(edgeX, 0, edgeX, height);
        }
    }
}
