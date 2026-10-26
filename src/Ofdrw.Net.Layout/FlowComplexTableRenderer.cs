using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Layout.Internal.Flow;

namespace Ofdrw.Net.Layout;

internal sealed partial class FlowDocumentRenderer
{
    private sealed class ComplexCell
    {
        internal Cell Source = null!;
        internal int Row;
        internal int Column;
        internal int ColumnSpan;
        internal int RowSpan;
        internal double Width;
        internal FlowCellContent? Text;
        internal ImageBlock? Image;
        internal byte[]? ImageBytes;
        internal double ContentHeight;
        internal double RequiredHeight => ContentHeight + 2 * Source.PaddingMillimeters;
    }

    private sealed class ComplexTable
    {
        internal Table Source = null!;
        internal Row[] Rows = null!;
        internal double[] Columns = null!;
        internal double[] Edges = null!;
        internal double[] Heights = null!;
        internal ComplexCell[][] Grid = null!;
        internal List<ComplexCell>[] Origins = null!;
        internal double Left;
        internal double Width;
        internal int HeaderRows;
        internal double HeaderHeight;
    }

    private readonly struct RowGroup
    {
        internal RowGroup(int first, int last, double height)
        { First = first; Last = last; Height = height; }
        internal int First { get; }
        internal int Last { get; }
        internal double Height { get; }
    }

    private readonly struct RowSlice
    {
        internal RowSlice(double start, double end, double leading, double trailing)
        { Start = start; End = end; Leading = leading; Trailing = trailing; }
        internal double Start { get; }
        internal double End { get; }
        internal double Leading { get; }
        internal double Trailing { get; }
        internal double Height => End - Start + Leading + Trailing;
        internal bool IsFinal(double rowHeight) => End >= rowHeight - 0.000001d;
    }

    private void RenderComplexTable(Table table)
    {
        var layout = MeasureComplexTable(table);
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

        var groups = BuildBodyGroups(layout);
        var first = groups[0];
        var firstCanSplit = first.First == first.Last && table.SplitRowsAcrossPages &&
            layout.Origins[first.First].All(cell => cell.Source.VerticalAlignment == CellVerticalAlignment.Top);
        var firstNeeded = firstCanSplit
            ? MinimumSliceHeight(layout, first.First) : first.Height;
        var freshPageHeight = ContentBottom - ContentTop - layout.HeaderHeight;
        if (firstNeeded > freshPageHeight + 0.000001d)
        {
            if (table.SplitRowsAcrossPages && first.First == first.Last && !firstCanSplit)
                throw new NotSupportedException("A split row needs Top vertical alignment in every cell.");
            throw new InvalidOperationException("Table header and first body fragment cannot fit on a fresh page.");
        }
        PlaceFirstLine(layout.HeaderHeight + firstNeeded, _pendingSpaceAfter + table.SpaceBeforeMillimeters);
        _pendingSpaceAfter = 0;
        var hasTableRowOnPage = false;
        if (layout.HeaderRows > 0)
        {
            DrawComplexRows(layout, 0, layout.HeaderRows - 1, true);
            hasTableRowOnPage = true;
        }
        foreach (var group in groups)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (group.First == group.Last && table.SplitRowsAcrossPages)
            {
                DrawSplitRow(layout, group.First, ref hasTableRowOnPage);
                continue;
            }
            if (group.Height > ContentBottom - ContentTop - layout.HeaderHeight + 0.000001d)
                throw new InvalidOperationException($"Table row group {group.First + 1}-{group.Last + 1} exceeds a page with its header.");
            if (_y + group.Height > ContentBottom + 0.000001d)
                BeginTableContinuation(layout, group.Height, ref hasTableRowOnPage);
            DrawComplexRows(layout, group.First, group.Last, !hasTableRowOnPage);
            hasTableRowOnPage = true;
        }
        _pendingSpaceAfter = table.SpaceAfterMillimeters;
    }

    private ComplexTable MeasureComplexTable(Table table)
    {
        ValidateLength(table.SpaceBeforeMillimeters, nameof(table.SpaceBeforeMillimeters));
        ValidateLength(table.SpaceAfterMillimeters, nameof(table.SpaceAfterMillimeters));
        ValidateLength(table.BorderWidthMillimeters, nameof(table.BorderWidthMillimeters));
        if (table.BorderColor is null) throw new ArgumentException("Table border color is required.");
        if (!Enum.IsDefined(typeof(ParagraphAlignment), table.Alignment)) throw new ArgumentOutOfRangeException(nameof(table.Alignment));
        var rows = table.Rows.ToArray();
        if (rows.Length == 0 || rows[0].Cells.Count == 0) throw new ArgumentException("A table must contain a nonempty row.");
        if (table.RepeatHeaderRowCount < 0 || table.RepeatHeaderRowCount >= rows.Length)
            throw new ArgumentOutOfRangeException(nameof(table.RepeatHeaderRowCount), "A repeated header needs at least one body row.");
        var bodyWidth = _options.PageWidthMillimeters - _options.MarginLeftMillimeters - _options.MarginRightMillimeters;
        var columns = table.ColumnWidthsMillimeters.ToArray();
        if (columns.Length == 0)
        {
            long count = 0;
            foreach (var cell in rows[0].Cells)
            {
                if (cell.ColumnSpan < 1) throw new ArgumentOutOfRangeException(nameof(cell.ColumnSpan));
                count += cell.ColumnSpan;
            }
            if (count > _options.MaxTableCells) throw new InvalidOperationException("Table grid exceeds MaxTableCells.");
            columns = Enumerable.Repeat(bodyWidth / count, (int)count).ToArray();
        }
        var gridSlots = (long)rows.Length * columns.Length;
        if (columns.Length > _options.MaxTableCells || gridSlots > _options.MaxTableGridSlots - _tableGridSlots)
            throw new InvalidOperationException("Complex table exceeds the grid-slot budget.");
        _tableGridSlots += gridSlots;
        foreach (var width in columns) ValidatePositive(width, nameof(table.ColumnWidthsMillimeters));
        var edges = new double[columns.Length + 1];
        for (var i = 0; i < columns.Length; i++) edges[i + 1] = edges[i] + columns[i];
        var tableWidth = edges[edges.Length - 1];
        if (tableWidth > bodyWidth + 0.000001d || double.IsInfinity(tableWidth))
            throw new ArgumentException("Table columns exceed the usable page width.");
        if (table.BorderWidthMillimeters >= columns.Min()) throw new ArgumentException("Table border is too wide for the grid.");
        var layout = new ComplexTable
        {
            Source = table, Rows = rows, Columns = columns, Edges = edges,
            Heights = new double[rows.Length], Grid = new ComplexCell[rows.Length][],
            Origins = new List<ComplexCell>[rows.Length], Width = tableWidth,
            Left = FlowParagraphLayout.Align(_options.MarginLeftMillimeters, bodyWidth, tableWidth, (FlowAlignment)table.Alignment),
            HeaderRows = table.RepeatHeaderRowCount
        };
        for (var row = 0; row < rows.Length; row++)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            layout.Grid[row] = new ComplexCell[columns.Length];
            layout.Origins[row] = new List<ComplexCell>();
            ValidateLength(rows[row].MinimumHeightMillimeters, nameof(Row.MinimumHeightMillimeters));
            layout.Heights[row] = Math.Max(rows[row].MinimumHeightMillimeters,
                Math.Max(table.BorderWidthMillimeters, 0.1d));
        }
        var spanning = new List<ComplexCell>();
        long admittedImageBytes = 0;
        for (var row = 0; row < rows.Length; row++)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var column = 0;
            foreach (var cell in rows[row].Cells)
            {
                while (column < columns.Length && layout.Grid[row][column] is not null) column++;
                if (column == columns.Length || cell.ColumnSpan < 1 || cell.ColumnSpan > columns.Length - column)
                    throw new ArgumentException($"Table row {row + 1} has a cell span outside its grid.");
                if (cell.RowSpan < 1 || cell.RowSpan > rows.Length - row)
                    throw new ArgumentOutOfRangeException(nameof(cell.RowSpan), $"Table row {row + 1} has an invalid RowSpan.");
                if (row < layout.HeaderRows && row + cell.RowSpan > layout.HeaderRows)
                    throw new ArgumentException("A header cell cannot span into body rows.");
                if (!Enum.IsDefined(typeof(CellVerticalAlignment), cell.VerticalAlignment))
                    throw new ArgumentOutOfRangeException(nameof(cell.VerticalAlignment));
                ValidateLength(cell.PaddingMillimeters, nameof(cell.PaddingMillimeters));
                var width = edges[column + cell.ColumnSpan] - edges[column];
                if (width <= 2 * cell.PaddingMillimeters)
                    throw new ArgumentException("Cell padding leaves no usable content width.");
                var measured = new ComplexCell { Source = cell, Row = row, Column = column,
                    ColumnSpan = cell.ColumnSpan, RowSpan = cell.RowSpan, Width = width };
                if (cell.Image is { } image)
                {
                    if (cell.Paragraphs.Count != 0) throw new NotSupportedException("An image cell cannot also contain paragraphs.");
                    if (image.PageBreakBefore || image.SpaceBeforeMillimeters != 0 || image.SpaceAfterMillimeters != 0)
                        throw new NotSupportedException("A cell image cannot request block page breaks or spacing.");
                    if (!Enum.IsDefined(typeof(ParagraphAlignment), image.Alignment))
                        throw new ArgumentOutOfRangeException(nameof(image.Alignment));
                    ValidatePositive(image.WidthMillimeters, nameof(image.WidthMillimeters));
                    ValidatePositive(image.HeightMillimeters, nameof(image.HeightMillimeters));
                    if (image.WidthMillimeters > width - 2 * cell.PaddingMillimeters + 0.000001d)
                        throw new ArgumentException("Cell image exceeds its padded width.");
                    if (image.Alpha is < 0 or > 255) throw new ArgumentOutOfRangeException(nameof(image.Alpha));
                    var bytes = image.SnapshotImage();
                    var limits = new OfdGraphicsOptions();
                    if (bytes.Length == 0 || bytes.Length > limits.MaxEncodedImageBytes ||
                        bytes.Length > _options.MaxTableImageBytes - _tableImageBytes - admittedImageBytes)
                        throw new ArgumentException("Cell image exceeds the encoded-image budget or is empty.");
                    admittedImageBytes += bytes.Length;
                    OfdGraphicsImage.Validate(bytes, limits.MaxImagePixels, _cancellationToken);
                    measured.Image = image;
                    measured.ImageBytes = bytes;
                    measured.ContentHeight = image.HeightMillimeters;
                }
                else
                {
                    measured.Text = MeasureCell(cell, width - 2 * cell.PaddingMillimeters);
                    measured.ContentHeight = measured.Text.Height;
                }
                for (var coveredRow = row; coveredRow < row + cell.RowSpan; coveredRow++)
                    for (var coveredColumn = column; coveredColumn < column + cell.ColumnSpan; coveredColumn++)
                    {
                        if (layout.Grid[coveredRow][coveredColumn] is not null)
                            throw new ArgumentException($"Table row {row + 1} overlaps another cell.");
                        layout.Grid[coveredRow][coveredColumn] = measured;
                    }
                layout.Origins[row].Add(measured);
                if (cell.RowSpan == 1) layout.Heights[row] = Math.Max(layout.Heights[row], measured.RequiredHeight);
                else spanning.Add(measured);
                column += cell.ColumnSpan;
            }
            if (layout.Grid[row].Any(slot => slot is null))
                throw new ArgumentException($"Table row {row + 1} does not cover every column.");
        }
        foreach (var cell in spanning.OrderBy(value => value.Row + value.RowSpan))
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var height = SumHeights(layout, cell.Row, cell.Row + cell.RowSpan - 1);
            if (cell.RequiredHeight > height) layout.Heights[cell.Row + cell.RowSpan - 1] += cell.RequiredHeight - height;
        }
        for (var row = 0; row < rows.Length; row++)
        {
            if (double.IsInfinity(layout.Heights[row]) || layout.Heights[row] <= 0)
                throw new ArgumentException($"Table row {row + 1} has invalid measured height.");
        }
        layout.HeaderHeight = layout.HeaderRows == 0 ? 0 : SumHeights(layout, 0, layout.HeaderRows - 1);
        return layout;
    }

    private static double SumHeights(ComplexTable layout, int first, int last)
    {
        var height = 0d;
        for (var row = first; row <= last; row++) height += layout.Heights[row];
        return height;
    }

    private List<RowGroup> BuildBodyGroups(ComplexTable layout)
    {
        var result = new List<RowGroup>();
        var next = layout.HeaderRows;
        while (next < layout.Rows.Length)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var last = next;
            for (var row = next; row <= last; row++)
                foreach (var cell in layout.Origins[row]) last = Math.Max(last, cell.Row + cell.RowSpan - 1);
            result.Add(new RowGroup(next, last, SumHeights(layout, next, last)));
            next = last + 1;
        }
        return result;
    }

    private void BeginTableContinuation(ComplexTable layout, double neededBodyHeight, ref bool hasTableRowOnPage)
    {
        if (layout.HeaderHeight + neededBodyHeight > ContentBottom - ContentTop + 0.000001d)
            throw new InvalidOperationException("Table header and next body fragment cannot fit on a fresh page.");
        StartPage();
        hasTableRowOnPage = false;
        if (layout.HeaderRows > 0)
        {
            DrawComplexRows(layout, 0, layout.HeaderRows - 1, true);
            hasTableRowOnPage = true;
        }
    }

    private void DrawComplexRows(ComplexTable layout, int first, int last, bool topBorder)
    {
        for (var row = first; row <= last; row++)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            CountTableFragment();
            var top = _y;
            var height = layout.Heights[row];
            foreach (var cell in layout.Origins[row])
            {
                var spanHeight = SumHeights(layout, row, row + cell.RowSpan - 1);
                DrawComplexFill(layout, cell, top, spanHeight);
            }
            foreach (var cell in layout.Origins[row])
            {
                var spanHeight = SumHeights(layout, row, row + cell.RowSpan - 1);
                DrawComplexContent(layout, cell, top, spanHeight);
            }
            DrawComplexGridRow(layout, row, top, height, topBorder && row == first, true);
            if (_onElement is not null)
                RecordPlacement(layout.Source, new FlowBoundary(_page!.Index, layout.Left, top, layout.Width, height));
            _y += height;
            _hasBodyContent = true;
        }
    }

    private void DrawComplexFill(ComplexTable layout, ComplexCell cell, double y, double height)
    {
        if (cell.Source.BackgroundColor is not { } fill) return;
        var x = layout.Left + layout.Edges[cell.Column];
        _page!.Elements.Add(new OfdPathElement { LayerType = "Body", XMillimeters = x, YMillimeters = y,
            WidthMillimeters = cell.Width, HeightMillimeters = height, Fill = true, Stroke = false,
            FillColor = fill, AbbreviatedData = FormattableString.Invariant(
                $"M 0 0 L {cell.Width} 0 L {cell.Width} {height} L 0 {height} C") });
    }

    private void DrawComplexContent(ComplexTable layout, ComplexCell cell, double y, double height)
    {
        var padding = cell.Source.PaddingMillimeters;
        var extra = Math.Max(0, height - cell.RequiredHeight);
        var vertical = cell.Source.VerticalAlignment == CellVerticalAlignment.Center ? extra / 2 :
            cell.Source.VerticalAlignment == CellVerticalAlignment.Bottom ? extra : 0;
        var top = y + padding + vertical;
        var left = layout.Left + layout.Edges[cell.Column];
        if (cell.Image is { } image)
        {
            var imageX = FlowParagraphLayout.Align(left + padding, cell.Width - 2 * padding,
                image.WidthMillimeters, (FlowAlignment)image.Alignment);
            DrawTableImage(cell.ImageBytes!, imageX, top, image.WidthMillimeters, image.HeightMillimeters, image.Alpha);
        }
        else foreach (var line in cell.Text!.Lines)
            DrawLine(line.Line, left + padding + line.Left, line.Width, top + line.Top);
    }

    private void DrawTableImage(byte[] imageBytes, double x, double y, double width, double height, int alpha)
    {
        if (imageBytes.Length > _options.MaxTableImageBytes - _tableImageBytes)
            throw new InvalidOperationException("Table images exceed MaxTableImageBytes.");
        _tableImageBytes += imageBytes.Length;
        new OfdGraphics(_package, _page!).DrawImage(imageBytes, x, y, width, height, alpha, _cancellationToken);
    }

    private void CountTableFragment()
    {
        if (_tableFragments >= _options.MaxTableFragments)
            throw new InvalidOperationException("Flow tables exceed MaxTableFragments.");
        _tableFragments++;
    }

    private void DrawComplexGridRow(ComplexTable layout, int row, double y, double height, bool topBorder, bool bottomBorder)
    {
        var stroke = layout.Source.BorderWidthMillimeters;
        if (stroke == 0) return;
        var inset = stroke / 2;
        void Edge(double x1, double y1, double x2, double y2) =>
            _page!.Elements.Add(new OfdPathElement { LayerType = "Body", XMillimeters = layout.Left, YMillimeters = y,
                WidthMillimeters = layout.Width, HeightMillimeters = height, Fill = false, Stroke = true,
                StrokeColor = layout.Source.BorderColor, LineWidthMillimeters = stroke,
                AbbreviatedData = FormattableString.Invariant($"M {x1} {y1} L {x2} {y2}") });
        if (topBorder) Edge(inset, inset, layout.Width - inset, inset);
        if (bottomBorder)
        {
            var start = -1;
            for (var column = 0; column <= layout.Columns.Length; column++)
            {
                var draw = column < layout.Columns.Length &&
                    (row == layout.Rows.Length - 1 ||
                     !ReferenceEquals(layout.Grid[row][column], layout.Grid[row + 1][column]));
                if (draw && start < 0) start = column;
                if (!draw && start >= 0)
                {
                    var from = start == 0 ? inset : layout.Edges[start];
                    var to = column == layout.Columns.Length ? layout.Width - inset : layout.Edges[column];
                    if (to > from) Edge(from, height - inset, to, height - inset);
                    start = -1;
                }
            }
        }
        Edge(inset, 0, inset, height);
        for (var column = 1; column < layout.Columns.Length; column++)
            if (!ReferenceEquals(layout.Grid[row][column - 1], layout.Grid[row][column]))
                Edge(layout.Edges[column], 0, layout.Edges[column], height);
        Edge(layout.Width - inset, 0, layout.Width - inset, height);
    }

    private double MinimumSliceHeight(ComplexTable layout, int row)
    {
        var height = layout.Heights[row];
        var gap = SplitGap(layout, row);
        var cuts = SafeSplitCuts(SplitItems(layout, row), height);
        return cuts.Count == 0 ? height : Math.Min(height, cuts[0] + gap);
    }

    private static double SplitGap(ComplexTable layout, int row) =>
        Math.Max(Math.Max(layout.Source.BorderWidthMillimeters, 0.1d),
            layout.Origins[row].Max(cell => cell.Source.PaddingMillimeters));

    private static List<(double Start, double End)> SplitItems(ComplexTable layout, int row)
    {
        var items = new List<(double Start, double End)>();
        foreach (var cell in layout.Origins[row])
        {
            var padding = cell.Source.PaddingMillimeters;
            if (cell.Image is { } image) items.Add((padding, padding + image.HeightMillimeters));
            else foreach (var line in cell.Text!.Lines)
                items.Add((padding + line.Top, padding + line.Top + line.Line.Height));
        }
        return items;
    }

    private static List<double> SafeSplitCuts(IReadOnlyList<(double Start, double End)> items, double height)
    {
        var starts = items.OrderBy(item => item.Start).ToArray();
        var ends = items.Select(item => item.End).Distinct().OrderBy(value => value);
        var result = new List<double>();
        var cursor = 0;
        var activeEnd = 0d;
        var finalItemEnd = items.Count == 0 ? 0 : items.Max(item => item.End);
        foreach (var cut in ends)
        {
            while (cursor < starts.Length && starts[cursor].Start < cut - 0.000001d)
            {
                activeEnd = Math.Max(activeEnd, starts[cursor].End);
                cursor++;
            }
            if (activeEnd <= cut + 0.000001d && cut < finalItemEnd - 0.000001d &&
                cut < height - 0.000001d)
                result.Add(cut);
        }
        return result;
    }

    private RowSlice? FindRowSlice(ComplexTable layout, int row, double start, double available,
        List<double> safeCuts)
    {
        var height = layout.Heights[row];
        var gap = SplitGap(layout, row);
        var leading = start > 0 ? gap : 0;
        if (height - start + leading <= available + 0.000001d)
            return new RowSlice(start, height, leading, 0);
        var limit = start + available - leading - gap;
        var index = safeCuts.BinarySearch(limit + 0.000001d);
        if (index < 0) index = ~index - 1;
        if (index < 0 || safeCuts[index] <= start + 0.000001d) return null;
        return new RowSlice(start, safeCuts[index], leading, gap);
    }

    private void DrawSplitRow(ComplexTable layout, int row, ref bool hasTableRowOnPage)
    {
        var items = SplitItems(layout, row);
        var cuts = SafeSplitCuts(items, layout.Heights[row]);
        var hasNonTop = layout.Origins[row].Any(cell => cell.Source.VerticalAlignment != CellVerticalAlignment.Top);
        var start = 0d;
        while (start < layout.Heights[row] - 0.000001d)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var available = ContentBottom - _y;
            var slice = FindRowSlice(layout, row, start, available, cuts);
            if (slice is not null && !slice.Value.IsFinal(layout.Heights[row]) && hasNonTop)
            {
                if (layout.Heights[row] > ContentBottom - ContentTop - layout.HeaderHeight + 0.000001d)
                    throw new NotSupportedException("A split row needs Top vertical alignment in every cell.");
                slice = null; // Preserve the intact row by moving it to a fresh page.
            }
            if (slice is null)
            {
                var fresh = FindRowSlice(layout, row, start,
                    ContentBottom - ContentTop - layout.HeaderHeight, cuts);
                if (fresh is null)
                    throw new InvalidOperationException($"Table row {row + 1} has no safe progressing page fragment.");
                if (!fresh.Value.IsFinal(layout.Heights[row]) && hasNonTop)
                    throw new NotSupportedException("A split row needs Top vertical alignment in every cell.");
                BeginTableContinuation(layout, fresh.Value.Height, ref hasTableRowOnPage);
                slice = fresh;
            }
            CountTableFragment();
            DrawRowSlice(layout, row, slice.Value, !hasTableRowOnPage);
            hasTableRowOnPage = true;
            start = slice.Value.End;
            if (slice.Value.IsFinal(layout.Heights[row])) break;
        }
    }

    private void DrawRowSlice(ComplexTable layout, int row, RowSlice slice, bool topBorder)
    {
        var top = _y;
        foreach (var cell in layout.Origins[row]) DrawComplexFill(layout, cell, top, slice.Height);
        foreach (var cell in layout.Origins[row])
        {
            var left = layout.Left + layout.Edges[cell.Column];
            var padding = cell.Source.PaddingMillimeters;
            var extra = slice.Start == 0 && slice.IsFinal(layout.Heights[row])
                ? Math.Max(0, layout.Heights[row] - cell.RequiredHeight) : 0;
            var vertical = cell.Source.VerticalAlignment == CellVerticalAlignment.Center ? extra / 2 :
                cell.Source.VerticalAlignment == CellVerticalAlignment.Bottom ? extra : 0;
            if (cell.Image is { } image)
            {
                var imageTop = padding + vertical;
                if (imageTop >= slice.Start - 0.000001d &&
                    imageTop + image.HeightMillimeters <= slice.End + 0.000001d)
                {
                    var imageX = FlowParagraphLayout.Align(left + padding, cell.Width - 2 * padding,
                        image.WidthMillimeters, (FlowAlignment)image.Alignment);
                    DrawTableImage(cell.ImageBytes!, imageX, top + slice.Leading + imageTop - slice.Start,
                        image.WidthMillimeters, image.HeightMillimeters, image.Alpha);
                }
            }
            else foreach (var line in cell.Text!.Lines)
            {
                var lineTop = padding + vertical + line.Top;
                if (lineTop < slice.Start - 0.000001d ||
                    lineTop + line.Line.Height > slice.End + 0.000001d) continue;
                DrawLine(line.Line, left + padding + line.Left, line.Width,
                    top + slice.Leading + lineTop - slice.Start);
            }
        }
        DrawComplexGridRow(layout, row, top, slice.Height, topBorder, true);
        if (_onElement is not null)
            RecordPlacement(layout.Source, new FlowBoundary(_page!.Index, layout.Left, top, layout.Width, slice.Height));
        _y += slice.Height;
        _hasBodyContent = true;
    }
}
