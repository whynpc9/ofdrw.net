using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Layout.Internal.Flow;

internal sealed class FlowCellSpec
{
    internal int ColumnSpan { get; set; } = 1;
    internal double Padding { get; set; }
    internal Func<double, FlowCellContent> Measure { get; set; } = null!;
}

internal sealed class FlowCellContent
{
    internal List<FlowCellLine> Lines { get; } = new();
    internal double Height { get; set; }
}

internal sealed class FlowCellLine
{
    internal FlowCellLine(FlowLine line, double top, double left, double width)
    { Line = line; Top = top; Left = left; Width = width; }
    internal FlowLine Line { get; }
    internal double Top { get; }
    internal double Left { get; }
    internal double Width { get; }
}

internal sealed class FlowMeasuredCell
{
    internal int ColumnSpan { get; set; }
    internal double Width { get; set; }
    internal double Padding { get; set; }
    internal FlowCellContent Content { get; set; } = null!;
    internal double Height => Content.Height + 2 * Padding;
}

/// <summary>Pure row geometry shared by public tables and DOCX Native. Adapters provide formatted line measurement.</summary>
internal static class FlowTableLayout
{
    internal static List<FlowMeasuredCell> MeasureRow(IReadOnlyList<double> columns,
        IEnumerable<FlowCellSpec> cells, CancellationToken cancellationToken)
    {
        var result = new List<FlowMeasuredCell>();
        var column = 0;
        foreach (var cell in cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.ColumnSpan < 1 || cell.ColumnSpan > columns.Count - column)
                throw new ArgumentException("A cell span exceeds the table grid.");
            var width = 0d;
            for (var i = 0; i < cell.ColumnSpan; i++) width += columns[column + i];
            if (double.IsNaN(cell.Padding) || double.IsInfinity(cell.Padding) || cell.Padding < 0 ||
                width <= 2 * cell.Padding)
                throw new ArgumentException("Cell padding leaves no usable text width.");
            var content = cell.Measure(width - 2 * cell.Padding);
            if (double.IsNaN(content.Height) || double.IsInfinity(content.Height) || content.Height < 0)
                throw new ArgumentException("Invalid measured cell height.");
            result.Add(new FlowMeasuredCell { ColumnSpan = cell.ColumnSpan, Width = width,
                Padding = cell.Padding, Content = content });
            column += cell.ColumnSpan;
        }
        if (column != columns.Count) throw new ArgumentException("A table row must cover every column exactly once.");
        return result;
    }

    internal static double VerticalOffset(double rowHeight, FlowMeasuredCell cell, int alignment)
    {
        var extra = Math.Max(0, rowHeight - cell.Height);
        return cell.Padding + (alignment == 1 ? extra / 2 : alignment == 2 ? extra : 0);
    }
}
