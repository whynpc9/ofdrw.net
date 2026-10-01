using System.Collections.Generic;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout;

/// <summary>A flowing table. Rows move intact to the next page; an oversized row fails.</summary>
public sealed class Table : FlowBlock
{
    /// <summary>Absolute column widths in millimeters. Empty means equal columns inferred from the first row.</summary>
    public IList<double> ColumnWidthsMillimeters { get; } = new List<double>();
    /// <summary>Rows in reading order. Each row must cover every column exactly once.</summary>
    public IList<Row> Rows { get; } = new List<Row>();
    /// <summary>Horizontal alignment of a table narrower than the body.</summary>
    public ParagraphAlignment Alignment { get; set; }
    /// <summary>Space before the table in millimeters.</summary>
    public double SpaceBeforeMillimeters { get; set; }
    /// <summary>Space after the table in millimeters.</summary>
    public double SpaceAfterMillimeters { get; set; }
    /// <summary>Starts a new page when preceding flow content exists.</summary>
    public bool PageBreakBefore { get; set; }
    /// <summary>Uniform grid stroke width in millimeters; zero disables borders. Strokes stay inside the table bounds.</summary>
    public double BorderWidthMillimeters { get; set; } = 0.2;
    /// <summary>Color of all grid edges; shared edges are drawn once.</summary>
    public OfdColor BorderColor { get; set; } = OfdColor.Black;
}

/// <summary>An indivisible table row; height is the larger of the minimum and measured cells.</summary>
public sealed class Row
{
    /// <summary>Cells in left-to-right order.</summary>
    public IList<Cell> Cells { get; } = new List<Cell>();
    /// <summary>Minimum row height in millimeters. No exact-height clipping is performed.</summary>
    public double MinimumHeightMillimeters { get; set; }
}

/// <summary>A cell with independently styled paragraphs. Paragraph alignment controls horizontal text alignment.</summary>
public sealed class Cell
{
    /// <summary>Creates an empty cell.</summary>
    public Cell() { }
    /// <summary>Creates a cell containing one paragraph.</summary>
    public Cell(string text) => Paragraphs.Add(new Paragraph(text));
    /// <summary>Paragraphs in reading order. Explicit page breaks are unsupported inside cells.</summary>
    public IList<Paragraph> Paragraphs { get; } = new List<Paragraph>();
    /// <summary>Number of consecutive columns covered by this cell.</summary>
    public int ColumnSpan { get; set; } = 1;
    /// <summary>Only 1 is supported. Vertical merging throws NotSupportedException.</summary>
    public int RowSpan { get; set; } = 1;
    /// <summary>Padding on each side in millimeters.</summary>
    public double PaddingMillimeters { get; set; } = 0.8;
    /// <summary>Optional cell fill, drawn before text and borders.</summary>
    public OfdColor? BackgroundColor { get; set; }
    /// <summary>Vertical placement of the entire paragraph group.</summary>
    public CellVerticalAlignment VerticalAlignment { get; set; }
}

/// <summary>Vertical alignment within the padded cell area.</summary>
public enum CellVerticalAlignment { Top, Center, Bottom }
