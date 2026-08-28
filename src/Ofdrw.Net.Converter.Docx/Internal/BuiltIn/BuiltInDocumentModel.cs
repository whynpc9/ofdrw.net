using System.Collections.Generic;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

internal sealed class BuiltInDocumentModel
{
    internal IList<BuiltInSectionModel> Sections { get; } = new List<BuiltInSectionModel>();
}

internal sealed class BuiltInSectionModel
{
    internal double PageWidthPoints { get; set; } = 595.3;
    internal double PageHeightPoints { get; set; } = 841.9;
    internal double MarginTopPoints { get; set; } = 72;
    internal double MarginRightPoints { get; set; } = 72;
    internal double MarginBottomPoints { get; set; } = 72;
    internal double MarginLeftPoints { get; set; } = 72;
    internal IList<BuiltInBlockModel> Blocks { get; } = new List<BuiltInBlockModel>();
    internal IList<BuiltInParagraphModel> Headers { get; } = new List<BuiltInParagraphModel>();
    internal IList<BuiltInParagraphModel> Footers { get; } = new List<BuiltInParagraphModel>();
}

internal abstract class BuiltInBlockModel
{
}

internal sealed class BuiltInParagraphModel : BuiltInBlockModel
{
    internal IList<BuiltInInlineModel> Inlines { get; } = new List<BuiltInInlineModel>();
    internal BuiltInParagraphFormat Format { get; } = new();
}

internal sealed class BuiltInTableModel : BuiltInBlockModel
{
    internal IList<double> ColumnWidthsPoints { get; } = new List<double>();
    internal IList<BuiltInTableRowModel> Rows { get; } = new List<BuiltInTableRowModel>();
    internal bool HasBorders { get; set; }
}

internal sealed class BuiltInTableRowModel
{
    internal IList<BuiltInTableCellModel> Cells { get; } = new List<BuiltInTableCellModel>();
    internal bool IsHeader { get; set; }
}

internal sealed class BuiltInTableCellModel
{
    internal IList<BuiltInParagraphModel> Paragraphs { get; } = new List<BuiltInParagraphModel>();
    internal int ColumnSpan { get; set; } = 1;
    internal string? ShadingHex { get; set; }
}

internal abstract class BuiltInInlineModel
{
}

internal sealed class BuiltInTextModel : BuiltInInlineModel
{
    internal string Text { get; set; } = string.Empty;
    internal BuiltInTextFormat Format { get; } = new();
}

internal sealed class BuiltInBreakModel : BuiltInInlineModel
{
    internal bool IsPageBreak { get; set; }
}

internal sealed class BuiltInTabModel : BuiltInInlineModel
{
}

internal sealed class BuiltInImageModel : BuiltInInlineModel
{
    internal byte[] Data { get; set; } = [];
    internal string Name { get; set; } = string.Empty;
    internal double? WidthPoints { get; set; }
    internal double? HeightPoints { get; set; }
}

internal sealed class BuiltInPageNumberModel : BuiltInInlineModel
{
}

internal sealed class BuiltInParagraphFormat
{
    internal BuiltInParagraphAlignment Alignment { get; set; }
    internal double? SpaceBeforePoints { get; set; }
    internal double? SpaceAfterPoints { get; set; }
    internal double? LineSpacingPoints { get; set; }
    internal double? LeftIndentPoints { get; set; }
    internal double? RightIndentPoints { get; set; }
    internal double? FirstLineIndentPoints { get; set; }
    internal bool PageBreakBefore { get; set; }
    internal bool KeepWithNext { get; set; }
    internal string? ListMarker { get; set; }
}

internal sealed class BuiltInTextFormat
{
    internal string? FontFamily { get; set; }
    internal double? FontSizePoints { get; set; }
    internal bool Bold { get; set; }
    internal bool Italic { get; set; }
    internal bool Underline { get; set; }
    internal string? ColorHex { get; set; }
}

internal enum BuiltInParagraphAlignment
{
    Left,
    Center,
    Right,
    Justify
}
