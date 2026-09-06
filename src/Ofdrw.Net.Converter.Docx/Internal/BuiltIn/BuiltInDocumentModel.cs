using System.Collections.Generic;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

internal sealed class BuiltInDocumentModel
{
    internal IList<BuiltInSectionModel> Sections { get; } = new List<BuiltInSectionModel>();
    internal IList<BuiltInSupplementalText> SupplementalText { get; } = new List<BuiltInSupplementalText>();
}

internal sealed class BuiltInSupplementalText
{
    internal string Kind { get; set; } = string.Empty;
    internal string Id { get; set; } = string.Empty;
    internal IList<BuiltInBlockModel> Blocks { get; } = new List<BuiltInBlockModel>();

    internal BuiltInParagraphModel CreateHeading()
    {
        var heading = new BuiltInParagraphModel();
        heading.Format.SpaceBeforePoints = 6;
        heading.Format.SpaceAfterPoints = 3;
        heading.Inlines.Add(new BuiltInTextModel { Text = $"[{Kind} {Id}]", Format = { Bold = true, FontSizePoints = 9 } });
        return heading;
    }
}

internal sealed class BuiltInSectionModel
{
    internal double PageWidthPoints { get; set; } = 595.3;
    internal double PageHeightPoints { get; set; } = 841.9;
    internal double MarginTopPoints { get; set; } = 72;
    internal double MarginRightPoints { get; set; } = 72;
    internal double MarginBottomPoints { get; set; } = 72;
    internal double MarginLeftPoints { get; set; } = 72;
    internal double HeaderDistancePoints { get; set; } = 36;
    internal double FooterDistancePoints { get; set; } = 36;
    internal bool DifferentFirstPage { get; set; }
    internal bool DifferentOddAndEvenPages { get; set; }
    internal int? PageNumberStart { get; set; }
    internal IList<BuiltInBlockModel> Blocks { get; } = new List<BuiltInBlockModel>();
    internal IList<BuiltInParagraphModel> Headers { get; } = new List<BuiltInParagraphModel>();
    internal IList<BuiltInParagraphModel> Footers { get; } = new List<BuiltInParagraphModel>();
    internal IList<BuiltInParagraphModel> FirstHeaders { get; } = new List<BuiltInParagraphModel>();
    internal IList<BuiltInParagraphModel> FirstFooters { get; } = new List<BuiltInParagraphModel>();
    internal IList<BuiltInParagraphModel> EvenHeaders { get; } = new List<BuiltInParagraphModel>();
    internal IList<BuiltInParagraphModel> EvenFooters { get; } = new List<BuiltInParagraphModel>();

    internal IList<BuiltInParagraphModel> GetHeaders(int sectionPageIndex, int pageNumber) =>
        DifferentFirstPage && sectionPageIndex == 0 ? FirstHeaders :
        DifferentOddAndEvenPages && pageNumber % 2 == 0 ? EvenHeaders : Headers;

    internal IList<BuiltInParagraphModel> GetFooters(int sectionPageIndex, int pageNumber) =>
        DifferentFirstPage && sectionPageIndex == 0 ? FirstFooters :
        DifferentOddAndEvenPages && pageNumber % 2 == 0 ? EvenFooters : Footers;
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
    internal Dictionary<string, BuiltInBorderModel> Borders { get; } = new();
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
    internal Dictionary<string, BuiltInBorderModel> Borders { get; } = new();
    internal string? ShadingHex { get; set; }
    internal BuiltInVerticalAlignment VerticalAlignment { get; set; } = BuiltInVerticalAlignment.Top;
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
    internal string MediaType { get; set; } = "image/png";
    internal string Name { get; set; } = string.Empty;
    internal double? WidthPoints { get; set; }
    internal double? HeightPoints { get; set; }
}

internal sealed class BuiltInPageNumberModel : BuiltInInlineModel
{
    internal BuiltInPageFieldKind Kind { get; set; }
    internal BuiltInTextFormat Format { get; } = new();
}

internal enum BuiltInPageFieldKind { Page, TotalPages, SectionPages }

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

internal enum BuiltInVerticalAlignment
{
    Top,
    Center,
    Bottom
}

internal sealed class BuiltInBorderModel
{
    internal bool Visible { get; set; }
    internal string? ColorHex { get; set; }
    internal double WidthPoints { get; set; } = 0.5;
}
