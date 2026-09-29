using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Layout.Internal.Flow;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

/// <summary>
/// Lays out a <see cref="BuiltInDocumentModel"/> directly as structured OFD page
/// objects (text + table borders) without PDF rasterization.
/// </summary>
internal sealed class BuiltInOfdRenderer : IFlowFontMetrics
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
    private int _pendingPageBreaks;
    private double _pendingSpaceAfter;

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
                Namespace = _options.OfdNamespace,
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
        for (var sectionIndex = 0; sectionIndex < sections.Count; sectionIndex++)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            lastState = RenderSection(package, sections[sectionIndex], sectionIndex + 1 < sections.Count);
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

    private LayoutState RenderSection(OfdDocumentPackage package, BuiltInSectionModel section, bool hasFollowingSection)
    {
        _pendingSpaceAfter = 0;
        var state = new LayoutState(section, section.PageNumberStart ?? _lastPageNumber + 1);
        EnsurePage(package, state);
        foreach (var block in section.Blocks)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            RenderBlock(package, state, block);
        }
        // A section boundary commits explicit breaks using the section that owns them.
        if (hasFollowingSection) ApplyPendingPageBreaks(package, state);
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
                ApplyPendingPageBreaks(package, state);
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
        var hadPendingBreaks = _pendingPageBreaks > 0;
        ApplyPendingPageBreaks(package, state);
        if (paragraph.Format.PageBreakBefore && (state.HasBodyContent || hadPendingBreaks))
        {
            StartNewPage(package, state);
            _pendingSpaceAfter = 0;
        }

        var spaceBefore = PointsToMillimeters(paragraph.Format.SpaceBeforePoints ?? 0);
        var spaceAfter = PointsToMillimeters(paragraph.Format.SpaceAfterPoints ?? 0);

        var leftIndent = PointsToMillimeters(paragraph.Format.LeftIndentPoints ?? 0);
        var rightIndent = PointsToMillimeters(paragraph.Format.RightIndentPoints ?? 0);
        var firstIndent = PointsToMillimeters(paragraph.Format.FirstLineIndentPoints ?? 0);
        var width = Math.Max(availableWidth - leftIndent - rightIndent, 5d);
        var firstLine = true;
        foreach (var line in LayoutParagraph(paragraph, width, firstIndent, state.PageNumber))
        {
            if (line.PageBreak)
            {
                _pendingPageBreaks++;
                continue;
            }
            ApplyPendingPageBreaks(package, state);
            if (firstLine)
            {
                PlaceFirstLine(package, state, line.Height, _pendingSpaceAfter + spaceBefore);
                _pendingSpaceAfter = 0;
                firstLine = false;
            }
            else EnsureVerticalSpace(package, state, line.Height);
            DrawLine(package, state.Page!, line, left + leftIndent, state.Y, width);
            state.Y += line.Height;
            state.HasBodyContent = true;
        }

        _pendingSpaceAfter += spaceAfter;
    }

    private void RenderTable(
        OfdDocumentPackage package,
        LayoutState state,
        BuiltInTableModel table)
    {
        var columnWidths = ResolveColumnWidths(table, state.ContentWidth);
        var firstRow = true;

        foreach (var row in table.Rows)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var cellLayouts = MeasureRow(row, columnWidths);
            var rowHeight = cellLayouts.Count == 0
                ? PointsToMillimeters(DefaultFontSizePoints) * 1.5d
                : cellLayouts.Max(cell => cell.Height);

            if (firstRow)
            {
                PlaceFirstLine(package, state, rowHeight, _pendingSpaceAfter);
                _pendingSpaceAfter = 0;
                firstRow = false;
            }
            else EnsureVerticalSpace(package, state, rowHeight);
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
            var lines = new List<FlowLine>();
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

    private string DeclaredFontName(BuiltInTextFormat format)
    {
        var family = NormalizeFontFamily(format.FontFamily);
        if (DocxFontCatalog.IsViewerLocalCjkFamily(family))
            return family;
        return FontKey(format);
    }

    // Preserve the source family and requested style even when a viewer substitutes
    // a regular CJK face. The PDF resolver inspects the actual bytes and simulates
    // only styles that the selected font does not supply.
    private (bool Bold, bool Italic) DeclaredResourceStyle(BuiltInTextFormat format) =>
        (format.Bold, format.Italic);

    private OfdFontResource GetOrAddFontResource(OfdDocumentPackage package, BuiltInTextFormat format, string declaredName)
    {
        var (bold, italic) = DeclaredResourceStyle(format);
        var existing = package.Fonts.FirstOrDefault(font =>
            font.FontName == declaredName && font.Bold == bold && font.Italic == italic);
        if (existing is not null)
            return existing;

        var resource = new OfdFontResource
        {
            Id = "F" + (package.Fonts.Count + 1),
            FontName = declaredName,
            FamilyName = declaredName,
            Charset = "unicode",
            Bold = bold,
            Italic = italic
        };
        if (_configuredFonts.ShouldEmbed(declaredName))
        {
            var resolveFamily = DocxFontCatalog.IsViewerLocalCjkFamily(declaredName)
                ? declaredName
                : NormalizeFontFamily(format.FontFamily);
            var resolveBold = format.Bold &&
                declaredName.Equals(NormalizeFontFamily(format.FontFamily), StringComparison.OrdinalIgnoreCase);
            var resolver = GlobalFontSettings.FontResolver;
            var face = resolver.ResolveTypeface(_configuredFonts.Resolve(resolveFamily, resolveBold, format.Italic), resolveBold, format.Italic);
            if (!_fontFiles.TryGetValue(face.FaceName, out var file))
            {
                var data = PdfFontRegistry.GetOriginalFont(face.FaceName);
                using var hash = SHA256.Create();
                var digest = BitConverter.ToString(hash.ComputeHash(data)).Replace("-", string.Empty).ToLowerInvariant();
                file = (data, "native-font-" + digest + ".ttf");
                _fontFiles[face.FaceName] = file;
            }
            resource.FileName = file.FileName;
            resource.Data = file.Data;
        }
        package.Fonts.Add(resource);
        return resource;
    }

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

    private XFont GetLatinCompatibleFont(BuiltInTextFormat format)
    {
        var family = NormalizeFontFamily(format.FontFamily);
        if (!DocxFontCatalog.IsViewerLocalCjkFamily(family) &&
            !family.StartsWith("Noto Sans CJK", StringComparison.OrdinalIgnoreCase))
        {
            return GetFont(format);
        }

        var key = "latin|" + (format.Bold ? "bold" : "regular") + (format.Italic ? "|italic" : "");
        if (_fonts.TryGetValue(key, out var font)) return font;
        PdfFontRegistry.EnsureInstalled();
        var style = (format.Bold ? XFontStyle.Bold : XFontStyle.Regular) |
                    (format.Italic ? XFontStyle.Italic : XFontStyle.Regular);
        font = new XFont("Arial", 1000, style);
        _fonts[key] = font;
        return font;
    }

    private double Advance(string text, BuiltInTextFormat format)
    {
        var key = FontKey(format) + "\n" + text;
        if (!_advances.TryGetValue(key, out var advance))
        {
            // linux1 / 宋体 CJK is one em. PDFsharp without a CJK face (typical Linux
            // ARM without FontDirectories) measures ideographs on a Latin substitute
            // at ~0.6em, which overlaps when the viewer later binds SimSun/SimHei.
            if (FlowTextMetrics.IsCjkTypographicUnit(text))
                advance = 1d;
            else
            {
                using var measure = XGraphics.CreateMeasureContext(new XSize(1000, 1000), XGraphicsUnit.Point, XPageDirection.Downwards);
                advance = measure.MeasureString(text == "\t" ? "    " : text, GetLatinCompatibleFont(format)).Width / 1000d;
            }
            _advances[key] = advance;
        }
        return advance * PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints);
    }

    public double AdvanceMillimeters(string grapheme, FlowTextStyle style) =>
        Advance(grapheme, (BuiltInTextFormat)style.Source!);

    private void ApplyPendingPageBreaks(OfdDocumentPackage package, LayoutState state)
    {
        if (_pendingPageBreaks > 0) _pendingSpaceAfter = 0;
        while (_pendingPageBreaks > 0)
        {
            StartNewPage(package, state);
            _pendingPageBreaks--;
        }
    }

    private IReadOnlyList<FlowLine> LayoutParagraph(BuiltInParagraphModel paragraph, double width, double firstIndent, int pageNumber = 1, int totalPages = 1, int sectionPages = 1)
    {
        var inlines = new List<FlowInline>();
        FlowTextStyle Style(BuiltInTextFormat format) => new()
        {
            FontFamily = NormalizeFontFamily(format.FontFamily),
            FontSizeMillimeters = PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints),
            Bold = format.Bold, Italic = format.Italic, Source = format
        };
        var fallbackFormat = paragraph.Inlines.OfType<BuiltInTextModel>().FirstOrDefault()?.Format ?? new BuiltInTextFormat();
        var fallback = Style(fallbackFormat);
        void Add(string value, BuiltInTextFormat format) =>
            inlines.Add(new FlowInline { Text = value, Style = Style(format) });
        foreach (var inline in paragraph.Inlines)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            switch (inline)
            {
                case BuiltInTextModel text: Add(text.Text, text.Format); break;
                case BuiltInTabModel: Add("\t", fallbackFormat); break;
                case BuiltInBreakModel br: Add(br.IsPageBreak ? "\f" : "\n", fallbackFormat); break;
                case BuiltInPageNumberModel field:
                    var value = field.Kind == BuiltInPageFieldKind.TotalPages ? totalPages :
                        field.Kind == BuiltInPageFieldKind.SectionPages ? sectionPages : pageNumber;
                    Add(value.ToString(CultureInfo.InvariantCulture), field.Format);
                    break;
                case BuiltInImageModel image:
                    var imageWidth = PointsToMillimeters(image.WidthPoints ?? 72);
                    var imageHeight = PointsToMillimeters(image.HeightPoints ?? 72);
                    inlines.Add(new FlowInline { Style = fallback, Image = image,
                        ImageWidthMillimeters = imageWidth, ImageHeightMillimeters = imageHeight });
                    break;
            }
        }
        var alignment = paragraph.Format.Alignment switch
        {
            BuiltInParagraphAlignment.Center => FlowAlignment.Center,
            BuiltInParagraphAlignment.Right => FlowAlignment.Right,
            _ => FlowAlignment.Left
        };
        return FlowParagraphLayout.Layout(inlines, new FlowParagraphFormat
        {
            Alignment = alignment,
            FirstLineIndentMillimeters = firstIndent,
            MinimumLineHeightMillimeters = PointsToMillimeters(paragraph.Format.LineSpacingPoints ?? 0)
        }, width, this, fallback, _cancellationToken, allowOversizeGlyph: true);
    }

    private void DrawLine(OfdDocumentPackage package, OfdPage page, FlowLine line, double left, double top, double width)
    {
        var x = FlowParagraphLayout.Align(left + line.Indent, width - line.Indent, line.Glyphs.Sum(g => g.Width), line.Alignment);
        var baseline = line.BaselineMillimeters;
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
            var format = (BuiltInTextFormat)line.Glyphs[i].Style.Source!;
            while (i < line.Glyphs.Count && line.Glyphs[i].Image is null && ReferenceEquals(line.Glyphs[i].Style.Source, format)) i++;
            var group = line.Glyphs.GetRange(start, i - start);
            var declaredName = DeclaredFontName(format);
            var resource = GetOrAddFontResource(package, format, declaredName);
            var text = new OfdTextElement
            {
                LayerType = "Body", XMillimeters = x, YMillimeters = top,
                WidthMillimeters = Math.Max(group.Sum(g => g.Width), 0.1), HeightMillimeters = line.Height,
                FontName = declaredName, FontResourceId = resource.Id,
                // Viewers apply bold/italic from the text object, not from the font
                // resource flags, so declare the requested style on both.
                Weight = format.Bold ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight,
                Italic = format.Italic,
                FontSizeMillimeters = PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints),
                FillColor = ParseColor(format.ColorHex), Text = string.Concat(group.Select(g => g.Text))
            };
            text.Runs.Add(new OfdTextRun { Text = text.Text, YMillimeters = baseline,
                DeltaX = group.Count > 1 ? string.Join(" ", group.Take(group.Count - 1).Select(g => g.Width.ToString("0.######", CultureInfo.InvariantCulture))) : null });
            page.Elements.Add(text);
            x += group.Sum(g => g.Width);
        }
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
        if (state.Page is null) StartNewPage(package, state);
        try
        {
            if (FlowPagination.NeedsNewPage(state.Y, needed, state.ContentTop, state.ContentBottom))
                StartNewPage(package, state);
            // Header and footer variants can change the usable area on the next page.
            if (FlowPagination.NeedsNewPage(state.Y, needed, state.ContentTop, state.ContentBottom))
                throw new InvalidDataException("DOCX content cannot fit on a fresh page.");
        }
        catch (InvalidOperationException exception)
        {
            throw new InvalidDataException("DOCX content exceeds the usable page area.", exception);
        }
    }

    private void PlaceFirstLine(OfdDocumentPackage package, LayoutState state, double height, double gap)
    {
        if (state.Page is null) StartNewPage(package, state);
        if (height > state.ContentBottom - state.ContentTop)
            throw new InvalidDataException("DOCX content exceeds the usable page area.");
        gap = Math.Max(0, gap);
        if (state.Y + gap + height > state.ContentBottom + 0.000001d &&
            (state.HasBodyContent || state.Y > state.ContentTop))
            StartNewPage(package, state);
        var availableGap = state.ContentBottom - state.ContentTop - height;
        if (availableGap < 0) throw new InvalidDataException("DOCX content cannot fit on a fresh page.");
        state.Y += Math.Min(gap, availableGap);
        EnsureVerticalSpace(package, state, height);
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
        internal List<(FlowLine Line, double Left, double Top, double Width)> Lines { get; } = new();
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
            List<FlowLine> lines,
            BuiltInVerticalAlignment verticalAlignment)
        {
            ColumnSpan = columnSpan;
            Height = height;
            Lines = lines;
            VerticalAlignment = verticalAlignment;
        }

        internal int ColumnSpan { get; }

        internal double Height { get; }

        internal List<FlowLine> Lines { get; }

        internal BuiltInVerticalAlignment VerticalAlignment { get; }
    }

}
