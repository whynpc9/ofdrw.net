using System;
using Ofdrw.Net.Core.Fonts;
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
    private int _totalPages = 1;
    private Dictionary<BuiltInSectionModel, int> _sectionPages = new();
    private readonly Dictionary<string, XFont> _fonts = new();
    private readonly Dictionary<string, (byte[] Data, string FileName, bool Bold, bool Italic)> _fontFiles = new();
    private readonly HashSet<string> _simulatedBoldFamilies = new(StringComparer.OrdinalIgnoreCase);
    private readonly NativeBoldStrokePlanner _boldStrokes;
    private int _pendingPageBreaks;
    private double _pendingSpaceAfter;
    private readonly List<double> _pendingLineBreakHeights = new();

    internal BuiltInOfdRenderer(
        DocxConversionOptions options,
        IList<DocxConversionDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        if (!Enum.IsDefined(typeof(DocxEmbeddedBoldBehavior), options.NativeEmbeddedBoldBehavior))
            throw new ArgumentOutOfRangeException(nameof(options), "Unknown native embedded bold behavior.");
        _options = options;
        _diagnostics = diagnostics;
        _cancellationToken = cancellationToken;
        _boldStrokes = new NativeBoldStrokePlanner(cancellationToken);
    }

    internal OfdDocumentPackage Render(BuiltInDocumentModel model)
    {
        _cancellationToken.ThrowIfCancellationRequested();

        RegisterConfiguredFonts();
        OfdDocumentPackage? result = null;
        FinalizePagination(() =>
        {
            _decorations.Clear();
            _lastPageNumber = 0;
            _pendingPageBreaks = 0;
            _pendingSpaceAfter = 0;
            _pendingLineBreakHeights.Clear();
            result = RenderPass(model);
            var counts = _decorations.GroupBy(value => value.Section)
                .ToDictionary(group => group.Key, group => group.Count());
            var stable = result.Pages.Count == _totalPages &&
                counts.All(pair => SectionPages(pair.Key) == pair.Value);
            _totalPages = result.Pages.Count;
            _sectionPages = counts;
            return stable;
        });
        return result!;
    }

    // A pass is committed only when every measured field context matches the actual
    // pagination. Rebuild all geometry after total digit widths change; never patch text.
    internal static void FinalizePagination(Func<bool> layoutPass)
    {
        const int maximumPasses = 8;
        for (var pass = 0; pass < maximumPasses; pass++)
            if (layoutPass()) return;
        throw new InvalidDataException("DOCX page field layout did not converge within 8 passes.");
    }

    private int SectionPages(BuiltInSectionModel section) =>
        _sectionPages.TryGetValue(section, out var count) ? count : 1;

    private OfdDocumentPackage RenderPass(BuiltInDocumentModel model)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        var package = new OfdDocumentPackage
        {
            Options = new OfdDocumentOptions
            {
                FontEmbedding = new OfdFontEmbeddingOptions
                {
                    Mode = _options.FontEmbedding.Mode,
                    MaximumFontBytes = Math.Min(_options.MaxEmbeddedFontBytes, _options.FontEmbedding.MaximumFontBytes),
                    MaximumCollectionBytes = _options.FontEmbedding.MaximumCollectionBytes,
                    MaximumUsedScalars = _options.FontEmbedding.MaximumUsedScalars,
                    MaximumCmapParseWork = _options.FontEmbedding.MaximumCmapParseWork,
                    MaximumGlyphClosureOperations = _options.FontEmbedding.MaximumGlyphClosureOperations
                },
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
        _pendingLineBreakHeights.Clear();
        var state = new LayoutState(section, section.PageNumberStart ?? _lastPageNumber + 1);
        EnsurePage(package, state);
        foreach (var block in section.Blocks)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            RenderBlock(package, state, block);
        }
        // A section boundary commits explicit breaks using the section that owns them.
        if (hasFollowingSection)
        {
            ApplyPendingPageBreaks(package, state);
            ApplyPendingLineBreaks(package, state);
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
        if (paragraph.Format.PageBreakBefore && (state.HasBodyContent || hadPendingBreaks || _pendingLineBreakHeights.Count > 0))
        {
            StartNewPage(package, state);
            _pendingSpaceAfter = 0;
            _pendingLineBreakHeights.Clear();
        }

        var spaceBefore = PointsToMillimeters(paragraph.Format.SpaceBeforePoints ?? 0);
        var spaceAfter = PointsToMillimeters(paragraph.Format.SpaceAfterPoints ?? 0);

        var leftIndent = PointsToMillimeters(paragraph.Format.LeftIndentPoints ?? 0);
        var rightIndent = PointsToMillimeters(paragraph.Format.RightIndentPoints ?? 0);
        var firstIndent = PointsToMillimeters(paragraph.Format.FirstLineIndentPoints ?? 0);
        var width = Math.Max(availableWidth - leftIndent - rightIndent, 5d);
        var firstLine = true;
        var preparedFirstLine = false;
        var leadingLines = 0;
        var lines = LayoutParagraph(paragraph, width, firstIndent, state.PageNumber, _totalPages, SectionPages(state.Section));
        if (paragraph.Inlines.OfType<BuiltInPageNumberModel>().Any(field => field.Kind == BuiltInPageFieldKind.Page))
        {
            // Body PAGE uses the paragraph's first body line page context. Resolve that context
            // after deferred pagination, and remeasure when its number gains a digit.
            while (leadingLines < lines.Count && lines[leadingLines].Glyphs.Count == 0)
            {
                var control = lines[leadingLines++];
                if (control.PageBreak) { _pendingPageBreaks++; continue; }
                if (control.TerminalNewline) { _pendingLineBreakHeights.Add(control.Height); continue; }
                ApplyPendingPageBreaks(package, state);
                if (firstLine)
                {
                    ApplyPendingLineBreaks(package, state);
                    PlaceFirstLine(package, state, control.Height, _pendingSpaceAfter + spaceBefore);
                    _pendingSpaceAfter = 0;
                    firstLine = false;
                }
                else EnsureVerticalSpace(package, state, control.Height);
                state.Y += control.Height;
                state.HasBodyContent = true;
            }
            ApplyPendingPageBreaks(package, state);
            ApplyPendingLineBreaks(package, state);
            var gap = firstLine ? Math.Max(0, _pendingSpaceAfter + spaceBefore) : 0;
            while (true)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                lines = LayoutParagraph(paragraph, width, firstIndent, state.PageNumber, _totalPages, SectionPages(state.Section));
                var first = lines.Skip(leadingLines).First(line => !line.PageBreak && !line.TerminalNewline);
                if (first.Height > state.ContentBottom - state.ContentTop)
                    throw new InvalidDataException("DOCX content exceeds the usable page area.");
                if (state.Y + gap + first.Height > state.ContentBottom + 0.000001d &&
                    (state.HasBodyContent || state.Y > state.ContentTop))
                {
                    StartNewPage(package, state);
                    continue;
                }
                PlaceFirstLine(package, state, first.Height, gap);
                _pendingSpaceAfter = 0;
                preparedFirstLine = true;
                firstLine = false;
                break;
            }
        }
        foreach (var line in lines.Skip(leadingLines))
        {
            if (line.PageBreak)
            {
                _pendingPageBreaks++;
                continue;
            }
            if (line.TerminalNewline) { _pendingLineBreakHeights.Add(line.Height); continue; }
            ApplyPendingPageBreaks(package, state);
            if (preparedFirstLine) preparedFirstLine = false;
            else if (firstLine)
            {
                ApplyPendingLineBreaks(package, state);
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

        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var row = table.Rows[rowIndex];
            if (firstRow) ApplyPendingLineBreaks(package, state);
            List<FlowMeasuredCell> cellLayouts;
            double rowHeight;
            // PAGE belongs to the row's actual placement page. A move can change
            // digit width and row height, so measure again before drawing any cell.
            while (true)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                var measuredPage = state.PageNumber;
                cellLayouts = MeasureRow(row, columnWidths, rowIndex + 1, state);
                rowHeight = cellLayouts.Count == 0
                    ? PointsToMillimeters(DefaultFontSizePoints) * 1.5d
                    : cellLayouts.Max(cell => cell.Height);
                if (firstRow) PlaceFirstLine(package, state, rowHeight, _pendingSpaceAfter);
                else EnsureVerticalSpace(package, state, rowHeight);
                if (measuredPage == state.PageNumber) break;
                // PlaceFirstLine already consumed the paragraph gap on the fresh page.
                // Undo it before remeasuring to avoid applying spacing twice.
                if (firstRow) state.Y = state.ContentTop;
            }
            if (firstRow) { _pendingSpaceAfter = 0; firstRow = false; }
            var rowTop = state.Y;
            var x = state.ContentLeft;
            var columnIndex = 0;
            var rowElementStart = state.Page!.Elements.Count;

            foreach (var cellLayout in cellLayouts)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                var cellWidth = cellLayout.Width;

                var cell = row.Cells[cellLayouts.IndexOf(cellLayout)];
                DrawCell(state.Page!, table, cell, rowIndex, columnIndex, columnWidths.Count,
                    x, rowTop, cellWidth, rowHeight);
                var textLeft = x + MinCellPaddingMillimeters;

                var localY = rowTop + FlowTableLayout.VerticalOffset(rowHeight, cellLayout, (int)cell.VerticalAlignment);
                foreach (var line in cellLayout.Content.Lines)
                    DrawLine(package, state.Page!, line.Line, textLeft + line.Left, localY + line.Top, line.Width);

                x += cellWidth;
                columnIndex += cellLayout.ColumnSpan;
            }

            // A cell's later paragraph is independent content, even when its
            // formatting matches. Keep the existing per-object policy for any
            // multi-paragraph row rather than flattening those identities.
            if (row.Cells.All(cell => cell.Paragraphs.Count == 1))
                HarmonizeRowBoldStroke(state.Page!, rowElementStart);
            state.Y = rowTop + rowHeight;
        }
    }

    private void HarmonizeRowBoldStroke(OfdPage page, int firstElement)
    {
        // These are freshly emitted native objects in this one measured row,
        // with exactly one paragraph per cell. Automatic wrapping keeps that
        // paragraph's lines together; a separate paragraph never participates.
        // Equal formatting should not become visibly different weights merely
        // because a sibling cell contains a narrow counter. Lower other cells
        // to the already checked minimum; never thicken the constrained glyph.
        // Pure Latin keeps its separate legacy emphasis contract.
        var candidates = page.Elements.Skip(firstElement).OfType<OfdTextElement>()
            .Where(text =>
            {
                _cancellationToken.ThrowIfCancellationRequested();
                return text.EmitSyntheticBoldStroke && !text.Italic &&
                    string.IsNullOrEmpty(text.SourceXml) &&
                    text.SyntheticBoldStrokeWidthMillimeters is > 0 &&
                    FontEmphasisMetrics.BoldStrokeWidth(text.FontSizeMillimeters, text.Text) <
                    FontEmphasisMetrics.BoldStrokeWidth(text.FontSizeMillimeters);
            })
            .GroupBy(text => (text.FontResourceId,
                Size: text.FontSizeMillimeters.ToString("0.###", CultureInfo.InvariantCulture),
                text.FillColor.Red, text.FillColor.Green, text.FillColor.Blue, text.FillColor.Alpha));
        foreach (var group in candidates)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var width = group.Min(text => text.SyntheticBoldStrokeWidthMillimeters!.Value);
            foreach (var text in group) text.SyntheticBoldStrokeWidthMillimeters = width;
        }
    }

    private List<FlowMeasuredCell> MeasureRow(BuiltInTableRowModel row, IReadOnlyList<double> columnWidths, int rowNumber, LayoutState state)
    {
        var specs = row.Cells.Select(cell => new FlowCellSpec
        {
            ColumnSpan = cell.ColumnSpan,
            Padding = MinCellPaddingMillimeters,
            Measure = width =>
            {
                var content = new FlowCellContent();
                // Preserve the existing Native cell paragraph policy while sharing row geometry.
                foreach (var paragraph in cell.Paragraphs)
                    foreach (var line in LayoutParagraph(paragraph, width, 0, state.PageNumber, _totalPages, SectionPages(state.Section)).Where(line => !line.PageBreak))
                    {
                        content.Lines.Add(new FlowCellLine(line, content.Height, 0, width));
                        content.Height += line.Height;
                    }
                return content;
            }
        });
        try { return FlowTableLayout.MeasureRow(columnWidths, specs, _cancellationToken); }
        catch (ArgumentException exception)
        {
            throw new InvalidDataException($"DOCX table row {rowNumber} has an invalid grid or unusable cell width: {exception.Message}", exception);
        }
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

    // Name-only resources retain requested style for viewer substitution.
    // Embedded resources override these flags with the physical face style.
    private (bool Bold, bool Italic) DeclaredResourceStyle(BuiltInTextFormat format) =>
        (format.Bold, format.Italic);

    private OfdFontResource GetOrAddFontResource(OfdDocumentPackage package, BuiltInTextFormat format, string declaredName)
    {
        var (bold, italic) = DeclaredResourceStyle(format);
        var embedded = _configuredFonts.ShouldEmbed(declaredName);
        var existing = package.Fonts.FirstOrDefault(font =>
            font.FontName == declaredName && (embedded || font.Bold == bold && font.Italic == italic));
        if (existing is not null)
            return existing;

        var resource = new OfdFontResource
        {
            Id = "F" + (package.Fonts.Count + 1),
            FontName = declaredName,
            FamilyName = NormalizeFontFamily(format.FontFamily),
            Charset = "unicode",
            Bold = bold,
            Italic = italic
        };
        if (embedded)
        {
            var resolveFamily = DocxFontCatalog.IsViewerLocalCjkFamily(declaredName)
                ? declaredName
                : NormalizeFontFamily(format.FontFamily);
            // Use the same physical style selection as GetFont (layout).
            var resolveBold = format.Bold;
            var resolver = GlobalFontSettings.FontResolver;
            var face = resolver.ResolveTypeface(_configuredFonts.Resolve(resolveFamily, resolveBold, format.Italic), resolveBold, format.Italic);
            if (!_fontFiles.TryGetValue(face.FaceName, out var file))
            {
                var data = PdfFontRegistry.GetOriginalFont(face.FaceName);
                using var hash = SHA256.Create();
                var digest = BitConverter.ToString(hash.ComputeHash(data)).Replace("-", string.Empty).ToLowerInvariant();
                var physicalStyle = new OpenTypeFace(data).Style;
                file = (data, "native-font-" + digest + ".ttf", physicalStyle.Bold, physicalStyle.Italic);
                _fontFiles[face.FaceName] = file;
            }
            resource.FileName = file.FileName;
            resource.Data = file.Data;
            // Font flags describe the actual embedded face, not requested text
            // emphasis. The latter is carried by Weight/Italic on TextObject.
            resource.Bold = file.Bold;
            resource.Italic = file.Italic;
            if (format.Bold && !file.Bold)
            {
                if (_options.NativeEmbeddedBoldBehavior == DocxEmbeddedBoldBehavior.Throw)
                    throw new InvalidDataException($"Native bold text requires an actual bold face for '{resource.FamilyName}'. Supply it through FontDirectories or explicitly allow TextStroke simulation.");
                if (_simulatedBoldFamilies.Add(resource.FamilyName))
                    _diagnostics.Add(new DocxConversionDiagnostic("DOCX_EMBEDDED_BOLD_SIMULATED",
                        $"'{resource.FamilyName}' has no selected bold face; native OFD uses a text stroke for emphasis while preserving the original font bytes and text."));
            }
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
        // Name-only CJK output retains compatibility metrics. In particular,
        // missing Noto must not measure a host resolver's unrelated stand-in.
        // Configured embedded faces measure Latin with their own bytes.
        var nameOnlyCjk = DocxFontCatalog.IsViewerLocalCjkFamily(family) ||
            DocxFontCatalog.IsNotoCjkFamily(family) && !_configuredFonts.ShouldEmbed(family);
        if (!nameOnlyCjk)
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
        // Missing whitespace glyphs must not inherit a .notdef advance.
        if (text.Length > 0 && text[0] == '\u00A0') return Advance(" ", format);
        if (text.Length > 0 && text[0] == '\u202F') return PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints) * 0.2d;
        if (text.Length > 0 && text[0] == '\u2007') return Advance("0", format);
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
                advance = measure.MeasureString(text == "\t" ? "    " : FlowTextMetrics.NormalizeLatinForMeasurement(text),
                    GetLatinCompatibleFont(format)).Width / 1000d;
            }
            _advances[key] = advance;
        }
        return advance * PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints);
    }

    public double AdvanceMillimeters(string grapheme, FlowTextStyle style) =>
        Advance(grapheme, (BuiltInTextFormat)style.Source!);

    private void ApplyPendingLineBreaks(OfdDocumentPackage package, LayoutState state)
    {
        foreach (var height in _pendingLineBreakHeights)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            EnsureVerticalSpace(package, state, height);
            state.Y += height;
            state.HasBodyContent = true;
        }
        _pendingLineBreakHeights.Clear();
    }

    private void ApplyPendingPageBreaks(OfdDocumentPackage package, LayoutState state)
    {
        if (_pendingPageBreaks > 0) { _pendingSpaceAfter = 0; _pendingLineBreakHeights.Clear(); }
        while (_pendingPageBreaks > 0)
        {
            StartNewPage(package, state);
            _pendingPageBreaks--;
        }
    }

    private IReadOnlyList<FlowLine> LayoutParagraph(BuiltInParagraphModel paragraph, double width, double firstIndent, int pageNumber, int totalPages, int sectionPages)
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
                case BuiltInBreakModel br: Add(br.IsPageBreak ? "\f" : "\n", br.Format ?? fallbackFormat); break;
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
        var x = FlowParagraphLayout.Align(left + line.Indent, width - line.Indent, line.AlignmentWidth, line.Alignment);
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
            var groupWidth = group.Sum(g => g.Width);
            var elementWidth = Math.Max(groupWidth, 0.1);
            var declaredName = DeclaredFontName(format);
            var resource = GetOrAddFontResource(package, format, declaredName);
            var text = new OfdTextElement
            {
                LayerType = "Body", XMillimeters = groupWidth == 0 ? Math.Min(x, left + width - elementWidth) : x,
                YMillimeters = top, WidthMillimeters = elementWidth, HeightMillimeters = line.Height,
                FontName = declaredName, FontResourceId = resource.Id,
                // Keep requested text style separate from the physical font face.
                EmitSyntheticBoldStroke = format.Bold && resource.Data.Length > 0 && !resource.Bold,
                UseNameOnlyBoldSelection = resource.Data.Length == 0 &&
                    DocxFontCatalog.IsNotoCjkFamily(NormalizeFontFamily(format.FontFamily)),
                Weight = format.Bold ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight,
                Italic = format.Italic,
                FontSizeMillimeters = PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints),
                FillColor = ParseColor(format.ColorHex), Text = string.Concat(group.Select(g => g.Text))
            };
            if (text.EmitSyntheticBoldStroke && !text.Italic)
                text.SyntheticBoldStrokeWidthMillimeters = _boldStrokes.Width(resource.Data, resource.Italic,
                    text.FontSizeMillimeters, text.Text);
            text.Runs.Add(new OfdTextRun { Text = text.Text, YMillimeters = baseline,
                DeltaX = group.Count > 1 ? string.Join(" ", group.Take(group.Count - 1).Select(g => g.Width.ToString("0.######", CultureInfo.InvariantCulture))) : null });
            page.Elements.Add(text);
            x += groupWidth;
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
        var header = LayoutDecoration(headers, state.ContentWidth, state.PageNumber, _totalPages, SectionPages(state.Section));
        var footer = LayoutDecoration(footers, state.ContentWidth, state.PageNumber, _totalPages, SectionPages(state.Section));
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

}
