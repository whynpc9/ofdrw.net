using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.Tables;
using MigraDocCore.Rendering;
using Ofdrw.Net.Converter.Pdf;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;
using PdfSharpCore.Utils;
using SixLabors.ImageSharp.PixelFormats;
using MigraImageSource = MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes.ImageSource;
using MigraParagraph = MigraDocCore.DocumentObjectModel.Paragraph;
using MigraSection = MigraDocCore.DocumentObjectModel.Section;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

internal sealed class BuiltInPdfRenderer
{
    private readonly DocxConversionOptions _options;
    private readonly IList<DocxConversionDiagnostic> _diagnostics;
    private readonly CancellationToken _cancellationToken;
    private readonly HashSet<string> _reportedDefaultFonts = new(StringComparer.OrdinalIgnoreCase);
    private int _imageIndex;

    internal BuiltInPdfRenderer(
        DocxConversionOptions options,
        IList<DocxConversionDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        _options = options;
        _diagnostics = diagnostics;
        _cancellationToken = cancellationToken;
    }

    internal void Render(BuiltInDocumentModel source, string outputPath)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        PdfFontRegistry.EnsureInstalled();
        RegisterConfiguredFonts();
        if (MigraImageSource.ImageSourceImpl is null)
        {
            MigraImageSource.ImageSourceImpl = new ImageSharpImageSource<Rgba32>();
        }

        var document = new Document();

        foreach (var sourceSection in source.Sections)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var section = document.AddSection();
            ApplyPageSetup(section, sourceSection);
            RenderHeaderFooter(section.Headers.Primary, sourceSection.Headers);
            RenderHeaderFooter(section.Footers.Primary, sourceSection.Footers);

            foreach (var block in sourceSection.Blocks)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                switch (block)
                {
                    case BuiltInParagraphModel paragraph:
                        RenderSectionParagraph(section, paragraph);
                        break;
                    case BuiltInTableModel table:
                        RenderTable(section, table);
                        break;
                }
            }
        }

        if (source.Sections.Count == 0)
        {
            document.AddSection().AddParagraph(string.Empty);
        }

        var renderer = new PdfDocumentRenderer(true)
        {
            Document = document,
            WorkingDirectory = Path.GetDirectoryName(outputPath)
        };
        _cancellationToken.ThrowIfCancellationRequested();
        renderer.RenderDocument();
        _cancellationToken.ThrowIfCancellationRequested();
        renderer.PdfDocument.Save(outputPath);
    }

    private void RegisterConfiguredFonts()
    {
        foreach (var directory in _options.FontDirectories)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                _diagnostics.Add(new DocxConversionDiagnostic(
                    "DOCX_FONT_DIRECTORY_MISSING",
                    "A configured BuiltIn font directory does not exist."));
                continue;
            }

            foreach (var path in Directory.EnumerateFiles(directory))
            {
                _cancellationToken.ThrowIfCancellationRequested();
                var extension = Path.GetExtension(path);
                if (!extension.Equals(".ttf", StringComparison.OrdinalIgnoreCase) &&
                    !extension.Equals(".otf", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    PdfFontRegistry.RegisterFont(
                        Path.GetFileNameWithoutExtension(path),
                        File.ReadAllBytes(path));
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    _diagnostics.Add(new DocxConversionDiagnostic(
                        "DOCX_FONT_LOAD_FAILED",
                        "A configured font file could not be loaded."));
                }
            }
        }
    }

    private static void ApplyPageSetup(MigraSection target, BuiltInSectionModel source)
    {
        target.PageSetup.PageWidth = Unit.FromPoint(source.PageWidthPoints);
        target.PageSetup.PageHeight = Unit.FromPoint(source.PageHeightPoints);
        target.PageSetup.Orientation = source.PageWidthPoints > source.PageHeightPoints
            ? Orientation.Landscape
            : Orientation.Portrait;
        target.PageSetup.TopMargin = Unit.FromPoint(source.MarginTopPoints);
        target.PageSetup.RightMargin = Unit.FromPoint(source.MarginRightPoints);
        target.PageSetup.BottomMargin = Unit.FromPoint(source.MarginBottomPoints);
        target.PageSetup.LeftMargin = Unit.FromPoint(source.MarginLeftPoints);
    }

    private void RenderHeaderFooter(
        HeaderFooter target,
        IEnumerable<BuiltInParagraphModel> paragraphs)
    {
        foreach (var paragraph in paragraphs)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            RenderParagraph(paragraph, target.AddParagraph, null);
        }
    }

    private void RenderSectionParagraph(MigraSection section, BuiltInParagraphModel source)
    {
        RenderParagraph(source, section.AddParagraph, section.AddPageBreak);
    }

    private void RenderParagraph(
        BuiltInParagraphModel source,
        Func<MigraParagraph> createParagraph,
        Action? pageBreak)
    {
        var target = createParagraph();
        ApplyParagraphFormat(target, source.Format);

        foreach (var inline in source.Inlines)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            switch (inline)
            {
                case BuiltInTextModel text:
                    AddText(target, text);
                    break;
                case BuiltInTabModel:
                    target.AddTab();
                    break;
                case BuiltInBreakModel { IsPageBreak: true }:
                    if (pageBreak is null)
                    {
                        target.AddLineBreak();
                    }
                    else
                    {
                        pageBreak();
                        target = createParagraph();
                        ApplyParagraphFormat(target, source.Format);
                    }
                    break;
                case BuiltInBreakModel:
                    target.AddLineBreak();
                    break;
                case BuiltInImageModel image:
                    AddImage(target, image);
                    break;
                case BuiltInPageNumberModel:
                    target.AddPageField();
                    break;
            }
        }
    }

    private static void ApplyParagraphFormat(MigraParagraph target, BuiltInParagraphFormat source)
    {
        target.Format.Alignment = source.Alignment switch
        {
            BuiltInParagraphAlignment.Center => ParagraphAlignment.Center,
            BuiltInParagraphAlignment.Right => ParagraphAlignment.Right,
            BuiltInParagraphAlignment.Justify => ParagraphAlignment.Justify,
            _ => ParagraphAlignment.Left
        };

        if (source.SpaceBeforePoints is double before)
        {
            target.Format.SpaceBefore = Unit.FromPoint(before);
        }

        if (source.SpaceAfterPoints is double after)
        {
            target.Format.SpaceAfter = Unit.FromPoint(after);
        }

        if (source.LineSpacingPoints is double line)
        {
            target.Format.LineSpacing = Unit.FromPoint(line);
            target.Format.LineSpacingRule = LineSpacingRule.AtLeast;
        }

        if (source.LeftIndentPoints is double left)
        {
            target.Format.LeftIndent = Unit.FromPoint(left);
        }

        if (source.RightIndentPoints is double right)
        {
            target.Format.RightIndent = Unit.FromPoint(right);
        }

        if (source.FirstLineIndentPoints is double first)
        {
            target.Format.FirstLineIndent = Unit.FromPoint(first);
        }

        target.Format.PageBreakBefore = source.PageBreakBefore;
        target.Format.KeepWithNext = source.KeepWithNext;
    }

    private void AddText(MigraParagraph target, BuiltInTextModel source)
    {
        var formatted = target.AddFormattedText(source.Text);
        var family = source.Format.FontFamily;
        if (string.IsNullOrWhiteSpace(family))
        {
            family = ResolveFallbackFamily();
            if (_reportedDefaultFonts.Add(family))
            {
                _diagnostics.Add(new DocxConversionDiagnostic(
                    "DOCX_FONT_DEFAULTED",
                    $"BuiltIn used '{family}' for text without an explicit font.",
                    DocxConversionDiagnosticSeverity.Information));
            }
        }

        if (!string.IsNullOrWhiteSpace(family))
        {
            formatted.Font.Name = family;
        }

        if (source.Format.FontSizePoints is double size)
        {
            formatted.Font.Size = Unit.FromPoint(size);
        }

        formatted.Font.Bold = source.Format.Bold;
        formatted.Font.Italic = source.Format.Italic;
        formatted.Font.Underline = source.Format.Underline ? Underline.Single : Underline.None;
        if (TryParseColor(source.Format.ColorHex, out var color))
        {
            formatted.Font.Color = color;
        }
    }

    private string ResolveFallbackFamily()
    {
        foreach (var family in _options.FontFallbackFamilies)
        {
            if (!string.IsNullOrWhiteSpace(family))
            {
                return family;
            }
        }

        return GlobalFontSettings.FontResolver?.DefaultFontName ?? "Arial";
    }

    private void AddImage(MigraParagraph target, BuiltInImageModel source)
    {
        if (source.Data.Length == 0)
        {
            return;
        }

        var name = "docx-image-" + (++_imageIndex).ToString(CultureInfo.InvariantCulture) +
            Path.GetExtension(source.Name);
        var data = source.Data;
        var image = target.AddImage(MigraImageSource.FromBinary(name, () => data));
        image.LockAspectRatio = true;
        if (source.WidthPoints is double width)
        {
            image.Width = Unit.FromPoint(width);
        }

        if (source.HeightPoints is double height)
        {
            image.Height = Unit.FromPoint(height);
        }
    }

    private void RenderTable(MigraSection section, BuiltInTableModel source)
    {
        var table = section.AddTable();
        table.Borders.Visible = source.HasBorders;
        foreach (var width in source.ColumnWidthsPoints)
        {
            table.AddColumn(Unit.FromPoint(Math.Max(1, width)));
        }

        foreach (var sourceRow in source.Rows)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var row = table.AddRow();
            row.HeadingFormat = sourceRow.IsHeader;
            var targetColumn = 0;
            foreach (var sourceCell in sourceRow.Cells)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                if (targetColumn >= table.Columns.Count)
                {
                    break;
                }

                var cell = row.Cells[targetColumn];
                cell.MergeRight = Math.Min(
                    sourceCell.ColumnSpan - 1,
                    table.Columns.Count - targetColumn - 1);
                if (TryParseColor(sourceCell.ShadingHex, out var shading))
                {
                    cell.Shading.Color = shading;
                }

                foreach (var paragraph in sourceCell.Paragraphs)
                {
                    RenderParagraph(paragraph, cell.AddParagraph, null);
                }

                targetColumn += sourceCell.ColumnSpan;
            }
        }
    }

    private static bool TryParseColor(string? value, out Color color)
    {
        color = Colors.Black;
        if (value is null || string.IsNullOrWhiteSpace(value) || value.Length != 6 ||
            !uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
        {
            return false;
        }

        color = new Color(
            (byte)((rgb >> 16) & 0xff),
            (byte)((rgb >> 8) & 0xff),
            (byte)(rgb & 0xff));
        return true;
    }
}
