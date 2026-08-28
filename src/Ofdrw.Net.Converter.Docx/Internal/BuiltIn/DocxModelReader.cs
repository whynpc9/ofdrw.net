using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using SixLabors.ImageSharp;
using WpParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WpRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WpTable = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

internal sealed class DocxModelReader
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly DocxConversionOptions _options;
    private readonly IList<DocxConversionDiagnostic> _diagnostics;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<string, Style> _styles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _numberingCounters = new(StringComparer.Ordinal);
    private MainDocumentPart _mainPart = null!;
    private int _elementCount;

    internal DocxModelReader(
        DocxConversionOptions options,
        IList<DocxConversionDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        _options = options;
        _diagnostics = diagnostics;
        _cancellationToken = cancellationToken;
    }

    internal BuiltInDocumentModel Read(string inputPath)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        DocxPackageValidator.Validate(inputPath, _options);

        using var document = WordprocessingDocument.Open(inputPath, false, new OpenSettings
        {
            AutoSave = false
        });

        _mainPart = document.MainDocumentPart ??
            throw new InvalidDataException("The DOCX package has no main document part.");
        var body = _mainPart.Document?.Body ??
            throw new InvalidDataException("The DOCX package has no document body.");

        CountElements(body);
        LoadStyles();

        var model = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel();
        model.Sections.Add(section);

        foreach (var element in body.ChildElements)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (element is WpParagraph paragraph)
            {
                section.Blocks.Add(ReadParagraph(paragraph, _mainPart));
                var paragraphSection = paragraph.ParagraphProperties?.SectionProperties;
                if (paragraphSection is not null)
                {
                    ApplySectionProperties(section, paragraphSection);
                    AddHeadersAndFooters(section, paragraphSection);
                    section = new BuiltInSectionModel();
                    model.Sections.Add(section);
                }

                continue;
            }

            if (element is WpTable table)
            {
                section.Blocks.Add(ReadTable(table));
                continue;
            }

            if (element is SectionProperties sectionProperties)
            {
                ApplySectionProperties(section, sectionProperties);
                AddHeadersAndFooters(section, sectionProperties);
                continue;
            }

            ReportUnsupported("DOCX_BLOCK_UNSUPPORTED", element.LocalName);
            if (_options.UnsupportedFeatureBehavior == UnsupportedDocxFeatureBehavior.Placeholder)
            {
                section.Blocks.Add(CreateUnsupportedParagraph(element.LocalName));
            }
        }

        if (model.Sections.Count > 1 &&
            model.Sections[model.Sections.Count - 1].Blocks.Count == 0)
        {
            model.Sections.RemoveAt(model.Sections.Count - 1);
        }

        return model;
    }

    private void CountElements(OpenXmlElement root)
    {
        foreach (var _ in root.Descendants())
        {
            _cancellationToken.ThrowIfCancellationRequested();
            _elementCount++;
            if (_elementCount > _options.MaxDocumentElements)
            {
                throw new InvalidDataException(
                    $"The DOCX contains more than {_options.MaxDocumentElements} XML elements.");
            }
        }
    }

    private void LoadStyles()
    {
        var styles = _mainPart.StyleDefinitionsPart?.Styles;
        if (styles is null)
        {
            return;
        }

        foreach (var style in styles.Elements<Style>())
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var id = style.StyleId?.Value;
            if (!string.IsNullOrWhiteSpace(id))
            {
                _styles[id!] = style;
            }
        }
    }

    private BuiltInParagraphModel ReadParagraph(WpParagraph paragraph, OpenXmlPart sourcePart)
    {
        var model = new BuiltInParagraphModel();
        ApplyParagraphFormat(model.Format, paragraph.ParagraphProperties);
        model.Format.ListMarker = ResolveListMarker(paragraph.ParagraphProperties);

        if (!string.IsNullOrEmpty(model.Format.ListMarker))
        {
            model.Inlines.Add(new BuiltInTextModel
            {
                Text = model.Format.ListMarker + " "
            });
        }

        var renderedSimpleFields = new HashSet<SimpleField>();
        var skippingPageFieldResult = false;
        foreach (var run in paragraph.Descendants<WpRun>())
        {
            _cancellationToken.ThrowIfCancellationRequested();

            var simpleField = run.Ancestors<SimpleField>().FirstOrDefault();
            if (simpleField?.Instruction?.Value?.IndexOf(
                    "PAGE",
                    StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (renderedSimpleFields.Add(simpleField))
                {
                    model.Inlines.Add(new BuiltInPageNumberModel());
                }

                continue;
            }

            var fieldCharacterType = run.GetFirstChild<FieldChar>()?.FieldCharType?.Value;
            if (skippingPageFieldResult)
            {
                if (fieldCharacterType == FieldCharValues.End)
                {
                    skippingPageFieldResult = false;
                }

                continue;
            }

            ReadRun(run, sourcePart, model);
            if (run.Elements<FieldCode>().Any(fieldCode =>
                    fieldCode.Text?.IndexOf("PAGE", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                skippingPageFieldResult = true;
            }
        }

        if (model.Inlines.Count == 0)
        {
            model.Inlines.Add(new BuiltInTextModel { Text = string.Empty });
        }

        return model;
    }

    private void ReadRun(WpRun run, OpenXmlPart sourcePart, BuiltInParagraphModel paragraph)
    {
        foreach (var child in run.ChildElements)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            switch (child)
            {
                case Text text:
                    var textModel = new BuiltInTextModel { Text = text.Text ?? string.Empty };
                    ApplyRunFormat(textModel.Format, run);
                    paragraph.Inlines.Add(textModel);
                    break;
                case TabChar:
                    paragraph.Inlines.Add(new BuiltInTabModel());
                    break;
                case Break br:
                    paragraph.Inlines.Add(new BuiltInBreakModel
                    {
                        IsPageBreak = br.Type?.Value == BreakValues.Page
                    });
                    break;
                case CarriageReturn:
                    paragraph.Inlines.Add(new BuiltInBreakModel());
                    break;
                case Drawing drawing:
                    ReadDrawing(drawing, sourcePart, paragraph);
                    break;
                case FieldCode fieldCode when
                    fieldCode.Text?.IndexOf("PAGE", StringComparison.OrdinalIgnoreCase) >= 0:
                    paragraph.Inlines.Add(new BuiltInPageNumberModel());
                    break;
                case FieldCode:
                    ReportUnsupported("DOCX_FIELD_UNSUPPORTED", "field code");
                    AddUnsupportedPlaceholder(paragraph, "field code");
                    break;
                case RunProperties:
                case LastRenderedPageBreak:
                    break;
                case FootnoteReference:
                    ReportUnsupported("DOCX_FOOTNOTE_UNSUPPORTED", "footnote");
                    AddUnsupportedPlaceholder(paragraph, "footnote");
                    break;
                case EmbeddedObject:
                case Picture:
                    ReportUnsupported("DOCX_DRAWING_UNSUPPORTED", child.LocalName);
                    AddUnsupportedPlaceholder(paragraph, child.LocalName);
                    break;
                case OpenXmlElement unsupported when
                    unsupported.LocalName is "commentReference" or "endnoteReference" or "sym":
                    ReportUnsupported("DOCX_INLINE_UNSUPPORTED", unsupported.LocalName);
                    AddUnsupportedPlaceholder(paragraph, unsupported.LocalName);
                    break;
            }
        }
    }

    private BuiltInTableModel ReadTable(WpTable table)
    {
        var model = new BuiltInTableModel
        {
            HasBorders = table.TableProperties?.TableBorders is not null
        };

        var grid = table.GetFirstChild<TableGrid>();
        if (grid is not null)
        {
            foreach (var column in grid.Elements<GridColumn>())
            {
                model.ColumnWidthsPoints.Add(TwipsToPoints(column.Width?.Value, 1440));
            }
        }

        foreach (var row in table.Elements<TableRow>())
        {
            _cancellationToken.ThrowIfCancellationRequested();
            var rowModel = new BuiltInTableRowModel
            {
                IsHeader = row.TableRowProperties?.GetFirstChild<TableHeader>() is not null
            };

            foreach (var cell in row.Elements<TableCell>())
            {
                _cancellationToken.ThrowIfCancellationRequested();
                var cellModel = new BuiltInTableCellModel
                {
                    ColumnSpan = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1),
                    ShadingHex = cell.TableCellProperties?.Shading?.Fill?.Value
                };

                foreach (var paragraph in cell.Elements<WpParagraph>())
                {
                    cellModel.Paragraphs.Add(ReadParagraph(paragraph, _mainPart));
                }

                if (cell.Elements<WpTable>().Any())
                {
                    ReportUnsupported("DOCX_NESTED_TABLE_UNSUPPORTED", "nested table");
                    if (_options.UnsupportedFeatureBehavior == UnsupportedDocxFeatureBehavior.Placeholder)
                    {
                        cellModel.Paragraphs.Add(CreateUnsupportedParagraph("nested table"));
                    }
                }

                if (cellModel.Paragraphs.Count == 0)
                {
                    cellModel.Paragraphs.Add(new BuiltInParagraphModel());
                }

                if (cell.TableCellProperties?.VerticalMerge is not null)
                {
                    ReportUnsupported("DOCX_VERTICAL_MERGE_DEGRADED", "vertical cell merge");
                }

                rowModel.Cells.Add(cellModel);
            }

            model.Rows.Add(rowModel);
        }

        if (model.ColumnWidthsPoints.Count == 0)
        {
            var count = Math.Max(1, model.Rows.Select(x => x.Cells.Sum(c => c.ColumnSpan)).DefaultIfEmpty(1).Max());
            for (var index = 0; index < count; index++)
            {
                model.ColumnWidthsPoints.Add(72);
            }
        }

        return model;
    }

    private void ReadDrawing(Drawing drawing, OpenXmlPart sourcePart, BuiltInParagraphModel paragraph)
    {
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        var relationshipId = blip?.Embed?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            ReportUnsupported("DOCX_DRAWING_UNSUPPORTED", "drawing without embedded image");
            AddUnsupportedPlaceholder(paragraph, "drawing without embedded image");
            return;
        }

        try
        {
            if (sourcePart.GetPartById(relationshipId!) is not ImagePart imagePart)
            {
                ReportUnsupported("DOCX_DRAWING_UNSUPPORTED", "non-image drawing");
                AddUnsupportedPlaceholder(paragraph, "non-image drawing");
                return;
            }

            using var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using var data = new MemoryStream();
            stream.CopyTo(data);
            if (data.Length > _options.MaxEmbeddedImageBytes)
            {
                throw new InvalidDataException(
                    $"An embedded DOCX image exceeds {_options.MaxEmbeddedImageBytes} bytes.");
            }

            var imageBytes = data.ToArray();
            var imageInfo = Image.Identify(imageBytes);
            if (imageInfo is null)
            {
                ReportUnsupported("DOCX_IMAGE_FORMAT_UNSUPPORTED", imagePart.ContentType);
                AddUnsupportedPlaceholder(paragraph, imagePart.ContentType);
                return;
            }

            var pixelCount = checked((long)imageInfo.Width * imageInfo.Height);
            if (pixelCount > _options.MaxEmbeddedImagePixels)
            {
                throw new InvalidDataException(
                    $"An embedded DOCX image exceeds {_options.MaxEmbeddedImagePixels} decoded pixels.");
            }

            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            paragraph.Inlines.Add(new BuiltInImageModel
            {
                Data = imageBytes,
                Name = "image" + GuessImageExtension(imagePart.ContentType),
                WidthPoints = EmuToPoints(extent?.Cx?.Value),
                HeightPoints = EmuToPoints(extent?.Cy?.Value)
            });
        }
        catch (KeyNotFoundException)
        {
            ReportUnsupported("DOCX_IMAGE_RELATIONSHIP_MISSING", "image relationship");
            AddUnsupportedPlaceholder(paragraph, "image relationship");
        }
        catch (UnknownImageFormatException)
        {
            ReportUnsupported("DOCX_IMAGE_FORMAT_UNSUPPORTED", "embedded image format");
            AddUnsupportedPlaceholder(paragraph, "embedded image format");
        }
    }

    private void ApplyParagraphFormat(BuiltInParagraphFormat target, ParagraphProperties? direct)
    {
        var layers = GetParagraphPropertyLayers(direct);
        foreach (var properties in layers)
        {
            if (properties.Justification?.Val?.Value is JustificationValues alignment)
            {
                if (alignment == JustificationValues.Center)
                {
                    target.Alignment = BuiltInParagraphAlignment.Center;
                }
                else if (alignment == JustificationValues.Right)
                {
                    target.Alignment = BuiltInParagraphAlignment.Right;
                }
                else if (alignment == JustificationValues.Both ||
                         alignment == JustificationValues.Distribute)
                {
                    target.Alignment = BuiltInParagraphAlignment.Justify;
                }
                else
                {
                    target.Alignment = BuiltInParagraphAlignment.Left;
                }
            }

            var spacing = properties.SpacingBetweenLines;
            if (spacing is not null)
            {
                target.SpaceBeforePoints = TwipsToNullablePoints(spacing.Before?.Value);
                target.SpaceAfterPoints = TwipsToNullablePoints(spacing.After?.Value);
                target.LineSpacingPoints = TwipsToNullablePoints(spacing.Line?.Value);
            }

            var indentation = properties.Indentation;
            if (indentation is not null)
            {
                target.LeftIndentPoints = TwipsToNullablePoints(indentation.Left?.Value ?? indentation.Start?.Value);
                target.RightIndentPoints = TwipsToNullablePoints(indentation.Right?.Value ?? indentation.End?.Value);
                target.FirstLineIndentPoints = TwipsToNullablePoints(indentation.FirstLine?.Value);
                if (indentation.Hanging?.Value is string hanging &&
                    double.TryParse(hanging, NumberStyles.Number, CultureInfo.InvariantCulture, out var hangingTwips))
                {
                    target.FirstLineIndentPoints = -hangingTwips / 20d;
                }
            }

            target.PageBreakBefore |= IsOn(properties.PageBreakBefore);
            target.KeepWithNext |= IsOn(properties.KeepNext);
        }
    }

    private void ApplyRunFormat(
        BuiltInTextFormat target,
        WpRun run)
    {
        foreach (var properties in GetRunPropertyLayers(
                     run.Ancestors<WpParagraph>().FirstOrDefault(),
                     run.RunProperties))
        {
            var fonts = properties.RunFonts;
            var family = fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
            if (!string.IsNullOrWhiteSpace(family))
            {
                target.FontFamily = family;
            }

            if (properties.FontSize?.Val?.Value is string halfPoints &&
                double.TryParse(halfPoints, NumberStyles.Number, CultureInfo.InvariantCulture, out var size))
            {
                target.FontSizePoints = Math.Max(1, size / 2d);
            }

            target.Bold = MergeOnOff(target.Bold, properties.Bold);
            target.Italic = MergeOnOff(target.Italic, properties.Italic);
            if (properties.Underline?.Val?.Value is UnderlineValues underline)
            {
                target.Underline = underline != UnderlineValues.None;
            }

            var color = properties.Color?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(color) && color != "auto")
            {
                target.ColorHex = color;
            }
        }

    }

    private IEnumerable<ParagraphProperties> GetParagraphPropertyLayers(ParagraphProperties? direct)
    {
        var defaults = _mainPart.StyleDefinitionsPart?.Styles?.DocDefaults?
            .ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        if (defaults is not null)
        {
            yield return new ParagraphProperties(defaults.OuterXml);
        }

        var styleId = direct?.ParagraphStyleId?.Val?.Value;
        foreach (var style in ResolveStyleChain(styleId))
        {
            if (style.StyleParagraphProperties is not null)
            {
                yield return new ParagraphProperties(style.StyleParagraphProperties.OuterXml);
            }
        }

        if (direct is not null)
        {
            yield return direct;
        }
    }

    private IEnumerable<RunProperties> GetRunPropertyLayers(WpParagraph? paragraph, RunProperties? direct)
    {
        var defaults = _mainPart.StyleDefinitionsPart?.Styles?.DocDefaults?
            .RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (defaults is not null)
        {
            yield return new RunProperties(defaults.OuterXml);
        }

        var paragraphStyleId = paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        foreach (var style in ResolveStyleChain(paragraphStyleId))
        {
            if (style.StyleRunProperties is not null)
            {
                yield return new RunProperties(style.StyleRunProperties.OuterXml);
            }
        }

        var characterStyleId = direct?.RunStyle?.Val?.Value;
        foreach (var style in ResolveStyleChain(characterStyleId))
        {
            if (style.StyleRunProperties is not null)
            {
                yield return new RunProperties(style.StyleRunProperties.OuterXml);
            }
        }

        if (direct is not null)
        {
            yield return direct;
        }
    }

    private IEnumerable<Style> ResolveStyleChain(string? styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId))
        {
            yield break;
        }

        var chain = new Stack<Style>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        var currentId = styleId;
        while (!string.IsNullOrWhiteSpace(currentId) &&
               visited.Add(currentId!) &&
               _styles.TryGetValue(currentId!, out var style))
        {
            chain.Push(style);
            currentId = style.BasedOn?.Val?.Value;
        }

        foreach (var style in chain)
        {
            yield return style;
        }
    }

    private string? ResolveListMarker(ParagraphProperties? properties)
    {
        NumberingProperties? numbering = null;
        foreach (var layer in GetParagraphPropertyLayers(properties))
        {
            if (layer.NumberingProperties is not null)
            {
                numbering = layer.NumberingProperties;
            }
        }

        var numberingId = numbering?.NumberingId?.Val?.Value;
        var level = numbering?.NumberingLevelReference?.Val?.Value ?? 0;
        if (numberingId is null)
        {
            return null;
        }

        var key = numberingId.Value.ToString(CultureInfo.InvariantCulture) + ":" + level;
        _numberingCounters.TryGetValue(key, out var count);
        count++;
        _numberingCounters[key] = count;

        var numberingPart = _mainPart.NumberingDefinitionsPart?.Numbering;
        var instance = numberingPart?.Elements<NumberingInstance>()
            .FirstOrDefault(x => x.NumberID?.Value == numberingId.Value);
        var abstractId = instance?.AbstractNumId?.Val?.Value;
        var abstractNumbering = numberingPart?.Elements<AbstractNum>()
            .FirstOrDefault(x => x.AbstractNumberId?.Value == abstractId);
        var levelDefinition = abstractNumbering?.Elements<Level>()
            .FirstOrDefault(x => x.LevelIndex?.Value == level);
        var format = levelDefinition?.NumberingFormat?.Val?.Value;
        if (format == NumberFormatValues.Bullet)
        {
            return levelDefinition?.LevelText?.Val?.Value ?? "•";
        }

        var pattern = levelDefinition?.LevelText?.Val?.Value ?? "%1.";
        return pattern.Replace("%" + (level + 1), count.ToString(CultureInfo.InvariantCulture));
    }

    private void ApplySectionProperties(BuiltInSectionModel target, SectionProperties properties)
    {
        var size = properties.GetFirstChild<PageSize>();
        if (size is not null)
        {
            target.PageWidthPoints = TwipsToPoints(size.Width?.Value.ToString(), 11906);
            target.PageHeightPoints = TwipsToPoints(size.Height?.Value.ToString(), 16838);
        }

        var margins = properties.GetFirstChild<PageMargin>();
        if (margins is not null)
        {
            target.MarginTopPoints = TwipsToPoints(margins.Top?.Value.ToString(), 1440);
            target.MarginRightPoints = TwipsToPoints(margins.Right?.Value.ToString(), 1440);
            target.MarginBottomPoints = TwipsToPoints(margins.Bottom?.Value.ToString(), 1440);
            target.MarginLeftPoints = TwipsToPoints(margins.Left?.Value.ToString(), 1440);
        }
    }

    private void AddHeadersAndFooters(BuiltInSectionModel target, SectionProperties properties)
    {
        var headerReference = properties.Elements<HeaderReference>()
            .FirstOrDefault(x => x.Type?.Value == HeaderFooterValues.Default) ??
            properties.Elements<HeaderReference>().FirstOrDefault();
        if (headerReference?.Id?.Value is string headerId &&
            _mainPart.GetPartById(headerId) is HeaderPart headerPart)
        {
            foreach (var paragraph in headerPart.Header?.Elements<WpParagraph>() ?? [])
            {
                target.Headers.Add(ReadParagraph(paragraph, headerPart));
            }
        }

        var footerReference = properties.Elements<FooterReference>()
            .FirstOrDefault(x => x.Type?.Value == HeaderFooterValues.Default) ??
            properties.Elements<FooterReference>().FirstOrDefault();
        if (footerReference?.Id?.Value is string footerId &&
            _mainPart.GetPartById(footerId) is FooterPart footerPart)
        {
            foreach (var paragraph in footerPart.Footer?.Elements<WpParagraph>() ?? [])
            {
                target.Footers.Add(ReadParagraph(paragraph, footerPart));
            }
        }
    }

    private void ReportUnsupported(string code, string feature)
    {
        if (_options.UnsupportedFeatureBehavior == UnsupportedDocxFeatureBehavior.Throw)
        {
            throw new NotSupportedException($"BuiltIn cannot render DOCX feature '{feature}'.");
        }

        _diagnostics.Add(new DocxConversionDiagnostic(
            code,
            $"BuiltIn degraded unsupported DOCX feature '{feature}'."));
    }

    private void AddUnsupportedPlaceholder(BuiltInParagraphModel paragraph, string feature)
    {
        if (_options.UnsupportedFeatureBehavior == UnsupportedDocxFeatureBehavior.Placeholder)
        {
            paragraph.Inlines.Add(new BuiltInTextModel
            {
                Text = $"[Unsupported DOCX feature: {feature}]"
            });
        }
    }

    private static BuiltInParagraphModel CreateUnsupportedParagraph(string feature)
    {
        var paragraph = new BuiltInParagraphModel();
        paragraph.Inlines.Add(new BuiltInTextModel
        {
            Text = $"[Unsupported DOCX feature: {feature}]"
        });
        return paragraph;
    }

    private static bool IsOn(OpenXmlElement? element)
    {
        if (element is null)
        {
            return false;
        }

        string? value = null;
        foreach (var attribute in element.GetAttributes())
        {
            if (attribute.LocalName == "val" && attribute.NamespaceUri == WordNamespace)
            {
                value = attribute.Value;
                break;
            }
        }

        return string.IsNullOrEmpty(value) ||
            !(value == "0" ||
              string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) ||
              value == "off");
    }

    private static bool MergeOnOff(bool current, OpenXmlElement? element)
    {
        return element is null ? current : IsOn(element);
    }

    private static double TwipsToPoints(string? value, double fallbackTwips)
    {
        return double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var twips)
            ? twips / 20d
            : fallbackTwips / 20d;
    }

    private static double? TwipsToNullablePoints(string? value)
    {
        return double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var twips)
            ? twips / 20d
            : null;
    }

    private static double? EmuToPoints(long? value)
    {
        return value is null ? null : value.Value / 12700d;
    }

    private static string GuessImageExtension(string contentType)
    {
        return contentType switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            _ => ".bin"
        };
    }
}
