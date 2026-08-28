using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Internal;

internal static class DocxSemanticTextExtractor
{
    internal static IReadOnlyList<string> ExtractPages(
        string inputPath,
        DocxConversionOptions options,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
        {
            throw new ArgumentException("A DOCX input path is required.", nameof(inputPath));
        }

        DocxPackageValidator.Validate(inputPath, options);
        using var document = WordprocessingDocument.Open(inputPath, false, new OpenSettings
        {
            AutoSave = false
        });
        var mainPart = document.MainDocumentPart ??
            throw new InvalidDataException("The DOCX package has no main document part.");
        var body = mainPart.Document?.Body ??
            throw new InvalidDataException("The DOCX package has no document body.");

        var pages = new List<StringBuilder> { new StringBuilder() };
        AppendParagraphs(body, pages, allowPageBreaks: true, cancellationToken);

        var headerText = JoinPartText(
            mainPart.HeaderParts.Select(part => part.Header),
            cancellationToken);
        var footerText = JoinPartText(
            mainPart.FooterParts.Select(part => part.Footer),
            cancellationToken);
        var supplementalText = JoinPartText(
            new OpenXmlElement?[]
            {
                mainPart.FootnotesPart?.Footnotes,
                mainPart.EndnotesPart?.Endnotes,
                mainPart.WordprocessingCommentsPart?.Comments
            },
            cancellationToken);

        if (!string.IsNullOrEmpty(supplementalText))
        {
            AppendSeparated(pages[pages.Count - 1], supplementalText);
        }

        return pages
            .Select(page => CombinePageText(headerText, page.ToString(), footerText))
            .ToArray();
    }

    private static string JoinPartText(
        IEnumerable<OpenXmlElement?> roots,
        CancellationToken cancellationToken)
    {
        var result = new StringBuilder();
        foreach (var root in roots.Where(root => root is not null))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var pages = new List<StringBuilder> { new StringBuilder() };
            AppendParagraphs(root!, pages, allowPageBreaks: false, cancellationToken);
            AppendSeparated(result, string.Join("\n", pages.Select(page => page.ToString())));
        }

        return result.ToString();
    }

    private static void AppendParagraphs(
        OpenXmlElement root,
        IList<StringBuilder> pages,
        bool allowPageBreaks,
        CancellationToken cancellationToken)
    {
        foreach (var paragraph in root.Descendants<W.Paragraph>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (allowPageBreaks && IsPageBreakBefore(paragraph))
            {
                StartPage(pages);
            }

            foreach (var node in paragraph.Descendants())
            {
                cancellationToken.ThrowIfCancellationRequested();
                switch (node)
                {
                    case W.Text text:
                        pages[pages.Count - 1].Append(text.Text);
                        break;
                    case W.TabChar:
                        pages[pages.Count - 1].Append('\t');
                        break;
                    case W.CarriageReturn:
                        pages[pages.Count - 1].AppendLine();
                        break;
                    case W.Break lineBreak when
                        allowPageBreaks && lineBreak.Type?.Value == W.BreakValues.Page:
                        StartPage(pages);
                        break;
                    case W.LastRenderedPageBreak when allowPageBreaks:
                        StartPage(pages);
                        break;
                    case W.Break:
                        pages[pages.Count - 1].AppendLine();
                        break;
                }
            }

            pages[pages.Count - 1].AppendLine();
        }
    }

    private static bool IsPageBreakBefore(W.Paragraph paragraph)
    {
        var pageBreak = paragraph.ParagraphProperties?.PageBreakBefore;
        return pageBreak is not null && (pageBreak.Val?.Value ?? true);
    }

    private static void StartPage(IList<StringBuilder> pages)
    {
        if (pages[pages.Count - 1].Length == 0)
        {
            return;
        }

        pages.Add(new StringBuilder());
    }

    private static string CombinePageText(string header, string body, string footer)
    {
        var result = new StringBuilder();
        AppendSeparated(result, header);
        AppendSeparated(result, body);
        AppendSeparated(result, footer);
        return result.ToString();
    }

    private static void AppendSeparated(StringBuilder target, string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return;
        }

        if (target.Length > 0 && target[target.Length - 1] != '\n')
        {
            target.AppendLine();
        }

        target.Append(value);
    }
}
