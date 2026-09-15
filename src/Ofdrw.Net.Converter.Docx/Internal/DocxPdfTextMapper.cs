using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using UglyToad.PdfPig;

namespace Ofdrw.Net.Converter.Docx.Internal;

internal sealed class DocxPageTextMap
{
    internal List<string> Pages { get; } = new();
    internal List<DocxConversionDiagnostic> Diagnostics { get; } = new();
    internal bool HasUnplacedSupplementalText { get; set; }
}

/// <summary>
/// Uses rendered PDF text only to locate original Word characters on pages.
/// Every emitted text fragment is copied/evaluated from OpenXML, never from PDF.
/// </summary>
internal static class DocxPdfTextMapper
{
    internal static DocxPageTextMap Map(
        DocxSourceDocument source, string pdfPath, int maximumPages, bool supplementalAfterBody,
        CancellationToken cancellationToken)
    {
        using var pdf = PdfDocument.Open(pdfPath);
        if (pdf.NumberOfPages <= 0 || pdf.NumberOfPages > maximumPages)
            throw new InvalidDataException("DOCX rendered page count is outside the configured limits.");
        var fullText = new StringBuilder();
        var fullPages = new List<int>();
        var bodyText = new StringBuilder();
        var bodyPages = new List<int>();
        var topMargin = source.Sections.Select(section => section.TopMargin).DefaultIfEmpty(72).Min();
        var bottomMargin = source.Sections.Select(section => section.BottomMargin).DefaultIfEmpty(72).Min();
        for (var pageIndex = 0; pageIndex < pdf.NumberOfPages; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = pdf.GetPage(pageIndex + 1);
            foreach (var letter in page.Letters)
            {
                Append(fullText, fullPages, letter.Value, pageIndex);
                if (letter.Location.Y >= bottomMargin - 1 && letter.Location.Y <= page.Height - topMargin + 1)
                    Append(bodyText, bodyPages, letter.Value, pageIndex);
            }
        }
        var full = Normalize(fullText.ToString(), fullPages);
        var body = Normalize(bodyText.ToString(), bodyPages);
        var result = new DocxPageTextMap();
        List<PartMapping> mappings;
        try { mappings = MapBody(source, body, pdf.NumberOfPages, cancellationToken); }
        catch (InvalidDataException)
        {
            // Floating text can legitimately extend into page margins. Retry
            // against actual page positions rather than inventing semantic pages.
            mappings = MapBody(source, full, pdf.NumberOfPages, cancellationToken);
            result.Diagnostics.Add(new DocxConversionDiagnostic("DOCX_TEXT_MAPPING_INCLUDED_MARGINS",
                "Original body text was aligned using rendered text including page margins.", DocxConversionDiagnosticSeverity.Information));
        }

        var sectionForPage = new int[pdf.NumberOfPages];
        var assigned = new bool[pdf.NumberOfPages];
        foreach (var mapping in mappings)
        {
            foreach (var page in mapping.Pages.Distinct())
            {
                if (!assigned[page]) { sectionForPage[page] = mapping.Part.SectionIndex; assigned[page] = true; }
            }
        }
        for (var page = 1; page < pdf.NumberOfPages; page++)
            if (!assigned[page]) sectionForPage[page] = sectionForPage[page - 1];
        var sectionCounts = sectionForPage.GroupBy(index => index).ToDictionary(group => group.Key, group => group.Count());
        var firstSectionPages = sectionForPage.Select((section, page) => (section, page))
            .GroupBy(value => value.section).ToDictionary(group => group.Key, group => group.First().page);
        var numbers = new int[pdf.NumberOfPages];
        for (var page = 0; page < numbers.Length; page++)
        {
            var section = source.Sections[sectionForPage[page]];
            numbers[page] = firstSectionPages[sectionForPage[page]] == page && section.PageNumberStart.HasValue
                ? section.PageNumberStart.Value : page == 0 ? 1 : numbers[page - 1] + 1;
        }
        var texts = Enumerable.Range(0, pdf.NumberOfPages).Select(_ => new StringBuilder()).ToArray();
        for (var page = 0; page < texts.Length; page++) AddDecoration(source.Sections[sectionForPage[page]].Headers, page);
        var references = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var mapping in mappings)
        {
            Emit(mapping.Part, mapping.Pages, mapping.FallbackPage);
            foreach (var reference in mapping.Part.References)
            {
                var key = reference.Kind + ":" + reference.Id;
                if (!references.ContainsKey(key)) references[key] = PageAt(mapping.Pages, reference.Offset, mapping.FallbackPage);
            }
        }
        var supplementalCursor = 0;
        foreach (var part in source.Supplemental)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var literal = Normalize(part.LiteralText);
            var start = supplementalCursor;
            if (supplementalAfterBody)
            {
                var label = Normalize($"[{part.Kind} {part.Id}]").Text;
                var labelIndex = full.Text.LastIndexOf(label, StringComparison.Ordinal);
                if (labelIndex >= 0) start = labelIndex + label.Length;
            }
            try
            {
                var cursor = start;
                var mapped = Align(literal, full, ref cursor);
                var pages = ExpandOriginalPages(part.LiteralText, literal.Offsets, mapped, full.Pages,
                    references.TryGetValue(part.Kind + ":" + part.Id, out var referencePage) ? referencePage : texts.Length - 1);
                Emit(part, pages, texts.Length - 1);
                supplementalCursor = cursor;
            }
            catch (InvalidDataException)
            {
                if (references.TryGetValue(part.Kind + ":" + part.Id, out var referencePage))
                {
                    // Comments need not be printed by Word. Their XML reference
                    // gives an explicit semantic anchor even without a PDF glyph.
                    Emit(part, Array.Empty<int>(), referencePage);
                    result.Diagnostics.Add(new DocxConversionDiagnostic("DOCX_SUPPLEMENTAL_REFERENCE_ANCHORED",
                        $"{part.Kind} {part.Id} text is preserved at its original reference on page {referencePage + 1}.",
                        DocxConversionDiagnosticSeverity.Information));
                }
                else
                {
                    Emit(part, Array.Empty<int>(), texts.Length - 1);
                    result.HasUnplacedSupplementalText = true;
                    result.Diagnostics.Add(new DocxConversionDiagnostic("DOCX_SUPPLEMENTAL_DOCUMENT_SCOPE",
                        $"{part.Kind} {part.Id} has no rendered or reference anchor; its original text is retained as document-scoped text on the final page."));
                }
            }
        }
        for (var page = 0; page < texts.Length; page++) AddDecoration(source.Sections[sectionForPage[page]].Footers, page);
        result.Pages.AddRange(texts.Select(text => text.ToString()));
        return result;

        void Emit(DocxSourcePart part, IReadOnlyList<int> pages, int fallback)
        {
            var offset = 0;
            foreach (var segment in part.Segments)
            {
                if (segment.Field is not null)
                {
                    var page = PageAt(pages, offset, fallback);
                    texts[page].Append(DocxSourcePart.FieldValue(segment.Field.Value, numbers[page], pdf.NumberOfPages, sectionCounts[sectionForPage[page]]));
                }
                else
                {
                    foreach (var character in segment.Text) texts[PageAt(pages, offset++, fallback)].Append(character);
                }
            }
            texts[PageAt(pages, Math.Max(0, offset - 1), fallback)].AppendLine();
        }

        void AddDecoration(IReadOnlyDictionary<string, DocxSourcePart> variants, int page)
        {
            var section = source.Sections[sectionForPage[page]];
            var first = page == firstSectionPages[sectionForPage[page]];
            var kind = first && section.DifferentFirstPage ? "first" : section.DifferentEvenPages && numbers[page] % 2 == 0 ? "even" : "default";
            if (variants.TryGetValue(kind, out var part)) texts[page].AppendLine(part.Evaluate(numbers[page], pdf.NumberOfPages, sectionCounts[sectionForPage[page]]));
        }
    }

    private static List<PartMapping> MapBody(DocxSourceDocument source, NormalizedText target, int pageCount, CancellationToken cancellationToken)
    {
        var result = new List<PartMapping>();
        var cursor = 0;
        var previousPage = 0;
        foreach (var part in source.Sections.SelectMany(section => section.Paragraphs))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var text = part.LiteralText;
            var normalized = Normalize(text);
            var mapped = Align(normalized, target, ref cursor);
            var fallback = Math.Min(pageCount - 1, previousPage + (part.PageBreakBefore && normalized.Text.Length == 0 ? 1 : 0));
            var pages = ExpandOriginalPages(text, normalized.Offsets, mapped, target.Pages, fallback);
            result.Add(new PartMapping(part, pages, fallback));
            if (pages.Count > 0) previousPage = pages[pages.Count - 1];
        }
        return result;
    }

    private static IReadOnlyList<int> Align(NormalizedText source, NormalizedText target, ref int cursor)
    {
        var mapping = new List<int>(source.Text.Length);
        var position = 0;
        var hasAnchor = source.Text.Length == 0;
        var minimumAnchor = Math.Min(8, source.Text.Length);
        while (position < source.Text.Length)
        {
            var length = Math.Min(32, source.Text.Length - position);
            var found = -1;
            while (length > 0)
            {
                found = target.Text.IndexOf(source.Text.Substring(position, length), cursor, StringComparison.Ordinal);
                if (found >= 0) break;
                length = length == 1 ? 0 : Math.Max(1, length / 2);
            }
            if (found < 0)
                throw new InvalidDataException("DOCX_TEXT_PAGE_MAPPING_FAILED: original OpenXML text could not be aligned with rendered pages. No partial text layer was written.");
            cursor = found;
            var runStart = position;
            while (position < source.Text.Length && cursor < target.Text.Length && source.Text[position] == target.Text[cursor])
            {
                mapping.Add(cursor++);
                position++;
            }
            hasAnchor |= position - runStart >= minimumAnchor;
        }
        if (!hasAnchor)
            throw new InvalidDataException("DOCX_TEXT_PAGE_MAPPING_FAILED: rendered text provides no reliable original-text anchor.");
        return mapping;
    }

    private static IReadOnlyList<int> ExpandOriginalPages(string text, IReadOnlyList<int> offsets,
        IReadOnlyList<int> mapping, IReadOnlyList<int> targetPages, int fallback)
    {
        var result = new int[text.Length];
        var normalized = 0;
        var page = mapping.Count > 0 ? targetPages[mapping[0]] : fallback;
        for (var offset = 0; offset < text.Length; offset++)
        {
            while (normalized < offsets.Count && offsets[normalized] <= offset)
            {
                page = targetPages[mapping[normalized++]];
            }
            result[offset] = page;
        }
        return result;
    }

    private static int PageAt(IReadOnlyList<int> pages, int offset, int fallback) =>
        pages.Count == 0 ? fallback : pages[Math.Max(0, Math.Min(offset, pages.Count - 1))];

    private static void Append(StringBuilder text, IList<int> pages, string value, int page)
    {
        text.Append(value);
        for (var index = 0; index < value.Length; index++) pages.Add(page);
    }

    private static NormalizedText Normalize(string text, IReadOnlyList<int>? pages = null)
    {
        var normalized = new NormalizedText();
        var content = new StringBuilder();
        var enumerator = StringInfo.GetTextElementEnumerator(text);
        while (enumerator.MoveNext())
        {
            var offset = enumerator.ElementIndex;
            foreach (var character in enumerator.GetTextElement().Normalize(NormalizationForm.FormKC))
            {
                if (char.IsWhiteSpace(character) || character is '\u00ad' or '\u200b' or '\ufeff') continue;
                content.Append(character);
                normalized.Offsets.Add(offset);
                normalized.Pages.Add(pages is null ? 0 : pages[offset]);
            }
        }
        normalized.Text = content.ToString();
        return normalized;
    }

    private sealed class NormalizedText
    {
        internal string Text { get; set; } = string.Empty;
        internal List<int> Offsets { get; } = new();
        internal List<int> Pages { get; } = new();
    }
    private sealed class PartMapping
    {
        internal PartMapping(DocxSourcePart part, IReadOnlyList<int> pages, int fallback) { Part = part; Pages = pages; FallbackPage = fallback; }
        internal DocxSourcePart Part { get; }
        internal IReadOnlyList<int> Pages { get; }
        internal int FallbackPage { get; }
    }
}
