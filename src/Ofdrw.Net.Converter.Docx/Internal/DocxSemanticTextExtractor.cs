using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Internal;

internal sealed class DocxSourceDocument
{
    internal List<DocxSourceSection> Sections { get; } = new();
    internal List<DocxSourcePart> Supplemental { get; } = new();
}

internal sealed class DocxSourceSection
{
    internal List<DocxSourcePart> Paragraphs { get; } = new();
    internal Dictionary<string, DocxSourcePart> Headers { get; } = new();
    internal Dictionary<string, DocxSourcePart> Footers { get; } = new();
    internal bool DifferentFirstPage { get; set; }
    internal bool DifferentEvenPages { get; set; }
    internal int? PageNumberStart { get; set; }
    internal double TopMargin { get; set; } = 72;
    internal double BottomMargin { get; set; } = 72;
}

internal sealed class DocxSourcePart
{
    internal string Kind { get; set; } = "body";
    internal string Id { get; set; } = string.Empty;
    internal int SectionIndex { get; set; }
    internal bool PageBreakBefore { get; set; }
    internal List<DocxSourceSegment> Segments { get; } = new();
    internal List<(string Kind, string Id, int Offset)> References { get; } = new();
    internal string LiteralText => string.Concat(Segments.Where(segment => segment.Field is null).Select(segment => segment.Text));
    internal string Evaluate(int page, int total, int sectionPages) => string.Concat(Segments.Select(segment =>
        segment.Field is null ? segment.Text : FieldValue(segment.Field.Value, page, total, sectionPages)));
    internal static string FieldValue(BuiltInPageFieldKind kind, int page, int total, int sectionPages) =>
        (kind == BuiltInPageFieldKind.Page ? page : kind == BuiltInPageFieldKind.TotalPages ? total : sectionPages).ToString(CultureInfo.InvariantCulture);
}

internal sealed class DocxSourceSegment
{
    internal string Text { get; set; } = string.Empty;
    internal BuiltInPageFieldKind? Field { get; set; }
}

/// <summary>Reads original Word text and field/reference structure, without inferring rendered pages.</summary>
internal static class DocxSemanticTextExtractor
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static DocxSourceDocument ExtractDocument(string path, DocxConversionOptions options, CancellationToken cancellationToken)
    {
        DocxPackageValidator.Validate(path, options);
        using var document = WordprocessingDocument.Open(path, false, new OpenSettings { AutoSave = false });
        var main = document.MainDocumentPart ?? throw new System.IO.InvalidDataException("The DOCX has no main document part.");
        var body = main.Document?.Body ?? throw new System.IO.InvalidDataException("The DOCX has no body.");
        var source = new DocxSourceDocument();
        var section = new DocxSourceSection();
        source.Sections.Add(section);
        var count = 0;
        void Check(OpenXmlElement root)
        {
            foreach (var _ in root.Descendants())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (++count > options.MaxDocumentElements) throw new System.IO.InvalidDataException("DOCX source text exceeds the configured XML element limit.");
            }
        }
        Check(body);
        foreach (var element in body.ChildElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (element is W.SectionProperties properties) ApplySection(properties, section);
            else
            {
                var paragraphs = element is W.Paragraph paragraph ? new[] { paragraph } : element.Descendants<W.Paragraph>();
                foreach (var item in paragraphs)
                {
                    var part = ReadParagraph(item);
                    part.SectionIndex = source.Sections.Count - 1;
                    section.Paragraphs.Add(part);
                    if (item.ParagraphProperties?.SectionProperties is W.SectionProperties paragraphSection)
                    {
                        ApplySection(paragraphSection, section);
                        var next = new DocxSourceSection();
                        foreach (var pair in section.Headers) next.Headers[pair.Key] = pair.Value;
                        foreach (var pair in section.Footers) next.Footers[pair.Key] = pair.Value;
                        section = next;
                        source.Sections.Add(section);
                    }
                }
            }
        }
        if (source.Sections.Count > 1 && source.Sections.Last().Paragraphs.Count == 0) source.Sections.RemoveAt(source.Sections.Count - 1);
        AddSupplemental("footnote", main.FootnotesPart?.Footnotes);
        AddSupplemental("endnote", main.EndnotesPart?.Endnotes);
        AddSupplemental("comment", main.WordprocessingCommentsPart?.Comments);
        return source;

        void ApplySection(W.SectionProperties properties, DocxSourceSection target)
        {
            var title = properties.GetFirstChild<W.TitlePage>();
            target.DifferentFirstPage = title is not null && (title.Val?.Value ?? true);
            var even = main.DocumentSettingsPart?.Settings?.GetFirstChild<W.EvenAndOddHeaders>();
            target.DifferentEvenPages = even is not null && (even.Val?.Value ?? true);
            if (int.TryParse(properties.GetFirstChild<W.PageNumberType>()?.Start?.Value.ToString(), out var start)) target.PageNumberStart = start;
            var margins = properties.GetFirstChild<W.PageMargin>();
            if (margins?.Top is not null) target.TopMargin = margins.Top.Value / 20d;
            if (margins?.Bottom is not null) target.BottomMargin = margins.Bottom.Value / 20d;
            foreach (var reference in properties.Elements<W.HeaderReference>())
            {
                if (reference.Id?.Value is not string id || main.GetPartById(id) is not HeaderPart part || part.Header is null) continue;
                Check(part.Header);
                target.Headers[reference.Type?.Value == W.HeaderFooterValues.First ? "first" : reference.Type?.Value == W.HeaderFooterValues.Even ? "even" : "default"] = ReadPart(part.Header, "header", id);
            }
            foreach (var reference in properties.Elements<W.FooterReference>())
            {
                if (reference.Id?.Value is not string id || main.GetPartById(id) is not FooterPart part || part.Footer is null) continue;
                Check(part.Footer);
                target.Footers[reference.Type?.Value == W.HeaderFooterValues.First ? "first" : reference.Type?.Value == W.HeaderFooterValues.Even ? "even" : "default"] = ReadPart(part.Footer, "footer", id);
            }
        }

        void AddSupplemental(string kind, OpenXmlElement? root)
        {
            if (root is null) return;
            Check(root);
            foreach (var node in root.ChildElements)
            {
                var id = Attribute(node, "id");
                var type = Attribute(node, "type");
                if (string.IsNullOrEmpty(id) || id!.StartsWith("-", StringComparison.Ordinal) || type is "separator" or "continuationSeparator") continue;
                source.Supplemental.Add(ReadPart(node, kind, id));
            }
        }
    }

    private static DocxSourcePart ReadPart(OpenXmlElement root, string kind, string id)
    {
        var result = new DocxSourcePart { Kind = kind, Id = id };
        foreach (var paragraph in root.Descendants<W.Paragraph>())
        {
            result.Segments.AddRange(ReadParagraph(paragraph).Segments);
            result.Segments.Add(new DocxSourceSegment { Text = "\n" });
        }
        return result;
    }

    private static DocxSourcePart ReadParagraph(W.Paragraph paragraph)
    {
        var part = new DocxSourcePart();
        var pageBreak = paragraph.ParagraphProperties?.PageBreakBefore;
        part.PageBreakBefore = pageBreak is not null && (pageBreak.Val?.Value ?? true);
        var fields = new Stack<(StringBuilder Instruction, bool Suppress)>();
        var text = new StringBuilder();
        var literalOffset = 0;
        void Flush()
        {
            if (text.Length == 0) return;
            part.Segments.Add(new DocxSourceSegment { Text = text.ToString() });
            literalOffset += text.Length;
            text.Clear();
        }
        foreach (var node in paragraph.Descendants().Where(node => ReferenceEquals(node.Ancestors<W.Paragraph>().FirstOrDefault(), paragraph)))
        {
            if (node.Ancestors<W.SimpleField>().Any(field => FieldKind(field.Instruction?.Value) is not null)) continue;
            if (node is W.SimpleField simple && FieldKind(simple.Instruction?.Value) is BuiltInPageFieldKind simpleKind)
            {
                Flush(); part.Segments.Add(new DocxSourceSegment { Field = simpleKind }); continue;
            }
            if (node is W.FieldChar marker)
            {
                if (marker.FieldCharType?.Value == W.FieldCharValues.Begin) fields.Push((new StringBuilder(), false));
                else if (marker.FieldCharType?.Value == W.FieldCharValues.Separate && fields.Count > 0)
                {
                    var field = fields.Pop();
                    var kind = FieldKind(field.Instruction.ToString());
                    if (kind is not null) { Flush(); part.Segments.Add(new DocxSourceSegment { Field = kind }); }
                    fields.Push((field.Instruction, kind is not null));
                }
                else if (marker.FieldCharType?.Value == W.FieldCharValues.End && fields.Count > 0) fields.Pop();
                continue;
            }
            if (node is W.FieldCode code) { if (fields.Count > 0) fields.Peek().Instruction.Append(code.Text); continue; }
            if (fields.Any(field => field.Suppress)) continue;
            switch (node)
            {
                case W.Text value: text.Append(value.Text); break;
                case W.TabChar: text.Append('\t'); break;
                case W.Break: text.Append('\n'); break;
                case W.CarriageReturn: text.Append('\n'); break;
                case W.FootnoteReference: part.References.Add(("footnote", Attribute(node, "id"), literalOffset + text.Length)); break;
                case W.EndnoteReference: part.References.Add(("endnote", Attribute(node, "id"), literalOffset + text.Length)); break;
                case W.CommentReference: part.References.Add(("comment", Attribute(node, "id"), literalOffset + text.Length)); break;
            }
        }
        Flush();
        return part;
    }

    private static BuiltInPageFieldKind? FieldKind(string? instruction) =>
        instruction?.Trim().Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.ToUpperInvariant() switch
        {
            "PAGE" => BuiltInPageFieldKind.Page, "NUMPAGES" => BuiltInPageFieldKind.TotalPages,
            "SECTIONPAGES" => BuiltInPageFieldKind.SectionPages, _ => null
        };
    private static string Attribute(OpenXmlElement node, string name) =>
        node.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name && attribute.NamespaceUri == WordNamespace).Value ?? string.Empty;
}
