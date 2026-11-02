using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.TestSupport;

// Shared deterministic inputs for existing regressions and current release evidence.
internal static class DocxReleaseFixtures
{
    internal static MemoryStream CreateMinimalDocx(
        string bodyContent,
        string? documentRelationships = null,
        IReadOnlyDictionary<string, byte[]>? binaryParts = null)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZipEntry(archive, "[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
                  <Default Extension="xml" ContentType="application/xml" />
                  <Default Extension="png" ContentType="image/png" />
                  <Override PartName="/word/document.xml"
                    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
                </Types>
                """);
            WriteZipEntry(archive, "_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1"
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    Target="word/document.xml" />
                </Relationships>
                """);
            WriteZipEntry(archive, "word/document.xml", $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>{{bodyContent}}<w:sectPr /></w:body>
                </w:document>
                """);
            if (documentRelationships is not null)
            {
                WriteZipEntry(
                    archive,
                    "word/_rels/document.xml.rels",
                    documentRelationships);
            }

            if (binaryParts is not null)
            {
                foreach (var part in binaryParts)
                {
                    var entry = archive.CreateEntry(part.Key);
                    using var entryStream = entry.Open();
                    entryStream.Write(part.Value);
                }
            }
        }

        stream.Position = 0;
        return stream;
    }

    private static void WriteZipEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }

    internal static MemoryStream CreateMultiSectionDocx()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<DocumentSettingsPart>().Settings = new W.Settings(new W.EvenAndOddHeaders());
            var odd = main.AddNewPart<HeaderPart>(); odd.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text("Odd Header"))));
            var even = main.AddNewPart<HeaderPart>(); even.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text("Even Header"))));
            var first = main.AddNewPart<HeaderPart>(); first.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text("First Header"))));
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new W.Footer(new W.Paragraph(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }),
                new W.Run(new W.Text("Page ")), new W.SimpleField(new W.Run(new W.Text("0"))) { Instruction = "PAGE" },
                new W.Run(new W.Text("/")), new W.SimpleField(new W.Run(new W.Text("0"))) { Instruction = "NUMPAGES" }));
            var firstSection = new W.SectionProperties(
                new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(odd) },
                new W.HeaderReference { Type = W.HeaderFooterValues.Even, Id = main.GetIdOfPart(even) },
                new W.HeaderReference { Type = W.HeaderFooterValues.First, Id = main.GetIdOfPart(first) },
                new W.FooterReference { Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) },
                new W.FooterReference { Type = W.HeaderFooterValues.Even, Id = main.GetIdOfPart(footer) },
                new W.FooterReference { Type = W.HeaderFooterValues.First, Id = main.GetIdOfPart(footer) },
                new W.PageNumberType { Start = 1 }, new W.TitlePage());
            var body = new W.Body();
            foreach (var (label, index) in new[] { ("one", 0), ("two", 1), ("three", 2) })
            {
                var properties = new W.ParagraphProperties();
                if (index > 0) properties.Append(new W.PageBreakBefore());
                if (index == 2) properties.Append(firstSection);
                body.Append(new W.Paragraph(properties, new W.Run(new W.Text("Body " + label))));
            }
            body.Append(new W.Paragraph(new W.Run(new W.Text("Body four"))),
                new W.Paragraph(new W.ParagraphProperties(new W.PageBreakBefore()), new W.Run(new W.Text("Body five"))),
                new W.SectionProperties(new W.PageNumberType { Start = 10 }));
            main.Document = new W.Document(body);
        }
        stream.Position = 0;
        return stream;
    }

    internal static (MemoryStream Docx, byte[] Jpeg) CreateInlineJpegDocx(string label = "Native JPEG image fixture")
    {
        var body = """
            <w:p><w:r><w:t>{LABEL}</w:t></w:r></w:p>
            <w:p><w:r><w:drawing><wp:inline><wp:extent cx="1440000" cy="720000"/><wp:docPr id="1" name="image"/>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image.jpg"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1440000" cy="720000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
            """.Replace("{LABEL}", System.Security.SecurityElement.Escape(label), StringComparison.Ordinal);
        using var pixels = new Image<Rgba32>(80, 40);
        for (var y = 0; y < 40; y++)
        for (var x = 0; x < 80; x++) pixels[x, y] = x < 40 ? new Rgba32(255, 0, 0) : new Rgba32(0, 0, 255);
        using var jpeg = new MemoryStream(); pixels.SaveAsJpeg(jpeg); jpeg.Position = 0;
        var input = CreateMinimalDocx(body);
        using (var document = WordprocessingDocument.Open(input, true))
            document.MainDocumentPart!.AddImagePart(ImagePartType.Jpeg, "rId2").FeedData(jpeg);
        input.Position = 0;
        return (input, jpeg.ToArray());
    }
}
