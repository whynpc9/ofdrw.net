using System.Globalization;
using System.IO.Compression;
using System.Text;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using PdfSharpCore.Fonts;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class DocxFontIdentityTests
{
    [RealTrustedFontsTheory]
    [InlineData(false, OfdFontEmbeddingMode.SubsetWhenSafe)]
    [InlineData(true, OfdFontEmbeddingMode.SubsetWhenSafe)]
    [InlineData(true, OfdFontEmbeddingMode.Full)]
    public async Task NativeNotoUsesOriginalLatinMetricsAndEmbedsTheSameFaceThroughPdf(bool explicitNative, OfdFontEmbeddingMode embedding)
    {
        var original = TrustedFontFixtures.Read("cjk");
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-noto-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.WriteAllBytes(Path.Combine(directory, "noto.ttf"), original);
            var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false };
            options.FontDirectories.Add(directory);
            options.FontEmbedding.Mode = embedding;
            if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
            using var input = Source(); using var ofd = new MemoryStream();
            await new DocxToOfdConverter(options).ConvertAsync(input, ofd);
            ofd.Position = 0;
            var package = await new OfdReader().ReadAsync(ofd);
            var texts = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().ToArray();
            Assert.Equal("A iW中文A iW中文A iW中文", string.Concat(texts.Select(text => text.Text)));
            Assert.Equal(3, package.Fonts.Count);
            var context = new DocumentFontContext(package.Fonts.ToArray(), new RejectHostFonts());
            foreach (var font in package.Fonts)
            {
                Assert.NotEmpty(font.Data);
                Assert.Equal("Noto Sans CJK SC", font.FamilyName);
                if (embedding == OfdFontEmbeddingMode.Full) Assert.Equal(original, font.Data);
                // Subsetting preserves original glyph IDs and original hmtx advances.
                AssertNotoMetrics(font.Data, original);
            }
            foreach (var text in texts)
            {
                var deltas = text.Runs.Single().DeltaX!.Split(' ').Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
                // Independent original-font oracle: A=608, space=224 units, UPEM=1000.
                // These assertions fail if measurement silently binds host Arial.
                Assert.InRange(deltas[0], text.FontSizeMillimeters * .608 - .001, text.FontSizeMillimeters * .608 + .001);
                Assert.InRange(deltas[1], text.FontSizeMillimeters * .224 - .001, text.FontSizeMillimeters * .224 + .001);
                Assert.False(string.IsNullOrEmpty(context.Resolve(text, out var resource)));
                Assert.NotNull(resource);
            }
            ofd.Position = 0; using var pdf = new MemoryStream();
            await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using var semantic = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
            Assert.Equal("A iW中文A iW中文A iW中文", semantic.GetPage(1).Text);
            using var parsed = PdfReader.Open(new MemoryStream(pdf.ToArray()), PdfDocumentOpenMode.Import);
            var descriptors = parsed.Internals.GetAllObjects().OfType<PdfDictionary>()
                .Where(value => value.Elements.ContainsKey("/FontFile2")).ToArray();
            Assert.NotEmpty(descriptors);
            foreach (var descriptor in descriptors)
                AssertNotoMetrics(descriptor.Elements.GetDictionary("/FontFile2").Stream.UnfilteredValue, original);
        }
        finally { Directory.Delete(directory, true); }
    }

    private static void AssertNotoMetrics(byte[] bytes, byte[] original)
    {
        var face = new OpenTypeFace(bytes);
        Assert.Equal(1000, OpenTypeFace.U16(face.Table("head"), 18));
        Assert.Equal(608, OpenTypeFace.U16(face.Table("hmtx"), 34 * 4)); // original A gid
        Assert.Equal(224, OpenTypeFace.U16(face.Table("hmtx"), 1 * 4)); // original space gid
        var source = new OpenTypeFace(original);
        var cmap = new OpenTypeCmap(source);
        foreach (var scalar in "AiW中文")
        {
            var glyph = cmap.Glyph(scalar);
            var expected = GlyphBytes(source, glyph);
            var actual = GlyphBytes(face, glyph);
            // The subset writer aligns glyph records to four bytes. Only
            // zero alignment padding may differ; outline/instruction bytes may not.
            Assert.InRange(actual.Length - expected.Length, 0, 3);
            Assert.Equal(expected, actual.Take(expected.Length).ToArray());
            Assert.All(actual.Skip(expected.Length), value => Assert.Equal(0, value));
        }
    }

    private static byte[] GlyphBytes(OpenTypeFace face, int glyph)
    {
        var loca = face.Table("loca");
        var shortOffsets = OpenTypeFace.U16(face.Table("head"), 50) == 0;
        int Offset(int index) => shortOffsets ? OpenTypeFace.U16(loca, index * 2) * 2 :
            checked((int)OpenTypeFace.U32(loca, index * 4));
        var start = Offset(glyph); var end = Offset(glyph + 1);
        return face.Table("glyf")[start..end];
    }

    private sealed class RejectHostFonts : IFontResolver
    {
        public string DefaultFontName => throw new InvalidOperationException("Embedded Noto must not consult host defaults.");
        public byte[] GetFont(string faceName) => throw new InvalidOperationException("Embedded Noto must not read host fonts.");
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic) =>
            throw new InvalidOperationException("Embedded Noto must not resolve host fonts.");
    }

    private static MemoryStream Source()
    {
        var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            void Add(string name, string xml)
            { using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)); writer.Write(xml); }
            Add("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            Add("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            var runs = string.Concat(new[] { "", "<w:b/>", "<w:i/>" }.Select(style => $"<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Noto Sans CJK SC\" w:hAnsi=\"Noto Sans CJK SC\" w:eastAsia=\"Noto Sans CJK SC\"/>{style}</w:rPr><w:t>A iW中文</w:t></w:r></w:p>"));
            Add("word/document.xml", $"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{runs}</w:body></w:document>");
        }
        stream.Position = 0; return stream;
    }
}
