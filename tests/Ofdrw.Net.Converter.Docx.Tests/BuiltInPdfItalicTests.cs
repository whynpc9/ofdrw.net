using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Converter.Pdf;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;
using PdfSharpCore.Pdf;
using Ofdrw.Net.Converter.Docx.Converters;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed partial class DocxConversionTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BuiltInPdf_RegularFaceItalicKeepsAnchorsAndAdjacentRegularUpright(bool wrap)
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-italic-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.Copy(Path.GetFullPath(Path.Combine(Path.GetDirectoryName(ResolveGeneratedSample())!,
                "../../Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf")), Path.Combine(directory, "fixture.ttf"));
            async Task<byte[]> Convert(bool italic)
            {
                var content = $"""
                    <w:p><w:r><w:rPr><w:rFonts w:ascii="Ofdrw Test Face"/><w:sz w:val="32"/></w:rPr><w:t>ABA</w:t></w:r>
                    <w:r><w:rPr><w:rFonts w:ascii="Ofdrw Test Face"/><w:sz w:val="32"/>{(italic ? "<w:i/>" : "")}</w:rPr><w:t>ABA</w:t></w:r>
                    <w:r><w:rPr><w:rFonts w:ascii="Ofdrw Test Face"/><w:sz w:val="32"/></w:rPr><w:t>ABA</w:t></w:r></w:p>
                    """;
                if (wrap)
                    content = content.Replace("</w:p>", "</w:p><w:sectPr><w:pgSz w:w=\"1500\" w:h=\"8000\"/><w:pgMar w:top=\"400\" w:bottom=\"400\" w:left=\"400\" w:right=\"400\"/></w:sectPr>");
                using var input = CreateMinimalDocx(content);
                var options = new DocxConversionOptions { Engine = DocxConversionEngine.BuiltIn };
                options.FontDirectories.Add(directory);
                using var output = new MemoryStream();
                await new DocxToPdfConverter(options).ConvertAsync(input, output);
                return output.ToArray();
            }
            using var control = PdfPigDocument.Open(await Convert(false));
            using var styled = PdfPigDocument.Open(await Convert(true));
            var original = control.GetPage(1).Letters;
            var letters = styled.GetPage(1).Letters;
            Assert.Equal("ABAABAABA", styled.GetPage(1).Text);
            Assert.Equal(9, letters.Count);
            if (wrap) Assert.True(letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count() > 1);
            for (var i = 0; i < letters.Count; i++)
            {
                Assert.Equal(original[i].StartBaseLine.X, letters[i].StartBaseLine.X, 3);
                Assert.Equal(original[i].StartBaseLine.Y, letters[i].StartBaseLine.Y, 3);
                var skew = letters[i].BoundingBox.TopLeft.X - letters[i].BoundingBox.BottomLeft.X;
                if (i is >= 3 and < 6) Assert.True(skew > 0.5, $"Glyph {i} must be visibly italic; skew={skew}");
                else Assert.Equal(0, skew, 3);
            }
        }
        finally { Directory.Delete(directory, true); }
    }
}

public sealed class BuiltInPdfItalicRendererTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactDrawingCallKeepsSelectedFontAndSkewsOnlyMissingItalic(bool physicalItalic)
    {
        PdfFontRegistry.EnsureInstalled();
        var root = FindRoot();
        var path = physicalItalic ? "e2e/Ofdrw.Net.SkiaSharp.E2E/testdata/fonts/italic.ttf"
            : "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf";
        var bytes = File.ReadAllBytes(Path.Combine(root, path));
        var host = new FalseFlagHost(bytes);
        var faces = new BuiltInPdfItalicFaces(host, host.GetFont);
        var family = faces.Resolve("host-style-probe", false);
        Assert.Equal(family, faces.Resolve("host-style-probe", false));
        Assert.Equal(1, host.Requests);
        var face = GlobalFontSettings.FontResolver.ResolveTypeface(family, false, true);
        Assert.Equal(!physicalItalic, face.MustSimulateItalic);
        Assert.Equal(PdfFontRegistry.RegisterFontFace(bytes, italic: true), family);
        var original = new Ofdrw.Net.Core.Fonts.OpenTypeFace(bytes);
        var embedded = new Ofdrw.Net.Core.Fonts.OpenTypeFace(GlobalFontSettings.FontResolver.GetFont(face.FaceName));
        foreach (var table in new[] { "glyf", "hmtx", "cmap", "OS/2" })
            Assert.Equal(original.Tables[table], embedded.Tables[table]);
        var font = new XFont(family, 20, XFontStyle.Italic);
        using var pdf = new PdfDocument();
        var page = pdf.AddPage();
        using (var target = XGraphics.FromPdfPage(page))
        using (var graphics = XGraphics.FromRenderer(new BuiltInPdfItalicRenderer(target, faces.ActualItalic),
                   new XSize(page.Width, page.Height), XGraphicsUnit.Point, XPageDirection.Downwards))
        {
            graphics.DrawString("ABA", font, XBrushes.Black, 40, 60);
            // The same physical face is deliberately drawn without the wrapper.
            // True italic must be byte-for-byte the same glyph geometry here.
            target.DrawString("ABA", font, XBrushes.Black, 40, 100);
            graphics.DrawString("ABA", new XFont(PdfFontRegistry.RegisterFontFace(bytes), 20), XBrushes.Black, 40, 140);
        }
        using var output = new MemoryStream(); pdf.Save(output, false);
        using var parsed = PdfPigDocument.Open(output.ToArray());
        var letters = parsed.GetPage(1).Letters;
        Assert.Equal("ABAABAABA", parsed.GetPage(1).Text);
        Assert.Equal(9, letters.Count);
        for (var i = 0; i < 3; i++)
        {
            Assert.Equal(letters[i + 3].StartBaseLine.X, letters[i].StartBaseLine.X, 3);
            Assert.Equal(40, letters[i].StartBaseLine.Y - letters[i + 3].StartBaseLine.Y, 3);
            var skew = letters[i].BoundingBox.TopLeft.X - letters[i].BoundingBox.BottomLeft.X;
            var controlSkew = letters[i + 3].BoundingBox.TopLeft.X - letters[i + 3].BoundingBox.BottomLeft.X;
            if (physicalItalic) Assert.Equal(controlSkew, skew, 4);
            else Assert.True(skew > controlSkew + 1);
        }
    }

    [Fact]
    public void UnknownItalicFaceOrUnsupportedAnchorIsRefusedBeforeDrawing()
    {
        PdfFontRegistry.EnsureInstalled();
        var bytes = File.ReadAllBytes(Path.Combine(FindRoot(), "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"));
        var alias = PdfFontRegistry.RegisterFontFace(bytes, italic: true);
        var font = new XFont(alias, 12, XFontStyle.Italic);
        using var pdf = new PdfDocument(); var page = pdf.AddPage();
        using var target = XGraphics.FromPdfPage(page);
        var renderer = new BuiltInPdfItalicRenderer(target, new Dictionary<string, bool>());
        Assert.Throws<InvalidDataException>(() => renderer.DrawString("A", font, XBrushes.Black, new XRect(10, 20, 0, 0), XStringFormats.BaseLineLeft));
        renderer = new BuiltInPdfItalicRenderer(target, new Dictionary<string, bool> { [font.Name] = false });
        Assert.Throws<InvalidDataException>(() => renderer.DrawString("A", font, XBrushes.Black, new XRect(10, 20, 80, 30), XStringFormats.TopLeft));
    }

    private static string FindRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
            if (File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln"))) return directory.FullName;
        throw new DirectoryNotFoundException("Repository fixture root not found.");
    }

    private sealed class FalseFlagHost(byte[] bytes) : IFontResolver
    {
        public int Requests { get; private set; }
        public string DefaultFontName => "host-style-probe";
        public FontResolverInfo ResolveTypeface(string family, bool bold, bool italic)
        {
            Requests++;
            return new FontResolverInfo("host-selected-face", false, false);
        }
        public byte[] GetFont(string face) => bytes;
    }
}
