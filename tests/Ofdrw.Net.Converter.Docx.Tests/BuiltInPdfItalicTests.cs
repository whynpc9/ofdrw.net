using SixLabors.ImageSharp;
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

    [Fact]
    public void DifferentPageSizesKeepHeadersFieldsTablesAndImagesThroughDrawingSeam()
    {
        var temporary = Path.Combine(Path.GetTempPath(), "ofdrw-italic-pages-" + Guid.NewGuid().ToString("N"));
        var directory = Environment.GetEnvironmentVariable("OFDRW_003_ARTIFACTS") ?? temporary;
        Directory.CreateDirectory(directory);
        BuiltInParagraphModel Text(string text, bool italic = false) => new()
        {
            Inlines = { new BuiltInTextModel { Text = text, Format = { FontFamily = "Arial", FontSizePoints = 12, Italic = italic } } }
        };
        try
        {
            using var image = new SixLabors.ImageSharp.Image<SixLabors.ImageSharp.PixelFormats.Rgba32>(12, 8,
                new SixLabors.ImageSharp.PixelFormats.Rgba32(50, 120, 210));
            using var encoded = new MemoryStream();
            image.Save(encoded, new SixLabors.ImageSharp.Formats.Png.PngEncoder());
            string Render(bool italic)
            {
                var model = new BuiltInDocumentModel();
                foreach (var (width, height, label) in new[] { (300d, 420d, "Portrait"), (320d, 440d, "SecondPortrait") })
                {
                    var section = new BuiltInSectionModel
                    {
                        PageWidthPoints = width, PageHeightPoints = height,
                        MarginTopPoints = 50, MarginBottomPoints = 50, MarginLeftPoints = 35, MarginRightPoints = 35,
                        HeaderDistancePoints = 20, FooterDistancePoints = 20
                    };
                    section.Headers.Add(Text(label + " header"));
                    var footer = Text("Page ");
                    footer.Inlines.Add(new BuiltInPageNumberModel());
                    footer.Inlines.Add(new BuiltInTextModel { Text = " of ", Format = { FontFamily = "Arial" } });
                    footer.Inlines.Add(new BuiltInPageNumberModel { Kind = BuiltInPageFieldKind.TotalPages });
                    section.Footers.Add(footer);
                    section.Blocks.Add(Text(label + " italic", italic));
                    section.Blocks.Add(Text(label + " regular"));
                    var table = new BuiltInTableModel { HasBorders = true };
                    table.ColumnWidthsPoints.Add(100); table.ColumnWidthsPoints.Add(100);
                    var row = new BuiltInTableRowModel();
                    row.Cells.Add(new BuiltInTableCellModel { ShadingHex = "D9EAF7", Paragraphs = { Text("Cell A") } });
                    row.Cells.Add(new BuiltInTableCellModel { Paragraphs = { Text("Cell B") } });
                    table.Rows.Add(row); section.Blocks.Add(table);
                    section.Blocks.Add(new BuiltInParagraphModel { Inlines = { new BuiltInImageModel
                    {
                        Name = "blue.png", Data = encoded.ToArray(), WidthPoints = 48, HeightPoints = 32
                    } } });
                    model.Sections.Add(section);
                }
                var path = Path.Combine(directory, italic ? "mixed-sections.pdf" : "mixed-sections-control.pdf");
                new BuiltInPdfRenderer(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default).Render(model, path);
                return path;
            }
            using var control = PdfPigDocument.Open(Render(false));
            using var styled = PdfPigDocument.Open(Render(true));
            Assert.Equal(2, styled.NumberOfPages);
            for (var number = 1; number <= 2; number++)
            {
                var expected = control.GetPage(number); var actual = styled.GetPage(number);
                Assert.Equal(number == 1 ? 300 : 320, actual.Width, 3);
                Assert.Equal(number == 1 ? 420 : 440, actual.Height, 3);
                Assert.Equal(expected.Rotation.Value, actual.Rotation.Value);
                Assert.Equal(expected.Text, actual.Text);
                Assert.Contains(number == 1 ? "Portraitheader" : "SecondPortraitheader", actual.Text);
                Assert.Contains($"Page{number}of2", actual.Text);
                Assert.Contains("CellACellB", actual.Text);
                Assert.Equal(1, actual.NumberOfImages);
                Assert.Equal(expected.GetImages().Single().BoundingBox, actual.GetImages().Single().BoundingBox);
                Assert.True(actual.Paths.Count >= 6); // Fill plus cell-border segments.
                Assert.Equal(expected.Paths.Count, actual.Paths.Count);
                for (var i = 0; i < actual.Letters.Count; i++)
                {
                    Assert.Equal(expected.Letters[i].StartBaseLine.X, actual.Letters[i].StartBaseLine.X, 3);
                    Assert.Equal(expected.Letters[i].StartBaseLine.Y, actual.Letters[i].StartBaseLine.Y, 3);
                }
            }
        }
        finally { if (directory == temporary) Directory.Delete(temporary, true); }
    }

    [Fact]
    public void PreparedPortraitAndLandscapePagesMatchOriginalMigraDocAnchors()
    {
        PdfFontRegistry.EnsureInstalled();
        var bytes = File.ReadAllBytes(Path.Combine(FindRoot(), "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"));
        var family = PdfFontRegistry.RegisterFontFace(bytes);
        var faces = new BuiltInPdfItalicFaces(GlobalFontSettings.FontResolver);
        var italicFamily = faces.Resolve(PdfFontRegistry.RegisterFontFace(bytes, italic: true), false);
        var directory = Environment.GetEnvironmentVariable("OFDRW_003_ARTIFACTS") ??
            Path.Combine(Path.GetTempPath(), "ofdrw-actual-orientation-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string Render(bool italic)
            {
                var document = new MigraDocCore.DocumentObjectModel.Document();
                foreach (var landscape in new[] { false, true })
                {
                    var section = document.AddSection();
                    section.PageSetup.PageWidth = 300;
                    section.PageSetup.PageHeight = 420;
                    section.PageSetup.Orientation = landscape ? MigraDocCore.DocumentObjectModel.Orientation.Landscape
                        : MigraDocCore.DocumentObjectModel.Orientation.Portrait;
                    section.PageSetup.LeftMargin = section.PageSetup.RightMargin = 35;
                    section.PageSetup.TopMargin = section.PageSetup.BottomMargin = 50;
                    var line = section.AddParagraph();
                    var first = line.AddFormattedText("ABA"); first.Font.Name = family; first.Font.Size = 16;
                    var middle = line.AddFormattedText("ABA"); middle.Font.Name = italic ? italicFamily : family;
                    middle.Font.Italic = italic; middle.Font.Size = 16;
                    var last = line.AddFormattedText("ABA"); last.Font.Name = family; last.Font.Size = 16;
                }
                var renderer = new MigraDocCore.Rendering.PdfDocumentRenderer(true) { Document = document };
                if (italic) BuiltInPdfRenderer.RenderPreparedPages(renderer, faces.ActualItalic, 10, default);
                else renderer.RenderDocument();
                var path = Path.Combine(directory, italic ? "mixed-orientations.pdf" : "mixed-orientations-control.pdf");
                renderer.PdfDocument.Save(path); return path;
            }
            using var control = PdfPigDocument.Open(Render(false)); using var styled = PdfPigDocument.Open(Render(true));
            Assert.Equal(2, styled.NumberOfPages);
            for (var number = 1; number <= 2; number++)
            {
                var actual = styled.GetPage(number); var expected = control.GetPage(number);
                Assert.Equal(number == 1 ? 300 : 420, actual.Width, 3);
                Assert.Equal(number == 1 ? 420 : 300, actual.Height, 3);
                Assert.Equal(expected.Width, actual.Width, 3); Assert.Equal(expected.Height, actual.Height, 3);
                Assert.Equal(expected.Rotation.Value, actual.Rotation.Value);
                Assert.Equal("ABAABAABA", actual.Text);
                for (var i = 0; i < actual.Letters.Count; i++)
                {
                    Assert.Equal(expected.Letters[i].StartBaseLine.X, actual.Letters[i].StartBaseLine.X, 3);
                    Assert.Equal(expected.Letters[i].StartBaseLine.Y, actual.Letters[i].StartBaseLine.Y, 3);
                    var skew = actual.Letters[i].BoundingBox.TopLeft.X - actual.Letters[i].BoundingBox.BottomLeft.X;
                    if (i is >= 3 and < 6) Assert.True(skew > 0.5);
                    else Assert.Equal(0, skew, 3);
                }
            }
        }
        finally { if (Environment.GetEnvironmentVariable("OFDRW_003_ARTIFACTS") is null) Directory.Delete(directory, true); }
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
