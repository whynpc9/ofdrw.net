using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Text.HarfBuzz;
using UglyToad.PdfPig;

namespace Ofdrw.Net.Converter.Pdf.Tests;

// F06 (Issue #109): bounded visual-line RTL import. Positive fixtures come from
// three independent generators (fpdf2+uharfbuzz, Pango/Cairo, LibreOffice) and
// carry a declared logical source; synthetic PDFs isolate each rejection guard.
public sealed class PdfRtlVisualLineImportTests
{
    public PdfRtlVisualLineImportTests() => PdfFontRegistry.EnsureInstalled();

    private static PdfToOfdOptions Profile(PdfTextBaseDirection direction = PdfTextBaseDirection.Auto) => new()
    {
        RightToLeftTextImport = PdfRightToLeftTextImport.VisualLineProfile,
        RightToLeftBaseDirection = direction,
    };

    [Theory]
    // fpdf2 shaped cell: content in visual order. Default GetWords reverses both words and their order.
    [InlineData("D3/fpdf2-two-word-cell.pdf", "Auto", "abc שלום מקום def ")]
    // Pango/Cairo: the same phrase with glyphs in logical content order; geometry alone decides.
    [InlineData("F06/pango-two-word.pdf", "Auto", "abc שלום מקום def ")]
    // Lam-alef: one glyph maps to two logical characters, in both base directions.
    [InlineData("F06/fpdf2-arabic-ltr-cell.pdf", "Auto", "abc مرحبا لا def ")]
    [InlineData("F06/fpdf2-arabic-rtl-cell.pdf", "Auto", "مرحبا بالعالم لا ")]
    // Wrapped paragraphs whose lines are ambiguous on their own need the caller's base direction.
    [InlineData("F06/fpdf2-hebrew-rtl-wrap.pdf", "RightToLeft", "שלום עולם abc מקום טוב ")]
    [InlineData("F06/pango-after-rtl-wrap.pdf", "LeftToRight", "abc שלום מקום def ")]
    [InlineData("F06/fpdf2-between-rtl-zero-margin.pdf", "LeftToRight", "abc שלום מקום def ")]
    // LibreOffice: RTL Hebrew, RTL Arabic and an LTR paragraph, each resolved by its own line.
    [InlineData("F06/libreoffice-rtl.pdf", "Auto", "שלום עולם abc מקום טוב مرحبا بالعالم لا abc שלום מקום def ")]
    public async Task ProfileImportsDeclaredLogicalTextFromIndependentGenerators(string fixture, string direction, string expected)
    {
        var ofd = await ImportAsync(Fixture(fixture), Profile(Enum.Parse<PdfTextBaseDirection>(direction)));
        Assert.Equal(expected, string.Concat(ofd.Pages.Single().Elements.OfType<OfdTextElement>().Select(text => text.Text)));
        Assert.NotEmpty(ofd.Pages.Single().Elements.OfType<OfdImageElement>().Single().Data);
    }

    [Fact]
    public async Task UnshapedNegativeImportsWhatItsPageShowsNotItsDeclaredSource()
    {
        // GetWords returns the declared source for the unshaped negative and the
        // reversed phrase for the shaped reference: wrong on both. The profile
        // follows the painted glyph order: the source for the shaped reference,
        // and the visible (reversed) reading for the negative.
        using (var parsed = PdfDocument.Open(Fixture("D3/fpdf2-two-word-fixed.pdf")))
            Assert.Equal("abc שלום מקום def", string.Join(" ", parsed.GetPage(1).GetWords().Select(word => word.Text)));
        var ofd = await ImportAsync(Fixture("D3/fpdf2-two-word-fixed.pdf"), Profile());
        Assert.Equal("abc םוקמ םולש def ", string.Concat(ofd.Pages[0].Elements.OfType<OfdTextElement>().Select(text => text.Text)));
    }

    [Fact]
    public async Task ImportedCharactersKeepTheirSourceGlyphOriginsInLogicalTextCode()
    {
        var pdf = Fixture("D3/fpdf2-two-word-cell.pdf");
        var expected = SourceOrigins(pdf);
        var ofd = await ImportAsync(pdf, Profile());
        var actual = new List<(char Char, double X, double Y)>();
        foreach (var text in ofd.Pages[0].Elements.OfType<OfdTextElement>())
        {
            var run = Assert.Single(text.Runs);
            Assert.Equal(text.Text, run.Text);
            Assert.Equal(0, text.FillColor.Alpha);
            var deltas = (run.DeltaX ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            Assert.Equal(run.Text.Length - 1, deltas.Length);
            var x = text.XMillimeters + run.XMillimeters;
            for (var index = 0; index < run.Text.Length; index++)
            {
                if (run.Text[index] != ' ') actual.Add((run.Text[index], x, text.YMillimeters + run.YMillimeters));
                if (index < deltas.Length) x += deltas[index];
            }
        }
        Assert.Equal(expected.Count, actual.Count);
        foreach (var (left, right) in expected.OrderBy(item => item.X).Zip(actual.OrderBy(item => item.X)))
        {
            Assert.Equal(left.Char, right.Char);
            Assert.InRange(right.X - left.X, -0.003, 0.003);
            Assert.InRange(right.Y - left.Y, -0.003, 0.003);
        }
        // Right-to-left words advance leftwards: their deltas are negative.
        var hebrew = ofd.Pages[0].Elements.OfType<OfdTextElement>().Single(text => text.Text == "שלום ");
        Assert.All(hebrew.Runs[0].DeltaX!.Split(' ').Take(3), value => Assert.StartsWith("-", value));
    }

    [Fact]
    public async Task LamAlefCharactersShareTheirLigatureGlyphOrigin()
    {
        var ofd = await ImportAsync(Fixture("F06/fpdf2-arabic-ltr-cell.pdf"), Profile());
        var ligature = ofd.Pages[0].Elements.OfType<OfdTextElement>().Single(text => text.Text == "لا ");
        Assert.StartsWith("0 ", ligature.Runs[0].DeltaX);
    }

    [Fact]
    public async Task ReexportCarriesLogicalOrderAtOriginalGlyphPositionsButIsNotReimportEvidence()
    {
        var source = Fixture("D3/fpdf2-two-word-cell.pdf");
        using var ofd = new MemoryStream();
        await new PdfToOfdConverter(Profile()).ConvertAsync(new MemoryStream(source), ofd);
        ofd.Position = 0;
        using var pdf = new MemoryStream();
        var face = File.ReadAllBytes(Path.Combine(Root(), "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-R22.ttf"));
        await new OfdToPdfConverter(new OfdToPdfOptions
        {
            NameOnlyFontBindings = new OfdNameOnlyFontBindings(new Dictionary<string, byte[]> { ["R22DejaVuSans"] = face })
        }).ConvertAsync(ofd, pdf);

        using (var exported = PdfDocument.Open(pdf.ToArray()))
        using (var original = PdfDocument.Open(source))
        {
            var letters = exported.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToList();
            // Content order is logical; positions remain the source's visual glyph origins.
            Assert.Equal("abcשלוםמקוםdef", string.Concat(letters.OrderBy(letter => letter.TextSequence).Select(letter => letter.Value)));
            var sourceLetters = original.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .OrderBy(letter => letter.StartBaseLine.X).ToList();
            foreach (var (left, right) in sourceLetters.Zip(letters.OrderBy(letter => letter.StartBaseLine.X)))
            {
                Assert.Equal(left.Value, right.Value);
                Assert.InRange(right.StartBaseLine.X - left.StartBaseLine.X, -0.05, 0.05);
            }
        }

        // The re-exported text layer is invisible (clipped); it is a producer
        // claim, not painted-glyph evidence, so the profile refuses to re-read it.
        var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new PdfToOfdConverter(Profile()).ConvertAsync(new MemoryStream(pdf.ToArray()), new MemoryStream()));
        Assert.Contains("clipping path", error.Message);
    }

    [Fact]
    public async Task FinalLayoutExportWithInvisibleSemanticTextIsRejectedRatherThanMisread()
    {
        // Regression from F06 probing: this PDF paints glyph outlines as paths and
        // carries a clipped semantic layer in logical order at LTR advances.
        // Reading that layer as geometry produced reversed Arabic.
        var face = new Ofdrw.Net.Layout.Text.TextFontFace("R22 DejaVu Sans",
            File.ReadAllBytes(Path.Combine(Root(), "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-R22.ttf")));
        var layout = HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest("مرحبا بالعالم لا",
            new TextStyle(new[] { new TextFontFamily("R22 DejaVu Sans", face) }, 5),
            new ParagraphLayoutOptions(150, null, TextParagraphDirection.RightToLeft)));
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(ParagraphLayoutOfdWriter.CreateDocument(layout, 170, 40, 10, 10), ofd);
        ofd.Position = 0;
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new PdfToOfdConverter(Profile(PdfTextBaseDirection.RightToLeft)).ConvertAsync(new MemoryStream(pdf.ToArray()), new MemoryStream()));
        Assert.Contains("PDF_TEXT_READING_ORDER_UNVERIFIED", error.Message);
        Assert.Contains("clipping path", error.Message);
    }

    [Theory]
    [InlineData("F06/pango-after-rtl-wrap.pdf", "Auto", "consistent with both base directions")]
    [InlineData("F06/fpdf2-hebrew-rtl-wrap.pdf", "Auto", "consistent with both base directions")]
    [InlineData("F06/pango-between-rtl-wrap.pdf", "LeftToRight", "/ActualText")]
    [InlineData("F06/fpdf2-hebrew-digits-cell.pdf", "Auto", "U+0031 (bidi class EN)")]
    [InlineData("F06/fpdf2-hebrew-niqqud-cell.pdf", "Auto", "bidi class NSM")]
    public async Task ProfileRejectsOutsideItsScopeWithoutWritingOutput(string fixture, string direction, string reason)
    {
        var prior = Encoding.ASCII.GetBytes("existing output");
        using var output = new MemoryStream();
        output.Write(prior);
        output.Position = 0;
        var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new PdfToOfdConverter(Profile(Enum.Parse<PdfTextBaseDirection>(direction))).ConvertAsync(new MemoryStream(Fixture(fixture)), output));
        Assert.Contains("PDF_TEXT_READING_ORDER_UNVERIFIED", error.Message);
        Assert.Contains(reason, error.Message);
        Assert.Contains("TextLayerMode=None", error.Message);
        Assert.Equal(prior, output.ToArray());
    }

    [Theory]
    [InlineData("F06/pango-two-word.pdf")]
    [InlineData("F06/fpdf2-arabic-ltr-cell.pdf")]
    [InlineData("F06/libreoffice-rtl.pdf")]
    public async Task DefaultOptionsKeepRejectingRightToLeftPages(string fixture)
    {
        var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new PdfToOfdConverter().ConvertAsync(new MemoryStream(Fixture(fixture)), new MemoryStream()));
        Assert.Contains("PDF_TEXT_READING_ORDER_UNVERIFIED", error.Message);
        Assert.Contains("contains right-to-left text", error.Message);
        Assert.Contains("RightToLeftTextImport=VisualLineProfile", error.Message);
    }

    [Fact]
    public async Task LeftToRightPagesAreUnchangedByTheProfile()
    {
        var pdf = Synthetic(Line(100, A, B, C, 0, D, E, F));
        var legacy = await ImportAsync(pdf, new PdfToOfdOptions());
        var profile = await ImportAsync(pdf, Profile());
        static string Shape(OfdDocumentPackage document) => string.Join("|", document.Pages[0].Elements.OfType<OfdTextElement>()
            .Select(text => $"{text.Text}@{text.XMillimeters},{text.YMillimeters},{text.WidthMillimeters},{text.HeightMillimeters},{text.Runs.Count}"));
        Assert.Equal("abc def ", string.Concat(legacy.Pages[0].Elements.OfType<OfdTextElement>().Select(text => text.Text)));
        Assert.Equal(Shape(legacy), Shape(profile));
    }

    [Fact]
    public void ContentStreamOrderDoesNotChangeTheReconstructedLine()
    {
        var visual = Line(100, A, B, C, 0, Mem, Vav, Lamed, Shin, 0, D, E, F);
        var shuffled = string.Join(" ", visual.Split("Tj ", StringSplitOptions.RemoveEmptyEntries).Reverse().Select(part => part + "Tj")) + " ";
        Assert.Equal("abc שלום def ", Lines(Synthetic(visual)).Single());
        Assert.Equal("abc שלום def ", Lines(Synthetic(shuffled)).Single());
    }

    [Fact]
    public async Task LeftToRightLineOnAnRtlPageKeepsTheLegacyWordWhileRtlLineIsRebuilt()
    {
        var pdf = Synthetic(Line(150, A, B, C) + Line(100, Mem, Vav, Lamed, Shin));
        using (var document = PdfDocument.Open(pdf))
        {
            var imported = PdfRtlVisualLineImport.Import(document, document.GetPage(1), PdfTextBaseDirection.Auto, 100, default);
            Assert.Equal("שלום ", string.Concat(Assert.Single(imported.Lines).Select(word => word.Text)));
            Assert.Equal(4, imported.Consumed.Count);
        }
        var ofd = await ImportAsync(pdf, Profile());
        var texts = ofd.Pages[0].Elements.OfType<OfdTextElement>().ToList();
        Assert.Equal(new[] { "abc ", "שלום " }, texts.Select(text => text.Text));
        Assert.True(string.IsNullOrEmpty(Assert.Single(texts[0].Runs).DeltaX));
        Assert.Equal(4, Assert.Single(texts[1].Runs).DeltaX!.Split(' ').Length);
    }

    [Theory]
    [InlineData("invisible", "rendering mode Neither")]
    [InlineData("clip-out", "is clipped")]
    [InlineData("clip-path", "non-rectangular clipping path")]
    [InlineData("alpha", "fully transparent")]
    [InlineData("reversed", "/ReversedChars")]
    [InlineData("optional", "optional content")]
    [InlineData("overlap", "overlapping glyphs")]
    [InlineData("rotated", "rotated or vertical")]
    [InlineData("gap", "gap wider than 3 em")]
    [InlineData("superscript", "different baseline")]
    [InlineData("presentation", "presentation form U+FEFB")]
    [InlineData("mixed-glyph", "both directions")]
    [InlineData("page-rotate", "rotated by 90 degrees")]
    [InlineData("no-order", "no logical order")]
    public void EachGuardRejectsItsOwnUnverifiedEvidence(string variant, string reason)
    {
        var line = Line(100, A, B, C, 0, Mem, Vav, Lamed, Shin, 0, D, E, F);
        var pdf = variant switch
        {
            "invisible" => Synthetic("3 Tr " + line),
            "clip-out" => Synthetic(line, prefix: "q 0 0 40 200 re W n ", suffix: "Q "),
            "clip-path" => Synthetic(line, prefix: "q 0 0 m 200 0 l 200 200 l h W n ", suffix: "Q "),
            "alpha" => Synthetic("/G0 gs " + line, resources: "/ExtGState << /G0 << /ca 0 >> >>"),
            "reversed" => Synthetic(line, prefix: "/ReversedChars BMC ", suffix: "EMC "),
            "optional" => Synthetic(line, prefix: "/OC /L0 BDC ", suffix: "EMC ",
                resources: "/Properties << /L0 << /Type /OCG /Name (L0) >> >>"),
            "overlap" => Synthetic(line + Show(38.5, 100, Mem)),
            "rotated" => Synthetic($"0 1 -1 0 60 60 Tm <{Mem:X4}> Tj "),
            "gap" => Synthetic(Line(100, A, B, C, 0, Mem, Vav, Lamed, Shin) + Show(150, 100, D)),
            "superscript" => Synthetic(line + Show(160, 104, A)),
            "presentation" => Synthetic(Line(100, A, B, C, 0, LamAlefForm, 0, D, E, F)),
            "mixed-glyph" => Synthetic(Line(100, A, B, C, 0, MixedGlyph, 0, D, E, F)),
            "page-rotate" => Synthetic(line, pageExtras: "/Rotate 90"),
            // Visually Hebrew sits at the left edge, Latin at the right: no base
            // direction resolves this logical candidate back to the observed order.
            "no-order" => Synthetic(Line(100, Mem, Vav, Lamed, Shin, 0, D, E, F)),
            _ => throw new ArgumentOutOfRangeException(nameof(variant)),
        };
        var error = Assert.Throws<NotSupportedException>(() => Lines(pdf));
        Assert.Contains("PDF_TEXT_READING_ORDER_UNVERIFIED", error.Message);
        Assert.Contains(reason, error.Message);
    }

    [Fact]
    public void RectangularClipThatContainsTheLineIsAccepted()
    {
        var line = Line(100, A, B, C, 0, Mem, Vav, Lamed, Shin, 0, D, E, F);
        Assert.Equal("abc שלום def ", Lines(Synthetic(line, prefix: "q 0 0 200 200 re W n ", suffix: "Q ")).Single());
        Assert.Equal("abc שלום def ", Lines(Synthetic("/G0 gs " + line, resources: "/ExtGState << /G0 << /ca 1 >> >>")).Single());
    }

    [Fact]
    public void ExplicitBaseDirectionResolvesAnAmbiguousLine()
    {
        var pdf = Synthetic(Line(100, A, B, C, 0, Mem, Vav, Lamed, Shin));
        Assert.Contains("both base directions", Assert.Throws<NotSupportedException>(() => Lines(pdf)).Message);
        Assert.Equal("abc שלום ", Lines(pdf, PdfTextBaseDirection.LeftToRight).Single());
        Assert.Equal("שלום abc ", Lines(pdf, PdfTextBaseDirection.RightToLeft).Single());
    }

    [Fact]
    public void BudgetAndCancellationStopTheImport()
    {
        var pdf = Synthetic(Line(100, A, B, C, 0, Mem, Vav, Lamed, Shin, 0, D, E, F));
        using var document = PdfDocument.Open(pdf);
        Assert.Throws<InvalidDataException>(() =>
            PdfRtlVisualLineImport.Import(document, document.GetPage(1), PdfTextBaseDirection.Auto, 2, default));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfRtlVisualLineImport.Import(document, document.GetPage(1), PdfTextBaseDirection.Auto, 100, cancelled.Token));
    }

    [Fact]
    public void OptionsRejectUndefinedValues()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfToOfdConverter(new PdfToOfdOptions { RightToLeftTextImport = (PdfRightToLeftTextImport)9 }));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfToOfdConverter(new PdfToOfdOptions { RightToLeftBaseDirection = (PdfTextBaseDirection)9 }));
    }

    // --- Synthetic one-font PDFs: CID n maps to Glyphs[n]; every glyph advances 0.5 em at 10 pt. ---

    private const int A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, Shin = 7, Lamed = 8, Vav = 9, Mem = 10, LamAlefForm = 11, MixedGlyph = 12;
    private static readonly string[] Glyphs = { "", "a", "b", "c", "d", "e", "f", "ש", "ל", "ו", "ם", "ﻻ", "aש" };

    private static string Show(double x, double y, int cid) =>
        string.Create(CultureInfo.InvariantCulture, $"1 0 0 1 {x} {y} Tm <{cid:X4}> Tj ");

    // Glyphs in visual left-to-right order; 0 inserts a 0.3 em word gap.
    private static string Line(double y, params int[] visual)
    {
        var builder = new StringBuilder();
        var x = 20d;
        foreach (var cid in visual)
        {
            if (cid == 0) { x += 3; continue; }
            builder.Append(Show(x, y, cid));
            x += 5;
        }
        return builder.ToString();
    }

    private static byte[] Synthetic(string text, string prefix = "", string suffix = "", string resources = "", string pageExtras = "")
    {
        var cmap = new StringBuilder("/CIDInit /ProcSet findresource begin 12 dict begin begincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Test def /CMapType 2 def\n1 begincodespacerange <0000> <FFFF> endcodespacerange\n");
        cmap.Append($"{Glyphs.Length - 1} beginbfchar\n");
        for (var cid = 1; cid < Glyphs.Length; cid++)
            cmap.Append($"<{cid:X4}> <{Convert.ToHexString(Encoding.BigEndianUnicode.GetBytes(Glyphs[cid]))}>\n");
        cmap.Append("endbfchar\nendcmap CMapName currentdict /CMap defineresource pop end end\n");
        static string Stream(string s) => $"<< /Length {Encoding.ASCII.GetByteCount(s)} >>\nstream\n{s}endstream";
        var objects = new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] {pageExtras} /Resources << /Font << /F1 4 0 R >> {resources} >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type0 /BaseFont /Synthetic /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 5 0 R >>",
            Stream(cmap.ToString()), Stream($"{prefix}BT /F1 10 Tf {text}ET {suffix}\n"),
            "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Synthetic /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 8 0 R /DW 500 /CIDToGIDMap /Identity >>",
            "<< /Type /FontDescriptor /FontName /Synthetic /Flags 4 /FontBBox [0 -200 500 800] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 >>" };
        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (var index = 0; index < objects.Length; index++)
        {
            offsets.Add(Encoding.UTF8.GetByteCount(pdf.ToString()));
            pdf.Append($"{index + 1} 0 obj\n{objects[index]}\nendobj\n");
        }
        var xref = Encoding.UTF8.GetByteCount(pdf.ToString());
        pdf.Append($"xref\n0 {objects.Length + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets) pdf.Append($"{offset:0000000000} 00000 n \n");
        pdf.Append($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        return Encoding.UTF8.GetBytes(pdf.ToString());
    }

    private static IReadOnlyList<string> Lines(byte[] pdf, PdfTextBaseDirection direction = PdfTextBaseDirection.Auto)
    {
        using var document = PdfDocument.Open(pdf);
        return PdfRtlVisualLineImport.Import(document, document.GetPage(1), direction, 1000, default)
            .Lines.Select(line => string.Concat(line.Select(word => word.Text))).ToList();
    }

    private static List<(char Char, double X, double Y)> SourceOrigins(byte[] pdf)
    {
        using var document = PdfDocument.Open(pdf);
        var page = document.GetPage(1);
        const double mm = 25.4 / 72;
        return page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Select(letter => (letter.Value[0], letter.StartBaseLine.X * mm, (page.Height - letter.StartBaseLine.Y) * mm)).ToList();
    }

    private static async Task<OfdDocumentPackage> ImportAsync(byte[] pdf, PdfToOfdOptions options)
    {
        using var output = new MemoryStream();
        await new PdfToOfdConverter(options).ConvertAsync(new MemoryStream(pdf), output);
        output.Position = 0;
        return await new OfdReader().ReadAsync(output);
    }

    private static readonly Dictionary<string, string> Hashes = new()
    {
        ["D3/fpdf2-two-word-cell.pdf"] = "7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8",
        ["D3/fpdf2-two-word-fixed.pdf"] = "7b69e87771891ba440bde051af33284f3e008b376744751bc27064391c366ea5",
        ["F06/fpdf2-arabic-ltr-cell.pdf"] = "a3ec912dff07c255dd6c04e9094b8834a924f254a7f64c15a4bff609104c1b68",
        ["F06/fpdf2-arabic-rtl-cell.pdf"] = "8abf6f54ae326e7bb7fbd597a09b88e37ab0b73e3959a80d176c6f88861052c0",
        ["F06/fpdf2-between-rtl-zero-margin.pdf"] = "9c40fc51de911841fecf8dc44dafe8c88bd392bf525a4f9f85112df8648eb3c1",
        ["F06/fpdf2-hebrew-digits-cell.pdf"] = "c9a087c9856f297caf8d4f38cf8545cc0db34a9b5e23dc9dedfd17d5e0f36f26",
        ["F06/fpdf2-hebrew-niqqud-cell.pdf"] = "f4ac0949f4b78e9298a03c01788d0b628676df3535736d100063eda322dd053f",
        ["F06/fpdf2-hebrew-rtl-wrap.pdf"] = "39678a380122d6e62940723cdf23a481ac467130ae14bc34254407dec7b5c02f",
        ["F06/libreoffice-rtl.pdf"] = "3db83d540e7e6d9772a5a01861550541ac69c74669e928d02754b601ad5c2f07",
        ["F06/pango-after-rtl-wrap.pdf"] = "3c7d4b6a973097e14c2761f0045d6ed86bfcc8d178f131f44d2f5ecdafe7e7bc",
        ["F06/pango-between-rtl-wrap.pdf"] = "e61b35acc0a14847253b68fd6a01728917ab025392b441637116cf9afbe417b5",
        ["F06/pango-two-word.pdf"] = "c20d0d39456727ea7f967b679120cb6ffc0f52bb89cfc7816b746f48d7a99513",
    };

    private static byte[] Fixture(string name)
    {
        var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
        Assert.Equal(Hashes[name], Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        return bytes;
    }

    private static string Root()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return root!.FullName;
    }
}
