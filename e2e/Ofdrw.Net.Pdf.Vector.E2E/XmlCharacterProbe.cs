using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using SkiaSharp;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Tokens;

internal static class XmlCharacterProbe
{
    internal static async Task Run(string output)
    {
        var rows = new List<object>();
        foreach (var (name, character, expectedSha) in new[] {
            ("ufffe", '\uFFFE', "9a6a3e951ff19a23cc558ec013e37cfeb145a16878853657c9a72caf69d7f86f"),
            ("uffff", '\uFFFF', "cba6ee6848a7efbcdcd9bcb138cdc79d74c35efbca0b77b05fc57cae45afd01c") })
        {
            using var compressed = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "testdata", "xml-" + name + ".pdf.gz"));
            using var gzip = new GZipStream(compressed, CompressionMode.Decompress); using var expanded = new MemoryStream(); gzip.CopyTo(expanded);
            var source = expanded.ToArray(); var sha = Convert.ToHexStringLower(SHA256.HashData(source));
            if (sha != expectedSha) throw new Exception("XML fixture input SHA changed.");
            using (var pdf = PdfDocument.Open(source))
            {
                if (pdf.NumberOfPages != 3 || pdf.GetPage(2).Letters.Single().Value != character.ToString())
                    throw new Exception("Actual PDF ToUnicode no longer maps the XML-invalid original glyph.");
                var scanner = pdf.Structure.TokenScanner;
                var stream = (StreamToken)scanner.Get(new UglyToad.PdfPig.Core.IndirectReference(7, 0))!.Data;
                using var fontData = SKData.CreateCopy(stream.Data.ToArray()); using var face = SKTypeface.FromData(fontData);
                if (face.GetGlyph(character) != 5) throw new Exception("Actual embedded cmap does not match source CID 0005.");
                using var font = new SKFont(face, 1000); using var path = font.GetGlyphPath(5);
                if (path is null || path.IsEmpty) throw new Exception("Mapped fixture glyph lost its visible Z outline.");
            }
            using (var input = new MemoryStream(source))
            using (var sentinel = new MemoryStream())
            {
                sentinel.Write(new byte[] { 4, 5, 6 }); sentinel.Position = 0;
                try { await new PdfVectorToOfdConverter().ConvertAsync(input, sentinel); throw new Exception("Invalid XML native text accepted."); }
                catch (NotSupportedException error) when (error.Message.Contains("TEXT_XML_CHARACTER")) { }
                if (!sentinel.ToArray().SequenceEqual(new byte[] { 4, 5, 6 }) || sentinel.Position != 0) throw new Exception("Fail policy changed sentinel.");
            }
            string defaultFailure;
            using (var input = new MemoryStream(source))
            using (var sentinel = new MemoryStream())
            {
                sentinel.Write(new byte[] { 4, 5, 6 }); sentinel.Position = 0;
                try { await new PdfVectorToOfdConverter(new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage }).ConvertAsync(input, sentinel); throw new Exception("Default invisible text unexpectedly serialized XML-invalid original text."); }
                catch (ArgumentException error) when (error.Message.Contains("invalid character")) { defaultFailure = error.GetType().FullName + ": " + error.Message; }
                if (!sentinel.ToArray().SequenceEqual(new byte[] { 4, 5, 6 }) || sentinel.Position != 0) throw new Exception("Default fallback failure changed sentinel.");
            }
            var options = new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage };
            options.Compatibility.TextLayerMode = PdfTextLayerMode.None;
            using var selected = new MemoryStream(source); using var destination = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(selected, destination);
            if (!result.Pages.Select(p => p.IsNative).SequenceEqual(new[] { true, false, true }) || !result.Pages[1].Diagnostic.Contains("TEXT_XML_CHARACTER"))
                throw new Exception("XML-invalid source did not fall back as a complete middle page.");
            destination.Position = 0; var package = await new OfdReader().ReadAsync(destination);
            var middle = package.Pages[1];
            if (package.Pages.Count != 3 || middle.Elements.Count != 1 || middle.Elements.Single() is not OfdImageElement ||
                !string.Concat(package.Pages[0].Elements.OfType<OfdTextElement>().Select(t => t.Text)).Contains("Before 正常中文") ||
                !string.Concat(package.Pages[2].Elements.OfType<OfdTextElement>().Select(t => t.Text)).Contains("After 前后"))
                throw new Exception("Raster middle page has native remnants or neighboring native text was lost.");
            var prefix = Path.Combine(output, "xml-" + name);
            File.WriteAllBytes(prefix + "-source.pdf", source); File.WriteAllBytes(prefix + ".ofd", destination.ToArray()); destination.Position = 0;
            await using (var pdf = File.Create(prefix + ".pdf")) await new OfdToPdfConverter().ConvertAsync(destination, pdf);
            rows.Add(new { name, sourceSha256 = sha, sourceGlyph = 5, codepoint = (int)character, result, defaultInvisibleFallbackFailure = defaultFailure,
                explicitTextLayerMode = "None", searchableOriginalTextOnRasterPage = false, sentinelUnchanged = true });
        }
        File.WriteAllText(Path.Combine(output, "xml-character-report.json"), JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true }));
    }
}
