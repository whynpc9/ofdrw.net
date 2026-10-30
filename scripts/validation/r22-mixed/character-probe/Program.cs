using System.Security.Cryptography;
using System.Text.Encodings.Web;
using System.Text.Json;
using UglyToad.PdfPig;

if (args.Length != 2) throw new ArgumentException("request.json output.json");
using var request = JsonDocument.Parse(File.ReadAllBytes(args[0]));
var rows = new List<object>();
foreach (var input in request.RootElement.GetProperty("inputs").EnumerateArray())
{
    var role = input.GetProperty("role").GetString()!;
    var path = input.GetProperty("path").GetString()!;
    var bytes = File.ReadAllBytes(path);
    var sha = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    if (sha != input.GetProperty("sha256").GetString()) throw new InvalidDataException("Input hash changed: " + role);
    foreach (var useActualText in new[] { false, true })
    {
        using var document = PdfDocument.Open(bytes, new ParsingOptions { UseActualText = useActualText });
        if (document.NumberOfPages != 1) throw new InvalidDataException("Expected one page: " + role);
        var page = document.GetPage(1);
        rows.Add(new {
            role, path, sha256 = sha, useActualText,
            pageWidth = page.Width, pageHeight = page.Height,
            rawPageText = page.Text,
            rawPageCodepoints = page.Text.Select(c => $"U+{(int)c:X4}").ToArray(),
            words = page.GetWords().Select((word, index) => new {
                index, word.Text,
                bounds = new { word.BoundingBox.Left, word.BoundingBox.Bottom,
                    word.BoundingBox.Right, word.BoundingBox.Top }
            }).ToArray(),
            letters = page.Letters.Select((letter, index) => new {
                index, letter.Value, letter.FontName,
                renderingMode = letter.RenderingMode.ToString(),
                glyphRectangle = new { letter.GlyphRectangle.Left, letter.GlyphRectangle.Bottom,
                    letter.GlyphRectangle.Right, letter.GlyphRectangle.Top },
                boundingBox = new { letter.BoundingBox.Left, letter.BoundingBox.Bottom,
                    letter.BoundingBox.Right, letter.BoundingBox.Top },
                baseline = new { x = letter.StartBaseLine.X, y = letter.StartBaseLine.Y }
            }).ToArray()
        });
    }
}
var assembly = typeof(PdfDocument).Assembly;
var output = new {
    assembly = assembly.FullName,
    assemblySha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(assembly.Location))).ToLowerInvariant(),
    rows
};
File.WriteAllText(args[1], JsonSerializer.Serialize(output, new JsonSerializerOptions {
    WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping }) + "\n");
