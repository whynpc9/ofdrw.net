using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Encodings.Web;
using UglyToad.PdfPig;

if (args.Length != 2) throw new ArgumentException("matrix.json output.json");
using var manifest = JsonDocument.Parse(File.ReadAllBytes(args[0]));
var rows = new List<object>();
foreach (var item in manifest.RootElement.GetProperty("cases").EnumerateArray())
{
    var path = item.GetProperty("path").GetString()!;
    var bytes = File.ReadAllBytes(path);
    var sha = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    if (sha != item.GetProperty("sha256").GetString()) throw new InvalidDataException("Input hash changed: " + path);
    foreach (var useActualText in new[] { false, true })
    {
        using var document = PdfDocument.Open(bytes, new ParsingOptions { UseActualText = useActualText });
        if (document.NumberOfPages != 1) throw new InvalidDataException("Expected one page: " + path);
        var page = document.GetPage(1);
        var words = page.GetWords().ToArray();
        var geometry = item.GetProperty("expectedWords").EnumerateArray().Select(expected =>
        {
            var word = expected.GetProperty("text").GetString()!;
            var matches = words.Where(value => value.Text == word).ToArray();
            var expectedCount = expected.GetProperty("count").GetInt32();
            var distinct = matches.Select(value => value.BoundingBox.ToString()).Distinct().Count();
            return new { Word = word, ExpectedCount = expectedCount, Count = matches.Length,
                DistinctRectangles = distinct,
                CountAndDistinctBounds = matches.Length == expectedCount && distinct == expectedCount &&
                    matches.All(value => value.BoundingBox.Width > 0 && value.BoundingBox.Height > 0) };
        }).ToArray();
        rows.Add(new { Name = item.GetProperty("name").GetString(), Path = path, Sha256 = sha,
            UseActualText = useActualText, Text = page.Text,
            Codepoints = page.Text.Select(value => $"U+{(int)value:X4}"),
            ExactSource = page.Text == item.GetProperty("expected").GetString(),
            Words = words.Select(word => new { word.Text, word.BoundingBox.Left, word.BoundingBox.Bottom,
                word.BoundingBox.Right, word.BoundingBox.Top }), Geometry = geometry,
            Scope = "GetWords token geometry; not reader Find/highlight acceptance" });
    }
}
var assembly = typeof(PdfDocument).Assembly;
var receipt = new { Assembly = assembly.FullName,
    AssemblySha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(assembly.Location))).ToLowerInvariant(),
    Rows = rows };
File.WriteAllText(args[1], JsonSerializer.Serialize(receipt, new JsonSerializerOptions {
    WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping }) + "\n");
Console.WriteLine($"Recorded {rows.Count} independent PdfPig reads, including failures.");
