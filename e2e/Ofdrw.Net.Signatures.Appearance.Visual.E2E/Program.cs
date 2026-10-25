using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

if (args.Length != 1) throw new ArgumentException("Expected an output directory.");
var output = Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
var pageMarks = new[]
{
    new OfdColor(35, 80, 130), new OfdColor(75, 110, 145),
    new OfdColor(125, 80, 55), new OfdColor(145, 100, 65),
    new OfdColor(65, 125, 70), new OfdColor(80, 145, 95),
    new OfdColor(135, 60, 130), new OfdColor(150, 85, 150)
};
var document = new OfdDocumentPackage();
for (var index = 0; index < 8; index++)
{
    var page = new OfdPage { Index = index, WidthMillimeters = 100, HeightMillimeters = 80 };
    page.Elements.Add(new OfdPathElement
    {
        XMillimeters = 5, YMillimeters = 5, WidthMillimeters = 12, HeightMillimeters = 12,
        AbbreviatedData = "M 0 0 L 12 0 L 12 12 L 0 12 C",
        Fill = true, Stroke = false, FillColor = pageMarks[index]
    });
    document.Pages.Add(page);
}
var unsignedPath = Path.Combine(output, "r11-unsigned.ofd");
using (var file = File.Create(unsignedPath)) await new OfdPackageWriter().WriteAsync(document, file);
using var picture = new Image<Rgb24>(80, 40);
for (var y = 0; y < picture.Height; y++)
for (var x = 0; x < picture.Width; x++)
    picture[x, y] = (x < 40, y < 20) switch
    {
        (true, true) => new Rgb24(225, 30, 30),
        (false, true) => new Rgb24(30, 180, 40),
        (true, false) => new Rgb24(35, 70, 220),
        _ => new Rgb24(220, 190, 30)
    };
using var image = new MemoryStream();
picture.SaveAsPng(image);
File.WriteAllBytes(Path.Combine(output, "r11-asymmetric-seal.png"), image.ToArray());
var options = new OfdSignatureOptions();
foreach (var (edge, firstPage) in new[]
{
    (OfdStampEdge.Left, 1), (OfdStampEdge.Right, 3),
    (OfdStampEdge.Top, 5), (OfdStampEdge.Bottom, 7)
}) options.StampAppearances.Add(new OfdRidingStampPlacement(edge, firstPage, 2, 40, 20)
{ MarginMillimeters = 2 });
var signedPath = Path.Combine(output, "r11-four-edges-signed.ofd");
using (var source = File.OpenRead(unsignedPath))
using (var target = File.Create(signedPath))
    await new OfdSignatureService().SignAsync(source, target, new SyntheticProvider(image.ToArray()), options);
using (var source = File.OpenRead(signedPath))
{
    var report = await new OfdSignatureVerifier().VerifyAsync(source);
    if (!report.ReferenceIntegrityValid || report.FullyValid)
        throw new InvalidDataException("Visual fixture must have intact references and an unsupported synthetic signed value.");
}
var pdfPath = Path.Combine(output, "r11-four-edges-signed.pdf");
using (var source = File.OpenRead(signedPath))
using (var target = File.Create(pdfPath)) await new OfdToPdfConverter().ConvertAsync(source, target);
for (var index = 0; index < 8; index++)
{
    using var source = File.OpenRead(signedPath);
    using var target = File.Create(Path.Combine(output, $"r11-page-{index + 1:00}.png"));
    await new OfdToImageConverter(new OfdToImageOptions
    { PixelsPerMillimeter = 3 }).ConvertAsync(source, target, index);
}
var hashes = Directory.GetFiles(output).OrderBy(path => Path.GetFileName(path), StringComparer.Ordinal)
    .ToDictionary(path => Path.GetFileName(path), path => new
    {
        bytes = new FileInfo(path).Length,
        sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)))
    });
File.WriteAllText(Path.Combine(output, "r11-visual-manifest.json"), JsonSerializer.Serialize(new
{
    source = "Synthetic picture provider; no cryptographic or independent reader acceptance",
    pageRanges = new[] { "left:1-2", "right:3-4", "top:5-6", "bottom:7-8" },
    dimensionsMillimeters = new { page = "100x80", seal = "40x20" },
    chain = "OFD -> SignAsync -> OFD -> PDF and per-page PNG", hashes
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine("Eight signed OFD pages, one PDF and eight PNGs generated for R11 visual inspection.");

sealed class SyntheticProvider(byte[] seal) : IOfdSignatureProvider
{
    public string ProviderName => "R11 visual fixture only";
    public string Company => "Ofdrw.Net Tests";
    public string Version => "1";
    public string SignatureMethod => "urn:ofdrw:r11:synthetic";
    public string SignatureType => "Seal";
    public byte[]? SealData => seal;
    public Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation,
        CancellationToken cancellationToken = default)
        => Task.FromResult(Encoding.UTF8.GetBytes("not-a-cryptographic-signature"));
}
