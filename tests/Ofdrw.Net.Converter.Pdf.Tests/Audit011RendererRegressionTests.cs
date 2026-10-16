using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[Collection("Font registry")]
public sealed class Audit011RendererRegressionTests
{
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts/narrow.ttf"));
    private static OfdDocumentPackage Package(byte[]? font = null)
    {
        var package = new OfdDocumentPackage(); package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 70 });
        if (font is not null)
        {
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "audit011", Data = font });
            package.Pages[0].Elements.Add(new OfdTextElement { FontResourceId = "10", Text = "AV office", XMillimeters = 8,
                YMillimeters = 12, WidthMillimeters = 85, HeightMillimeters = 12, FontSizeMillimeters = 4 });
        }
        return package;
    }
    private static async Task<byte[]> Save(OfdDocumentPackage package)
    { using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); return output.ToArray(); }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Audit011_PdfAndSvgEntry_EnforceTinyBudgetBeforeOutput(bool svg)
    {
        using var ofd = new MemoryStream(await Save(Package(Font))); using var output = new MemoryStream();
        var failure = await Assert.ThrowsAsync<InvalidDataException>(() => svg
            ? new OfdToSvgConverter(new OfdPackageLoadOptions(), 1).ConvertAsync(ofd, output)
            : new OfdToPdfConverter(new OfdToPdfOptions(), 1).ConvertAsync(ofd, output));
        Assert.Contains("MaximumCmapParseWork", failure.Message); Assert.Equal(0, output.Length);
    }

    private sealed class Resolver(byte[] primary, byte[] fallback, Action<int>? read = null) : IFontResolver
    {
        public int Reads { get; private set; }
        public string DefaultFontName => "audit-default";
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic) => new(familyName);
        public byte[] GetFont(string faceName) { Reads++; read?.Invoke(Reads); return faceName == DefaultFontName ? fallback : primary; }
    }

    [Theory]
    [InlineData("embedded")] [InlineData("name-only")] [InlineData("selected")] [InlineData("fallback")]
    public void Audit011_Pdf_AllFourFontPathsPropagateCumulativeExhaustion(string path)
    {
        var source = new OpenTypeFace(Font);
        source.Tables["cmap"] = OpenTypeCmap.Build(new Dictionary<int, int> { ['A'] = 1 });
        var primary = source.Build();
        var probe = new FontParseBudget(); _ = new OpenTypeCmap(new OpenTypeFace(primary), probe);
        var budget = new FontParseBudget(path == "fallback" ? probe.Consumed : 1);
        var resource = new OfdFontResource { Id = "10", FontName = "audit-optional-probe", Bold = path == "name-only", Data = path == "embedded" ? primary : [] };
        var resolver = new Resolver(primary, Font);
        var failure = Assert.Throws<InvalidDataException>(() =>
        {
            var context = new DocumentFontContext([resource], resolver, budget);
            context.Resolve(new OfdTextElement { FontResourceId = "10", Text = path == "fallback" ? "B" : "A" }, out _);
        });
        Assert.Contains("MaximumCmapParseWork", failure.Message); Assert.True(budget.IsExhausted);
        Assert.Equal(path == "embedded" ? 0 : path == "fallback" ? 2 : 1, resolver.Reads);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void Audit011_Pdf_OptionalFontDiscoveryAndFallbackDoNotSwallowCancellation(bool fallback)
    {
        var source = new OpenTypeFace(Font);
        source.Tables["cmap"] = OpenTypeCmap.Build(new Dictionary<int, int> { ['A'] = 1 });
        using var cancellation = new CancellationTokenSource();
        var resolver = new Resolver(source.Build(), Font, count => { if (count == (fallback ? 2 : 1)) cancellation.Cancel(); });
        var resource = new OfdFontResource { Id = "10", FontName = "audit-cancel", Bold = !fallback };
        Assert.Throws<OperationCanceledException>(() =>
        {
            var context = new DocumentFontContext([resource], resolver,
                new FontParseBudget(cancellationToken: cancellation.Token), cancellation.Token);
            context.Resolve(new OfdTextElement { FontResourceId = "10", Text = "B" }, out _);
        });
        Assert.Equal(fallback ? 2 : 1, resolver.Reads);
    }

    [Fact]
    public async Task Audit011_PdfNestedSignature_SharesParentParseBudget()
    {
        var font = Font;
        var budget = new FontParseBudget(); _ = new OpenTypeCmap(new OpenTypeFace(font), budget);
        var nested = await Save(Package(font));
        using var outer = new MemoryStream(); outer.Write(await Save(Package(font)));
        using (var zip = new ZipArchive(outer, ZipArchiveMode.Update, true))
        {
            var entry = zip.GetEntry("OFD.xml")!; XDocument root;
            using (var input = entry.Open()) root = XDocument.Load(input);
            var body = root.Descendants().Single(x => x.Name.LocalName == "DocBody");
            var documentPath = body.Elements().Single(x => x.Name.LocalName == "DocRoot").Value;
            var prefix = documentPath[..documentPath.LastIndexOf('/')];
            XDocument document; using (var input = zip.GetEntry(documentPath)!.Open()) document = XDocument.Load(input);
            var pageId = document.Descendants().First(x => x.Name.LocalName == "Page").Attribute("ID")!.Value;
            body.Add(new XElement(body.Name.Namespace + "Signatures", prefix + "/Signs/Signatures.xml"));
            entry.Delete(); using (var output = zip.CreateEntry("OFD.xml").Open()) root.Save(output);
            void Add(string name, byte[] bytes) { using var stream = zip.CreateEntry(name).Open(); stream.Write(bytes); }
            Add(prefix + "/Signs/Signatures.xml", Encoding.UTF8.GetBytes("<Signatures><Signature BaseLoc=\"S/Signature.xml\"/></Signatures>"));
            Add(prefix + "/Signs/S/Signature.xml", Encoding.UTF8.GetBytes($"<Signature><SignedInfo><Seal BaseLoc=\"Seal.ofd\"/><StampAnnot PageRef=\"{pageId}\" Boundary=\"0 0 10 10\"/></SignedInfo></Signature>"));
            Add(prefix + "/Signs/S/Seal.ofd", nested);
        }
        outer.Position = 0; using var pdf = new MemoryStream();
        var failure = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToPdfConverter(new OfdToPdfOptions(), budget.Consumed).ConvertAsync(outer, pdf));
        Assert.Contains("MaximumCmapParseWork", failure.Message); Assert.Equal(0, pdf.Length);
        // A finite default permits the same ordinary+nested document.
        outer.Position = 0; await new OfdToPdfConverter().ConvertAsync(outer, pdf); Assert.True(pdf.Length > 0);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Audit011_PdfAndSvgEntry_ObserveCallerCancellation(bool svg)
    {
        using var input = new MemoryStream(await Save(Package(Font))); using var output = new MemoryStream();
        var token = new CancellationToken(true);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => svg
            ? new OfdToSvgConverter().ConvertAsync(input, output, cancellationToken: token)
            : new OfdToPdfConverter().ConvertAsync(input, output, cancellationToken: token));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task Audit011_RealFontArtifact_ProducesFreshEmbeddedPdfAndSvg()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(directory!.FullName, "Ofdrw.Net.sln"))) directory = directory.Parent;
        var artifacts = Path.Combine(directory.FullName, "artifacts/plan-011/font-pages"); Directory.CreateDirectory(artifacts);
        var package = Package(File.ReadAllBytes(Path.Combine(directory.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/LXGWWenKai-Regular.ttf")));
        var text = (OfdTextElement)package.Pages[0].Elements[0]; text.Text = "中文 AV office ffi é e\u0301";
        text.FillColor = new OfdColor(20, 40, 150);
        package.Pages[0].Elements.Add(new OfdTextElement { FontResourceId = "10", Text = "Bold office", Weight = 700,
            XMillimeters = 8, YMillimeters = 30, WidthMillimeters = 80, HeightMillimeters = 12, FontSizeMillimeters = 4 });
        var bytes = await Save(package); File.WriteAllBytes(Path.Combine(artifacts, "font-page.ofd"), bytes);
        using var input = new MemoryStream(bytes); using var pdf = File.Create(Path.Combine(artifacts, "font-page.pdf"));
        await new OfdToPdfConverter().ConvertAsync(input, pdf);
        input.Position = 0; using var svg = File.Create(Path.Combine(artifacts, "font-page.svg"));
        await new OfdToSvgConverter().ConvertAsync(input, svg);
        Assert.True(pdf.Length > 0); Assert.True(svg.Length > 0);
    }
}
