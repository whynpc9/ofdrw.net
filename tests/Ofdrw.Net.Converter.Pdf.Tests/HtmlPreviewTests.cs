using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Converter.Svg.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class HtmlPreviewTests
{
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private const string Attack = "</title><script>fetch('http://127.0.0.1:9/attack')</script>&\"";
    internal static async Task<byte[]> Fixture(bool fonts = false, bool images = true)
    {
        var package = new OfdDocumentPackage();
        if (fonts) package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "fixture", Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "style-metrics.ttf")) });
        for (var i = 0; i < 3; i++)
        {
            var page = new OfdPage { Index = i, WidthMillimeters = i == 1 ? 150 : 100, HeightMillimeters = 120 };
            package.Pages.Add(page);
            page.Elements.Add(new OfdTextElement { FontName = fonts ? "fixture" : "sans-serif", FontResourceId = fonts ? "10" : "",
                Text = fonts ? "ABC" : $"Page {i + 1} 中文 Wi </text><script> & \"", XMillimeters = 5, YMillimeters = 8, WidthMillimeters = 90, HeightMillimeters = 16,
                FontSizeMillimeters = 4, Weight = i == 1 ? 700 : 400, Italic = i == 2, FillColor = new OfdColor((byte)(i * 80), 20, 70) });
            page.Elements.Add(new OfdPathElement { XMillimeters = 5, YMillimeters = 30, WidthMillimeters = 50, HeightMillimeters = 20,
                AbbreviatedData = "M 0 0 L 50 0 L 50 20 L 0 20 C", Stroke = true, Fill = true, FillColor = new OfdColor(230, 240, 250), LineWidthMillimeters = 0.4,
                ClippingXml = "<Clips xmlns=\"http://www.ofdspec.org\"><Clip><Area><Path><AbbreviatedData>M 0 0 L 45 0 L 45 20 C</AbbreviatedData></Path></Area></Clip></Clips>" });
            if (images)
            {
                using var pixels = new Image<Rgba32>(4, 4, i == 2 ? Color.Blue : Color.Red); using var png = new MemoryStream(); pixels.SaveAsPng(png);
                page.Elements.Add(new OfdImageElement { XMillimeters = 60, YMillimeters = 35, WidthMillimeters = 20, HeightMillimeters = 15, Data = png.ToArray() });
            }
        }
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); return output.ToArray();
    }
    private static async Task<byte[]> Html(byte[] ofd, OfdHtmlOptions? options = null, IReadOnlyList<int>? pages = null)
    { using var input = new MemoryStream(ofd); using var output = new MemoryStream(); await new OfdToHtmlConverter(options ?? new()).ConvertAsync(input, output, pages); return output.ToArray(); }
    private static XDocument Parse(byte[] html) => XDocument.Parse(System.Text.Encoding.UTF8.GetString(html));

    [Fact]
    public async Task OneRead_IdIsolation_FontScope_DuplicateOrder_AndEscaping()
    {
        var source = await Fixture(fonts: true); var calls = 0;
        var converter = new OfdToHtmlConverter(new() { Title = Attack }, null, async (stream, options, token) =>
        { calls++; return await new OfdReader().ReadAsync(stream, options, token); });
        using var input = new NonSeekable(source); using var output = new MemoryStream();
        await converter.ConvertAsync(input, output, new[] { 2, 0, 2 });
        Assert.Equal(1, calls); Assert.True(input.ReadBytes >= source.Length); Assert.Equal(0, input.Seeks);
        var document = Parse(output.ToArray()); var sections = document.Descendants("section").ToArray();
        Assert.Equal(new[] { "3", "1", "3" }, sections.Select(s => s.Attribute("data-source-page")!.Value));
        Assert.Equal(Attack, document.Descendants("title").Single().Value);
        Assert.Empty(document.Descendants().Where(n => n.Name.LocalName == "script"));
        Assert.Empty(document.Descendants().Attributes().Where(a => a.Name.LocalName.StartsWith("on", StringComparison.OrdinalIgnoreCase)));
        var ids = document.Descendants().Attributes("id").Select(a => a.Value).ToArray(); Assert.Equal(ids.Length, ids.Distinct().Count());
        for (var i = 0; i < sections.Length; i++)
        {
            var section = sections[i]; var localIds = section.Descendants().Attributes("id").Select(a => a.Value).ToArray();
            foreach (var reference in section.Descendants().Attributes("clip-path")) Assert.Contains(reference.Value[5..^1], localIds);
            var family = section.Descendants(Svg + "text").First().Attribute("font-family")!.Value;
            Assert.StartsWith($"ofd-preview-font-{i}-", family); Assert.Contains(family, section.Descendants(Svg + "style").Single().Value);
        }
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } directory)
        { Directory.CreateDirectory(directory); await File.WriteAllBytesAsync(Path.Combine(directory, "three-pages.ofd"), source); await File.WriteAllBytesAsync(Path.Combine(directory, "duplicate-fonts.html"), output.ToArray());
            var mixed = await Fixture(); await File.WriteAllBytesAsync(Path.Combine(directory, "mixed.ofd"), mixed);
            await File.WriteAllBytesAsync(Path.Combine(directory, "mixed.html"), await Html(mixed, new() { Title = Attack }, new[] { 2, 0, 2 }));
            await File.WriteAllBytesAsync(Path.Combine(directory, "mixed-all.html"), await Html(mixed)); }
    }
    [Fact]
    public async Task ControlledStrokeRemainsIdenticalToSingleSvg()
    {
        var source = await Audit015RegressionTests.Fixture("Alpha='128' DashPattern='6 2' DashOffset='1.5' Cap='Round' Join='Round'");
        var html = Parse(await Html(source)); using var standalone = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(source), standalone);
        var expected = XDocument.Parse(System.Text.Encoding.UTF8.GetString(standalone.ToArray())).Descendants(Svg + "path").Last();
        var actual = html.Descendants(Svg + "path").Last(); Assert.True(XNode.DeepEquals(expected, actual));
        Assert.Equal("6 2", actual.Attribute("stroke-dasharray")!.Value);
    }
    [Fact]
    public async Task BlankPageIsPreserved()
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 90, HeightMillimeters = 70 });
        using var source = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, source); var bytes = await Html(source.ToArray());
        var document = Parse(bytes); Assert.Single(document.Descendants("section")); Assert.Empty(document.Descendants(Svg + "svg").Single().Elements());
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } directory) await File.WriteAllBytesAsync(Path.Combine(directory, "blank.html"), bytes);
    }
    [Fact]
    public async Task TitleAndTextStayLiteral()
    {
        var document = Parse(await Html(await Fixture(), new() { Title = Attack }));
        Assert.Equal(Attack, document.Descendants("title").Single().Value);
        Assert.Contains("</text><script> & \"", document.Descendants(Svg + "text").First().Value);
    }
    [Fact]
    public async Task ExactUtf8OutputBoundaryIncludesShellResourcesEscapingAndDuplicates()
    {
        var source = await Fixture(); var expected = await Html(source, pages: new[] { 2, 0, 2 });
        Assert.Equal(expected, await Html(source, new() { MaxOutputBytes = expected.Length, MaxPageBytes = expected.Length }, new[] { 2, 0, 2 }));
        await Assert.ThrowsAsync<InvalidDataException>(() => Html(source, new() { MaxOutputBytes = expected.Length - 1, MaxPageBytes = expected.Length - 1 }, new[] { 2, 0, 2 }));
    }
    [Theory]
    [InlineData("page")][InlineData("nodes")][InlineData("characters")][InlineData("resources")]
    public async Task LimitsPreservePublicOutput(string limit)
    {
        var options = new OfdHtmlOptions();
        switch (limit) { case "page": options.MaxPageBytes = 10; break; case "nodes": options.MaxGeneratedNodes = 40; break; case "characters": options.MaxGeneratedCharacters = 10000; break; case "resources": options.MaxEmbeddedResourceBytes = 1; break; }
        using var output = new MemoryStream(); output.WriteByte(42);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToHtmlConverter(options).ConvertAsync(new MemoryStream(Fixture().GetAwaiter().GetResult()), output));
        Assert.Equal(new byte[] { 42 }, output.ToArray());
    }
    [Theory]
    [InlineData(-1)][InlineData(3)]
    public async Task InvalidPagePreservesOutput(int page)
    { using var output = new MemoryStream(); output.WriteByte(42); await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(Fixture().GetAwaiter().GetResult()), output, new[] { page })); Assert.Equal(new byte[] { 42 }, output.ToArray()); }
    [Fact]
    public async Task EmptySelectionAndPreCancellationReadNothing()
    {
        using var input = new NonSeekable(await Fixture()); using var output = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentException>(() => new OfdToHtmlConverter().ConvertAsync(input, output, Array.Empty<int>()));
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new OfdToHtmlConverter().ConvertAsync(input, output, cancellationToken: cancel.Token));
        Assert.Equal(0, input.ReadBytes); Assert.Empty(output.ToArray());
    }
    [Fact]
    public async Task OptionsAndSelectionAreSnapshots_ConcurrentCallsAreIsolated()
    {
        var options = new OfdHtmlOptions { Title = "snapshot" }; var converter = new OfdToHtmlConverter(options); options.Title = "changed";
        var source = await Fixture(); var outputs = await Task.WhenAll(Enumerable.Range(0, 4).Select(async _ =>
        { using var output = new MemoryStream(); await converter.ConvertAsync(new MemoryStream(source), output); return output.ToArray(); }));
        Assert.All(outputs, o => Assert.Equal(outputs[0], o)); Assert.Equal("snapshot", Parse(outputs[0]).Descendants("title").Single().Value);
        var selection = new[] { 2, 0 }; var entered = new TaskCompletionSource(); var resume = new TaskCompletionSource();
        var slow = new OfdToHtmlConverter(new(), null, async (s, l, t) => { entered.SetResult(); await resume.Task; return await new OfdReader().ReadAsync(s, l, t); });
        using var final = new MemoryStream(); var task = slow.ConvertAsync(new MemoryStream(source), final, selection); await entered.Task;
        selection[0] = 1; resume.SetResult(); await task; Assert.Equal("3", Parse(final.ToArray()).Descendants("section").First().Attribute("data-source-page")!.Value);
    }
    [Fact]
    public async Task BadLaterPageAndOpaqueDrawingFailBeforePublicWrite()
    {
        var source = await Fixture(); using var input = new MemoryStream(); input.Write(source);
        using (var zip = new ZipArchive(input, ZipArchiveMode.Update, true))
        {
            var entry = zip.GetEntry("Doc_0/Pages/Page_2/Content.xml")!; XDocument document; using (var stream = entry.Open()) document = XDocument.Load(stream);
            var layer = document.Descendants().First(n => n.Name.LocalName == "Layer"); layer.Add(new XElement(layer.Name.Namespace + "VendorDrawing", new XAttribute("ID", "999")));
            entry.Delete(); using var replacement = zip.CreateEntry("Doc_0/Pages/Page_2/Content.xml").Open(); document.Save(replacement);
        }
        input.Position = 0; using var output = new MemoryStream(); output.WriteByte(42);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(input, output)); Assert.Equal(new byte[] { 42 }, output.ToArray());
    }
    [Theory]
    [InlineData("onload", "alert(1)")][InlineData("href", "https://127.0.0.1:9/attack")][InlineData("clip-path", "url(#missing)")][InlineData("xml:base", "file:///tmp/")]
    public void GeneratedDomPolicyRejectsEventsExternalLinksAndUnclosedReferences(string name, string value)
    {
        var element = new XElement(Svg + "svg", new XAttribute(name == "xml:base" ? XNamespace.Xml + "base" : name, value));
        Assert.Throws<NotSupportedException>(() => SvgPreviewPolicy.Validate(element, "ofd-preview-svg-0-", "ofd-preview-font-0-", default));
    }
    [Fact]
    public void GeneratedDomPolicyRejectsDuplicateAndCrossPageIds()
    {
        var node = new XElement(Svg + "svg", new XElement(Svg + "clipPath", new XAttribute("id", "ofd-preview-svg-0-a")), new XElement(Svg + "clipPath", new XAttribute("id", "ofd-preview-svg-0-a")));
        Assert.Throws<NotSupportedException>(() => SvgPreviewPolicy.Validate(node, "ofd-preview-svg-0-", "ofd-preview-font-0-", default));
        node.Elements().Last().Remove(); node.Add(new XElement(Svg + "g", new XAttribute("clip-path", "url(#ofd-preview-svg-1-a)")));
        Assert.Throws<NotSupportedException>(() => SvgPreviewPolicy.Validate(node, "ofd-preview-svg-0-", "ofd-preview-font-0-", default));
    }
    [Theory]
    [InlineData("image/svg+xml")][InlineData("text/html")][InlineData("image/png")][InlineData("image/jpeg")]
    public void FakeMediaIsRejected(string mime) => Assert.ThrowsAny<Exception>(() => HtmlImageAdmission.Validate(System.Text.Encoding.UTF8.GetBytes("<svg onload='fetch(\"http://127.0.0.1:9\")'/>"), mime, default));
    [Fact]
    public void PngAndJpegRequireCompleteFraming()
    {
        using var image = new Image<Rgba32>(4, 4, Color.Red); using var png = new MemoryStream(); using var jpeg = new MemoryStream(); image.SaveAsPng(png); image.SaveAsJpeg(jpeg);
        HtmlImageAdmission.Validate(png.ToArray(), "image/png", default); HtmlImageAdmission.Validate(jpeg.ToArray(), "image/jpeg", default);
        Assert.Throws<InvalidDataException>(() => HtmlImageAdmission.Validate(png.ToArray()[..^1], "image/png", default));
        Assert.Throws<InvalidDataException>(() => HtmlImageAdmission.Validate(jpeg.ToArray().Concat(new byte[] { 1 }).ToArray(), "image/jpeg", default));
        var broken = png.ToArray(); broken[40] ^= 1; Assert.Throws<InvalidDataException>(() => HtmlImageAdmission.Validate(broken, "image/png", default));
    }
    [Theory]
    [InlineData(false)][InlineData(true)]
    public async Task ImagePixelBudgetPrecedesDecodeAndCountsDuplicateOccurrences(bool jpeg)
    {
        using var pixels = new Image<Rgba32>(64, 64, Color.Black); using var encoded = new MemoryStream();
        if (jpeg) pixels.SaveAsJpeg(encoded); else pixels.SaveAsPng(encoded);
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 120 }; package.Pages.Add(page);
        page.Elements.Add(new OfdImageElement { WidthMillimeters = 50, HeightMillimeters = 50, Data = encoded.ToArray(), MediaType = jpeg ? "image/jpeg" : "image/png" });
        using var source = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, source); source.Position = 0;
        using var output = new MemoryStream(); output.WriteByte(42);
        var options = new OfdHtmlOptions { MaxDecodedImagePixels = 8192 };
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToHtmlConverter(options).ConvertAsync(source, output, new[] { 0, 0, 0 }));
        Assert.Contains("decoded image pixels", error.Message); Assert.Equal(new byte[] { 42 }, output.ToArray());
        source.Position = 0; using var success = new MemoryStream(); await new OfdToHtmlConverter(options).ConvertAsync(source, success, new[] { 0, 0 }); Assert.True(success.Length > 0);
        // Budget callback runs at the encoded dimension gate, before decode work.
        using var cancel = new CancellationTokenSource(); long charged = 0;
        Assert.ThrowsAny<OperationCanceledException>(() => HtmlImageAdmission.Validate(encoded.ToArray(), jpeg ? "image/jpeg" : "image/png", cancel.Token, n => { charged += n; cancel.Cancel(); }));
        Assert.Equal(4096, charged);
    }
    [Fact]
    public void IndexedPngIsExplicitlyOutsideTheNarrowProfile()
    {
        using var pixels = new Image<Rgba32>(4, 4, Color.Red); using var encoded = new MemoryStream();
        pixels.Save(encoded, new SixLabors.ImageSharp.Formats.Png.PngEncoder { ColorType = SixLabors.ImageSharp.Formats.Png.PngColorType.Palette });
        Assert.Throws<NotSupportedException>(() => HtmlImageAdmission.Validate(encoded.ToArray(), "image/png", default));
    }
    [Fact]
    public async Task FinalCopyFailureAndCancellationDoNotPretendRollback()
    {
        var source = await Fixture(); using var fail = new FailingOutput();
        await Assert.ThrowsAsync<IOException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(source), fail)); Assert.True(fail.Length > 0);
        using var cancel = new CancellationTokenSource(); using var destination = new CancelOutput(cancel);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(source), destination, cancellationToken: cancel.Token)); Assert.True(destination.Length > 0);
    }
    [Fact]
    public async Task CancellationAfterReadLeavesOutputUnchanged()
    {
        var source = await Fixture(); using var cancel = new CancellationTokenSource();
        var converter = new OfdToHtmlConverter(new(), null, async (s, l, t) => { var result = await new OfdReader().ReadAsync(s, l, t); cancel.Cancel(); return result; });
        using var output = new MemoryStream(); output.WriteByte(42); await Assert.ThrowsAnyAsync<OperationCanceledException>(() => converter.ConvertAsync(new MemoryStream(source), output, cancellationToken: cancel.Token)); Assert.Equal(new byte[] { 42 }, output.ToArray());
    }
    private sealed class FailingOutput : MemoryStream
    { public override Task WriteAsync(byte[] b, int o, int c, CancellationToken token) { base.Write(b, o, Math.Min(10, c)); throw new IOException("controlled copy failure"); } }
    private sealed class CancelOutput(CancellationTokenSource cancel) : MemoryStream
    { public override Task WriteAsync(byte[] b, int o, int c, CancellationToken token) { base.Write(b, o, c); cancel.Cancel(); return Task.CompletedTask; } }
    private sealed class NonSeekable(byte[] data) : Stream
    {
        private readonly MemoryStream _inner = new(data); internal int ReadBytes; internal int Seeks;
        public override int Read(byte[] b, int o, int c) { var n = _inner.Read(b, o, c); ReadBytes += n; return n; }
        public override async Task<int> ReadAsync(byte[] b, int o, int c, CancellationToken token) { var n = await _inner.ReadAsync(b, o, c, token); ReadBytes += n; return n; }
        public override bool CanRead => true; public override bool CanSeek => false; public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException(); public override long Position { get => throw new NotSupportedException(); set { Seeks++; throw new NotSupportedException(); } }
        public override long Seek(long o, SeekOrigin s) { Seeks++; throw new NotSupportedException(); }
        public override void Flush() { } public override void SetLength(long v) => throw new NotSupportedException(); public override void Write(byte[] b, int o, int c) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
