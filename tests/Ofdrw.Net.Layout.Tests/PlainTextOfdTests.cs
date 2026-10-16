using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class PlainTextOfdTests
{
    [Theory]
    [InlineData(1)] [InlineData(2)] [InlineData(4096)]
    public async Task Utf8BomAndChunkBoundaries(int chunk)
    {
        const string value = "中英 Alpha e\u0301 😀\U00020000\uFEFF end";
        foreach (var prefix in new[] { "", "\uFEFF", "\uFEFF\uFEFF" })
        {
            using var input = new ChunkStream(Encoding.UTF8.GetBytes(prefix + value), chunk);
            var package = await new PlainTextOfdConverter().RenderUtf8Async(input);
            Assert.Equal((prefix.Length == 2 ? "\uFEFF" : "") + value, Text(package));
            Assert.True(input.CanRead);
        }
        Assert.Equal("\uFEFFX", Text(new PlainTextOfdConverter().Render("\uFEFFX")));
    }

    [Theory]
    [InlineData("")] [InlineData("\n\r\n\r")] [InlineData("\f\f")]
    public async Task EmptyAndControlOnlyInputUsesOneBlankPage(string text)
    {
        var converter = new PlainTextOfdConverter();
        Assert.Empty(Assert.Single(converter.Render(text).Pages).Elements);
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("\uFEFF" + text));
        Assert.Empty(Assert.Single((await converter.RenderUtf8Async(input)).Pages).Elements);
    }

    [Theory]
    [InlineData("A\r\nB", 2)] [InlineData("A\rB", 2)] [InlineData("A\nB", 2)]
    [InlineData("A\n\nB", 3)] [InlineData("A\n", 1)] [InlineData("A\n\n", 1)]
    [InlineData("A\fB", 2)] [InlineData("A\f\fB", 3)] [InlineData("A\f", 1)]
    public void InteriorAndTerminalBreaksUseFlowPagination(string text, int pages)
    {
        var package = new PlainTextOfdConverter(Small()).Render(text);
        Assert.Equal(pages, package.Pages.Count);
        Assert.Equal(text.Replace("\r", "").Replace("\n", "").Replace("\f", ""), Text(package));
        foreach (var element in package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>())
            Assert.Equal(5, element.YMillimeters, 5);
        if (pages == 3) Assert.Empty(package.Pages[1].Elements);
    }

    [Fact]
    public async Task WhitespaceAndGraphemesRoundTripAsNativeTextWithoutNormalization()
    {
        const string value = "  中英 iii WWW\tX\u00A0Y e\u0301 😀\U00020000\u2028\u2029  ";
        var package = new PlainTextOfdConverter().Render(value);
        Assert.Equal(value, Text(package));
        Assert.All(package.Pages.SelectMany(p => p.Elements), e => Assert.IsType<OfdTextElement>(e));
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0;
        var parsed = await new OfdReader().ReadAsync(output);
        Assert.Equal(value, Text(parsed));
        Assert.All(package.Fonts, f => Assert.Empty(f.Data));
    }

    [Fact]
    public void LongLineMatchesOneDirectFlowParagraph()
    {
        var text = string.Concat(Enumerable.Repeat("中英 iii WWW proportional ", 40)) + new string('W', 200);
        var options = Small();
        options.Layout.PageWidthMillimeters = 80;
        var direct = new FlowDocument();
        direct.Options.PageWidthMillimeters = 80; direct.Options.PageHeightMillimeters = 30;
        direct.Options.MarginLeftMillimeters = direct.Options.MarginRightMillimeters = 5;
        direct.Options.MarginTopMillimeters = direct.Options.MarginBottomMillimeters = 5;
        direct.Options.DefaultFontSizeMillimeters = 10;
        direct.Blocks.Add(new Paragraph(text));
        var expected = direct.Render();
        var actual = new PlainTextOfdConverter(options).Render(text);
        Assert.True(actual.Pages.Count > 2);
        Assert.Equal(text, Text(actual));
        Assert.Equal(expected.Pages.Count, actual.Pages.Count);
        Assert.Equal(expected.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(e => (e.Text, e.XMillimeters, e.YMillimeters)),
            actual.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(e => (e.Text, e.XMillimeters, e.YMillimeters)));
    }

    [Theory]
    [InlineData(new byte[] { 0xC3, 0x28 })]
    [InlineData(new byte[] { 0xE4, 0xB8 })]
    [InlineData(new byte[] { 0xED, 0xA0, 0x80 })]
    [InlineData(new byte[] { 0xC0, 0xAF })]
    [InlineData(new byte[] { 0xFF, 0xFE })]
    public async Task InvalidUtf8FailsWithoutReplacement(byte[] bytes)
    {
        using var input = new ChunkStream(bytes, 1);
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new PlainTextOfdConverter().RenderUtf8Async(input));
        Assert.IsType<DecoderFallbackException>(error.InnerException);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData("\0")] [InlineData("\v")] [InlineData("\uFFFE")] [InlineData("\uFFFF")]
    public async Task XmlInvalidScalarsFailAtBothEntrypoints(string text)
    {
        Assert.Throws<InvalidDataException>(() => new PlainTextOfdConverter().Render(text));
        using var input = new ChunkStream(Encoding.UTF8.GetBytes(text), 1);
        await Assert.ThrowsAsync<InvalidDataException>(() => new PlainTextOfdConverter().RenderUtf8Async(input));
    }

    [Fact]
    public void UnpairedSurrogatesFail()
    {
        foreach (var value in new[] { "\uD800", "\uDC00", "\uD800X", "\uD800\uD800" })
            Assert.Throws<InvalidDataException>(() => new PlainTextOfdConverter().Render(value));
    }

    [Fact]
    public async Task InputBudgetsCountBomCrLfAndSurrogateCodeUnitsBeforeLayout()
    {
        var options = new PlainTextOfdOptions { MaxInputBytes = 4 };
        options.Layout.MaxCharacters = 2;
        var converter = new PlainTextOfdConverter(options);
        using var exact = new ChunkStream(Encoding.UTF8.GetBytes("😀"), 1);
        Assert.Equal("😀", Text(await converter.RenderUtf8Async(exact)));
        using var over = new ChunkStream(Encoding.UTF8.GetBytes("\uFEFF中"), 4096);
        await Assert.ThrowsAsync<InvalidDataException>(() => converter.RenderUtf8Async(over));
        Assert.Equal(5, over.BytesRead); // limit + exactly one byte
        Assert.Equal("AB", Text(converter.Render("AB")));
        Assert.Throws<InvalidDataException>(() => converter.Render("A\r\n"));
        Assert.Throws<InvalidDataException>(() => converter.Render("😀A"));
        using var chars = new ChunkStream(Encoding.UTF8.GetBytes("ABC"), 1);
        await Assert.ThrowsAsync<InvalidDataException>(() => converter.RenderUtf8Async(chars));
        Assert.Equal(3, chars.BytesRead);
    }

    [Theory]
    [InlineData("pages")] [InlineData("elements")] [InlineData("height")] [InlineData("width")]
    public void OutputLimitsAndUnfittableTextRetainFlowError(string kind)
    {
        var options = Small();
        if (kind == "pages") options.Layout.MaxPageCount = 1;
        if (kind == "elements") options.Layout.MaxTextElements = 1;
        if (kind == "height") options.Layout.PageHeightMillimeters = 11;
        if (kind == "width") options.Layout.PageWidthMillimeters = 11;
        var error = Assert.Throws<InvalidDataException>(() => new PlainTextOfdConverter(options).Render("WW\nW"));
        Assert.IsType<InvalidOperationException>(error.InnerException);
    }

    [Fact]
    public async Task OptionsAreValidatedAtConstructionAndCallsAreIndependent()
    {
        var options = Small();
        var converter = new PlainTextOfdConverter(options);
        options.Layout.PageWidthMillimeters = double.NaN;
        options.Layout.MaxPageCount = 0;
        options.MaxInputBytes = 0;
        var packages = await Task.WhenAll(Enumerable.Range(0, 8).Select(i => Task.Run(() => converter.Render("A\nB"))));
        Assert.All(packages, p => Assert.Equal(2, p.Pages.Count));
        packages[0].Pages.Clear(); Assert.Equal(2, packages[1].Pages.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PlainTextOfdConverter(options));
        options = Small(); options.Layout.DefaultColor = null!;
        Assert.Throws<ArgumentException>(() => new PlainTextOfdConverter(options));
        options = Small(); options.Layout.MarginLeftMillimeters = 40;
        Assert.Throws<ArgumentException>(() => new PlainTextOfdConverter(options));
        options = Small(); options.Layout.DefaultFontSizeMillimeters = double.PositiveInfinity;
        Assert.Throws<ArgumentOutOfRangeException>(() => new PlainTextOfdConverter(options));
    }

    [Fact]
    public async Task CancellationAndIoErrorsPropagateWithoutClosingBorrowedStream()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var converter = new PlainTextOfdConverter();
        Assert.ThrowsAny<OperationCanceledException>(() => converter.Render("A", cancellation.Token));
        using var before = new ChunkStream(new byte[] { 65 }, 1);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => converter.RenderUtf8Async(before, cancellation.Token));
        Assert.Equal(0, before.BytesRead);
        using var during = new ChunkStream(new byte[] { 65, 66 }, 1) { AfterRead = () => cancellation.Cancel() };
        using var later = new CancellationTokenSource();
        during.AfterRead = () => later.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => converter.RenderUtf8Async(during, later.Token));
        Assert.Equal(1, during.BytesRead); Assert.True(during.CanRead);
        using var io = new ChunkStream(Array.Empty<byte>(), 1) { AfterRead = () => throw new IOException("read failure") };
        await Assert.ThrowsAsync<IOException>(() => converter.RenderUtf8Async(io)); Assert.True(io.CanRead);
    }

    [Fact]
    public async Task CurrentPositionAndArgumentContracts()
    {
        var converter = new PlainTextOfdConverter();
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("skip\uFEFF中")); input.Position = 4;
        Assert.Equal("中", Text(await converter.RenderUtf8Async(input))); Assert.True(input.CanRead);
        Assert.Throws<ArgumentNullException>(() => converter.Render(null!));
        Assert.Throws<ArgumentNullException>(() => new PlainTextOfdConverter(null!));
        await Assert.ThrowsAsync<ArgumentNullException>(() => converter.RenderUtf8Async(null!));
        using var closed = new MemoryStream(); closed.Close();
        await Assert.ThrowsAsync<ArgumentException>(() => converter.RenderUtf8Async(closed));
    }

    private static PlainTextOfdOptions Small()
    {
        var options = new PlainTextOfdOptions(); var layout = options.Layout;
        layout.PageWidthMillimeters = 40; layout.PageHeightMillimeters = 30;
        layout.MarginLeftMillimeters = layout.MarginRightMillimeters = layout.MarginTopMillimeters = layout.MarginBottomMillimeters = 5;
        layout.DefaultFontSizeMillimeters = 10;
        return options;
    }
    private static string Text(OfdDocumentPackage package) => string.Concat(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(e => e.Text));

    private sealed class ChunkStream : Stream
    {
        private readonly byte[] _bytes; private readonly int _chunk; private bool _disposed;
        internal ChunkStream(byte[] bytes, int chunk) { _bytes = bytes; _chunk = chunk; }
        internal int BytesRead { get; private set; }
        internal Action? AfterRead { get; set; }
        public override bool CanRead => !_disposed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        { token.ThrowIfCancellationRequested(); var read = Read(buffer, offset, count); AfterRead?.Invoke(); return Task.FromResult(read); }
        public override int Read(byte[] buffer, int offset, int count)
        { var read = Math.Min(Math.Min(_chunk, count), _bytes.Length - BytesRead); Array.Copy(_bytes, BytesRead, buffer, offset, read); BytesRead += read; return read; }
        protected override void Dispose(bool disposing) { _disposed = true; base.Dispose(disposing); }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
