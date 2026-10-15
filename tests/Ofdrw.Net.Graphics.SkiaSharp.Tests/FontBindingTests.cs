using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Graphics.SkiaSharp;
using SkiaSharp;
using Binding = Ofdrw.Net.Graphics.SkiaSharp.SkiaDrawEvent.FixedFontBinding;

namespace Ofdrw.Net.Graphics.SkiaSharp.Tests;

public sealed class FontBindingTests
{
    private static byte[] Bytes(string name = "narrow") => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", name + ".ttf"));
    private static (OfdDocumentPackage Package, OfdPage Page) Target(byte[] bytes, Binding binding, string id = "face")
    {
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        package.Fonts.Add(new OfdFontResource { Id = id, FontName = binding.FamilyName, Data = bytes, Bold = binding.Bold, Italic = binding.Italic });
        return (package, page);
    }
    private static SkiaDrawEvent Snapshot(Binding binding, SKPaint paint, CancellationToken token = default)
        => binding.SnapshotText("A  B", new(2, 3), 4, paint, new float[] { 2, 2, 2 }, SKMatrix.Identity, 100, token);

    [Theory]
    [InlineData(1)] [InlineData(10)] [InlineData(50)] [InlineData(100)]
    public void OwnedSnapshotSurvivesInputMutationAndBindingDisposal(int count)
    {
        var original = Bytes(); var caller = (byte[])original.Clone();
        using var binding = Binding.FromOriginalBytes(caller, "face", original.Length);
        var (package, page) = Target(original, binding);
        // Before first identity validation, overwrite every byte passed to the
        // factory. An identity based on a retained caller array would fail.
        Array.Clear(caller);
        using var paint = new SKPaint { Color = SKColors.Blue };
        var events = Enumerable.Range(0, count).Select(_ => Snapshot(binding, paint)).ToArray();
        binding.Dispose(); paint.Color = SKColors.Red;
        Assert.Throws<ObjectDisposedException>(() => Snapshot(binding, paint));
        Assert.Throws<ObjectDisposedException>(() => binding.GetGlyph('A'));
        OfdSkiaAdapter.Append(package, page, events);
        Assert.Equal(count, page.Elements.Count);
        Assert.All(page.Elements, value =>
        {
            var text = Assert.IsType<OfdTextElement>(value);
            Assert.Equal("A  B", text.Text); Assert.Equal("face", text.FontResourceId);
            Assert.Equal(255, text.FillColor.Blue);
        });
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void DestinationMutationBetweenAppendsIsReauthenticated(bool inPlace)
    {
        var bytes = Bytes(); using var binding = Binding.FromOriginalBytes(bytes, "face", bytes.Length);
        var (package, page) = Target(bytes, binding); using var paint = new SKPaint(); var value = Snapshot(binding, paint);
        OfdSkiaAdapter.Append(package, page, new[] { value }); var first = Assert.Single(page.Elements);
        if (inPlace) bytes[^1] ^= 1; else package.Fonts[0].Data = Bytes("wide");
        Assert.Throws<ArgumentException>(() => OfdSkiaAdapter.Append(package, page, new[] { value, value }));
        Assert.Same(first, Assert.Single(page.Elements));
    }

    [Fact]
    public void EqualIdsDoNotShareDifferentFontIdentityAndLateMismatchIsAtomic()
    {
        var bytes = Bytes(); using var first = Binding.FromOriginalBytes(bytes, "face", bytes.Length);
        using var second = Binding.FromOriginalBytes(Bytes("wide"), "face", 64000000);
        var (package, page) = Target(bytes, first); using var paint = new SKPaint();
        Assert.Throws<ArgumentException>(() => OfdSkiaAdapter.Append(package, page, new[] { Snapshot(first, paint), Snapshot(second, paint) }));
        Assert.Empty(page.Elements);
    }

    [Fact]
    public void EqualBytesWithDifferentIdsPreserveRouting()
    {
        var bytes = Bytes(); using var first = Binding.FromOriginalBytes(bytes, "a", bytes.Length);
        using var second = Binding.FromOriginalBytes(bytes, "b", bytes.Length);
        var (package, page) = Target(bytes, first, "a");
        package.Fonts.Add(new OfdFontResource { Id = "b", FontName = second.FamilyName, Data = (byte[])bytes.Clone(), Bold = second.Bold, Italic = second.Italic });
        using var paint = new SKPaint();
        OfdSkiaAdapter.Append(package, page, new[] { Snapshot(first, paint), Snapshot(second, paint), Snapshot(first, paint) });
        Assert.Equal(new[] { "a", "b", "a" }, page.Elements.Cast<OfdTextElement>().Select(x => x.FontResourceId));
    }

    [Theory]
    [InlineData("regular", false, false)] [InlineData("bold", true, false)] [InlineData("italic", false, true)]
    public void PhysicalStyleAndDestinationStyleRemainAuthenticated(string name, bool bold, bool italic)
    {
        var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "style-fonts", name + ".ttf"));
        using var binding = Binding.FromOriginalBytes(bytes, "face", bytes.Length); using var paint = new SKPaint();
        Assert.Equal(bold, binding.Bold); Assert.Equal(italic, binding.Italic);
        var (package, page) = Target(bytes, binding); var value = Snapshot(binding, paint);
        package.Fonts[0].Italic = !italic;
        Assert.Throws<ArgumentException>(() => OfdSkiaAdapter.Append(package, page, new[] { value })); Assert.Empty(page.Elements);
        package.Fonts[0].Italic = italic; OfdSkiaAdapter.Append(package, page, new[] { value }); Assert.Single(page.Elements);
    }

    [Theory]
    [InlineData("mismatch")] [InlineData("missing-os2")] [InlineData("short-os2")]
    public void InvalidStyleTablesNeverCreateAReusableBinding(string name)
    {
        var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "style-fonts", name + ".ttf"));
        Assert.Throws<NotSupportedException>(() => Binding.FromOriginalBytes(bytes, "face", bytes.Length));
    }

    [Fact]
    public void CancelledOrInvalidRunDoesNotPoisonLaterSnapshot()
    {
        var bytes = Bytes(); using var binding = Binding.FromOriginalBytes(bytes, "face", bytes.Length); using var paint = new SKPaint();
        Assert.Throws<OperationCanceledException>(() => Snapshot(binding, paint, new CancellationToken(true)));
        Assert.Throws<ArgumentException>(() => binding.SnapshotText("AA", new(1, 2), 4, paint, Array.Empty<float>(), SKMatrix.Identity, 100, default));
        Assert.Throws<System.Xml.XmlException>(() => binding.SnapshotText("\0", new(1, 2), 4, paint, Array.Empty<float>(), SKMatrix.Identity, 100, default));
        Assert.Throws<InvalidOperationException>(() => binding.SnapshotText("AAA", new(1, 2), 4, paint, new float[] { 1, 1 }, SKMatrix.Identity, 2, default));
        Assert.Throws<ArgumentException>(() => binding.SnapshotText("A", new(1, 2), 4, paint, Array.Empty<float>(), SKMatrix.CreateScale(0, 1), 100, default));
        paint.Style = SKPaintStyle.Stroke; Assert.Throws<NotSupportedException>(() => Snapshot(binding, paint)); paint.Style = SKPaintStyle.Fill;
        var (package, page) = Target(bytes, binding); OfdSkiaAdapter.Append(package, page, new[] { Snapshot(binding, paint) }); Assert.Single(page.Elements);
    }

    [Fact]
    public async Task ConcurrentSnapshotsAndDisposalAreSerializedWithoutNativeUseAfterFree()
    {
        var bytes = Bytes(); using var binding = Binding.FromOriginalBytes(bytes, "face", bytes.Length);
        var (package, page) = Target(bytes, binding);
        var tasks = Enumerable.Range(0, 20).Select(_ => Task.Run(() =>
        {
            using var paint = new SKPaint();
            try { return Snapshot(binding, paint); } catch (ObjectDisposedException) { return null; }
        })).ToArray();
        await Task.Run(binding.Dispose); var snapshots = (await Task.WhenAll(tasks)).Where(x => x is not null).Cast<SkiaDrawEvent>().ToArray();
        OfdSkiaAdapter.Append(package, page, snapshots); Assert.Equal(snapshots.Length, page.Elements.Count);
        Assert.Throws<ObjectDisposedException>(() => binding.GetGlyph('A'));
    }
}
