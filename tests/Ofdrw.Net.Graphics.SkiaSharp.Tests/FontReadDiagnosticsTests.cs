using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Graphics.SkiaSharp;
using Ofdrw.Net.TestSupport;
using SkiaSharp;
using Binding = Ofdrw.Net.Graphics.SkiaSharp.SkiaDrawEvent.FixedFontBinding;
using Access = Ofdrw.Net.Graphics.SkiaSharp.SkiaDrawEvent.NativeFontTableAccess;

namespace Ofdrw.Net.Graphics.SkiaSharp.Tests;

public sealed class FontReadDiagnosticsTests
{
    private static byte[] Bytes(string name = "narrow") => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", name + ".ttf"));
    private static JsonElement Evidence(Exception error) => JsonSerializer.SerializeToElement(
        error.Data[SkiaDrawEvent.FontReadFailureKey], error.Data[SkiaDrawEvent.FontReadFailureKey]!.GetType(),
        new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FirstShortSizeRemainsFailureWithoutASecondGoodQuery(bool recordHistory)
    {
        var bytes = Bytes(); var caller = bytes.ToArray(); var reads = 0; var creates = 0;
        var access = new Access(size: (face, tag, phase) =>
        {
            if (phase == "constructor.style.os2") return ++reads == 1 ? 0 : face.GetTableSize(tag);
            return face.GetTableSize(tag);
        }, beforeCreate: () => { creates++; Array.Clear(caller); }, recordHistory: recordHistory);
        var error = Assert.Throws<NotSupportedException>(() => Binding.FromOriginalBytes(caller, "face", bytes.Length, access));
        Assert.Equal("Font style table is missing or truncated.", error.Message);
        Assert.Equal(1, reads); Assert.Equal(1, creates);
        var data = Evidence(error);
        Assert.True(data.GetProperty("rawStylePreflightPassed").GetBoolean());
        Assert.False(data.GetProperty("decoratedAfterOriginalThrow").GetBoolean());
        Assert.Equal(Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(), data.GetProperty("originalSha256").GetString());
        var step = data.GetProperty("steps").EnumerateArray().Last();
        Assert.Equal("constructor.style.os2", step.GetProperty("phase").GetString());
        Assert.Equal(0, step.GetProperty("size").GetInt32());
        Assert.True(data.GetProperty("rawStyle").GetProperty("os2Length").GetInt32() >= 64);
        Assert.InRange(data.GetProperty("steps").GetArrayLength(), 1, 32);
        Assert.InRange(data.GetProperty("processHistory").GetArrayLength(), recordHistory ? 1 : 0, recordHistory ? 128 : 0);
    }

    [Fact]
    public void FirstByteFailureDoesNotReadTheNextByte()
    {
        var bytes = Bytes(); var offsets = new List<int>();
        var access = new Access(readByte: (face, tag, offset, buffer, phase) =>
        {
            offsets.Add(offset);
            return phase == "constructor.style.os2" ? false : face.TryGetTableData(tag, offset, 1, buffer);
        });
        var error = Assert.Throws<NotSupportedException>(() => Binding.FromOriginalBytes(bytes, "face", bytes.Length, access));
        Assert.Equal(new[] { 62 }, offsets);
        var step = Evidence(error).GetProperty("steps").EnumerateArray().Last();
        Assert.Equal("byte", step.GetProperty("operation").GetString());
        Assert.Equal(62, step.GetProperty("offset").GetInt32());
        Assert.False(step.GetProperty("success").GetBoolean());
        Assert.Equal(JsonValueKind.Null, step.GetProperty("value").ValueKind);
    }

    [Fact]
    public void NativeUpperBoundStillRejectsWithoutQueryingTheNextTable()
    {
        var bytes = Bytes(); var phases = new List<string>();
        var error = Assert.Throws<NotSupportedException>(() => Binding.FromOriginalBytes(bytes, "face", bytes.Length,
            new Access(size: (_, _, phase) => { phases.Add(phase); return 1025; })));
        Assert.Equal("PDFV_FONT_STYLE_TABLE", error.Message);
        Assert.Equal(new[] { "constructor.upper.os2" }, phases);
        Assert.Equal(1025, Evidence(error).GetProperty("steps").EnumerateArray().Last().GetProperty("size").GetInt32());
    }

    [Theory]
    [InlineData("missing-os2")]
    [InlineData("short-os2")]
    [InlineData("mismatch")]
    public void OriginalStyleRejectionPrecedesEveryNativeConstructionAndQuery(string name)
    {
        var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "style-fonts", name + ".ttf"));
        var creates = 0; var reads = 0;
        var access = new Access(size: (face, tag, _) => { reads++; return face.GetTableSize(tag); }, beforeCreate: () => creates++);
        var error = Assert.Throws<NotSupportedException>(() => Binding.FromOriginalBytes(bytes, "face", bytes.Length, access));
        Assert.Equal(0, creates); Assert.Equal(0, reads);
        var data = Evidence(error);
        Assert.Equal("raw-style-preflight-rejected", data.GetProperty("classification").GetString());
        Assert.False(data.GetProperty("rawStylePreflightPassed").GetBoolean());
        Assert.DoesNotContain(data.GetProperty("steps").EnumerateArray(), step => step.GetProperty("operation").GetString() == "native-create-begin");
    }

    [Fact]
    public void ValidRawNativeFailureDuringLazyBatchLeavesExistingDestinationAndResourcesUnchanged()
    {
        var firstBytes = Bytes(); var secondBytes = Bytes("wide");
        using var first = Binding.FromOriginalBytes(firstBytes, "first", firstBytes.Length);
        var reads = 0;
        using var second = Binding.FromOriginalBytes(secondBytes, "second", secondBytes.Length,
            new Access(size: (face, tag, phase) => phase == "identity.style.os2" ? (++reads == 1 ? 0 : face.GetTableSize(tag)) : face.GetTableSize(tag)));
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page); var sentinel = new OfdPathElement(); page.Elements.Add(sentinel);
        package.Fonts.Add(new OfdFontResource { Id = "first", FontName = first.FamilyName, Data = firstBytes, Bold = first.Bold, Italic = first.Italic });
        package.Fonts.Add(new OfdFontResource { Id = "second", FontName = second.FamilyName, Data = secondBytes, Bold = second.Bold, Italic = second.Italic });
        using var paint = new SKPaint();
        IEnumerable<SkiaDrawEvent> Events()
        {
            yield return first.SnapshotText("AA", new(2, 3), 4, paint, new float[] { 4 }, SKMatrix.Identity, 100, default);
            yield return second.SnapshotText("AA", new(2, 3), 4, paint, new float[] { 4 }, SKMatrix.Identity, 100, default);
        }
        var error = Assert.Throws<NotSupportedException>(() => OfdSkiaAdapter.Append(package, page, Events()));
        Assert.Equal(1, reads); Assert.Same(sentinel, Assert.Single(page.Elements));
        Assert.Same(firstBytes, package.Fonts[0].Data); Assert.Same(secondBytes, package.Fonts[1].Data);
        Assert.Equal("identity.style.os2", Evidence(error).GetProperty("steps").EnumerateArray().Last().GetProperty("phase").GetString());
    }

    [Fact]
    public void DirectReadExceptionIsNotReplacedAndIsExplicitlyDecoratedAfterItsOriginalThrow()
    {
        var bytes = Bytes(); var reads = 0; var original = new IOException("controlled table read exception");
        var access = new Access(size: (face, tag, phase) =>
        {
            if (phase == "constructor.style.os2") { reads++; throw original; }
            return face.GetTableSize(tag);
        });
        var error = Assert.Throws<IOException>(() => Binding.FromOriginalBytes(bytes, "face", bytes.Length, access));
        Assert.Same(original, error); Assert.Equal(1, reads);
        var data = Evidence(error); Assert.True(data.GetProperty("decoratedAfterOriginalThrow").GetBoolean());
        var step = data.GetProperty("steps").EnumerateArray().Last();
        Assert.Equal("size-threw", step.GetProperty("operation").GetString());
        Assert.Equal(JsonValueKind.Null, step.GetProperty("size").ValueKind);
    }

    [Fact]
    public void CollectorCapturesAnAlreadyLoadedNativeModuleAndCannotMaskFailureWhenItsSinkFails()
    {
        var bytes = Bytes();
        var error = Assert.Throws<NotSupportedException>(() => Binding.FromOriginalBytes(bytes, "face", bytes.Length,
            new Access(size: (face, tag, phase) => phase == "constructor.style.os2" ? 0 : face.GetTableSize(tag))));
        var evidence = error.Data[SkiaDrawEvent.FontReadFailureKey];
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-skia-diagnostic-" + Guid.NewGuid().ToString("N"));
        var blocked = Path.GetTempFileName();
        try
        {
            var path = SkiaFontDiagnosticHost.Capture(error, directory); Assert.NotNull(path);
            using var document = JsonDocument.Parse(File.ReadAllBytes(path));
            var root = document.RootElement;
            Assert.Equal(Environment.ProcessId, root.GetProperty("processId").GetInt32());
            Assert.Equal("loaded-skia-module-captured", root.GetProperty("native").GetProperty("status").GetString());
            Assert.Contains(root.GetProperty("native").GetProperty("modules").EnumerateArray(), module =>
                Path.GetFileName(module.GetProperty("path").GetString()!).StartsWith("libSkiaSharp", StringComparison.OrdinalIgnoreCase) &&
                module.GetProperty("sha256").GetString() is { Length: 64 });
            Assert.Null(SkiaFontDiagnosticHost.Capture(error, blocked));
            Assert.Same(evidence, error.Data[SkiaDrawEvent.FontReadFailureKey]);
            Assert.Equal("Font style table is missing or truncated.", error.Message);
        }
        finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); File.Delete(blocked); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DiagnosticFailuresCannotMaskAdmissionFailureOrSkipNativeCleanup(bool failSnapshot)
    {
        var bytes = Bytes(); SKTypeface? captured = null;
        var access = new Access(size: (face, tag, phase) =>
        {
            captured = face;
            return phase == "constructor.style.os2" ? 0 : face.GetTableSize(tag);
        }, beforeDiagnosticRecord: () => throw new IOException("controlled record failure"),
            beforeDiagnosticSnapshot: failSnapshot ? () => throw new IOException("controlled snapshot failure") : null);
        var error = Assert.Throws<NotSupportedException>(() => Binding.FromOriginalBytes(bytes, "face", bytes.Length, access));
        Assert.Equal("Font style table is missing or truncated.", error.Message);
        Assert.NotNull(captured); Assert.Equal(IntPtr.Zero, captured.Handle);
        if (failSnapshot) Assert.False(error.Data.Contains(SkiaDrawEvent.FontReadFailureKey));
        else Assert.True(Evidence(error).GetProperty("recordingFailed").GetBoolean());
    }

    [Fact]
    public void FailedOptionalRecordingDoesNotChangeSuccessfulIdentityOrDisposal()
    {
        var bytes = Bytes(); SKTypeface? captured = null;
        var access = new Access(size: (face, tag, _) => { captured = face; return face.GetTableSize(tag); },
            beforeDiagnosticRecord: () => throw new IOException("controlled record failure"));
        using var binding = Binding.FromOriginalBytes(bytes, "face", bytes.Length, access);
        using var paint = new SKPaint();
        var value = binding.SnapshotText("AA", new(2, 3), 4, paint, new float[] { 4 }, SKMatrix.Identity, 100, default);
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page); package.Fonts.Add(new OfdFontResource
        { Id = "face", FontName = binding.FamilyName, Data = bytes, Bold = binding.Bold, Italic = binding.Italic });
        OfdSkiaAdapter.Append(package, page, new[] { value });
        Assert.Equal("AA", Assert.IsType<OfdTextElement>(Assert.Single(page.Elements)).Text);
        binding.Dispose(); Assert.NotNull(captured); Assert.Equal(IntPtr.Zero, captured.Handle);
    }

    [Fact]
    public void CollectorDeduplicationAndBothReportScopesRemainBoundedUnderConcurrentPressure()
    {
        var budget = new SkiaFontDiagnosticHost.ReportBudget();
        var accepted = new System.Collections.Concurrent.ConcurrentBag<(string Id, bool Controlled)>();
        Parallel.For(0, 1000, index =>
        {
            foreach (var controlled in new[] { false, true })
            {
                var id = controlled + ":" + index;
                if (budget.TryReserve(id, controlled)) accepted.Add((id, controlled));
            }
        });
        Assert.Equal(32, accepted.Count); Assert.Equal((16, 16, 32), budget.Counts);
        foreach (var value in accepted) Assert.False(budget.TryReserve(value.Id, value.Controlled));
        Assert.Equal((16, 16, 32), budget.Counts);
    }

    [Fact]
    public void CollectorCannotLetADataAccessorFailureEscape()
    {
        var error = new BrokenDataException();
        Assert.Null(SkiaFontDiagnosticHost.Capture(error, Path.GetTempPath()));
    }
    private sealed class BrokenDataException : Exception
    { public override System.Collections.IDictionary Data => throw new IOException("controlled Data accessor failure"); }
}
