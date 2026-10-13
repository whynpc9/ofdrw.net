using System.Security.Cryptography;
using Ofdrw.Net.Converter.Pdf.Geometry;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[CollectionDefinition("Trusted font geometry", DisableParallelization = true)]
public sealed class TrustedFontGeometryCollection;

[Collection("Trusted font geometry")]
public sealed class TrustedFontGeometryTests
{
    private readonly Xunit.Abstractions.ITestOutputHelper output;
    public TrustedFontGeometryTests(Xunit.Abstractions.ITestOutputHelper output) => this.output = output;
    [Fact]
    public void ActualReaderSnapshotsRetainIdentityAcrossOneTenAndHundredRuns()
    {
        foreach (var runs in new[] { 1, 10, 100 })
        {
            var original = Fixture("narrow"); var font = Resource("10", original);
            using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(font, "10");
            var provider = new OfdTrustedEmbeddedFontGeometryProvider([Identity(original)]);
            var session = (OfdTrustedEmbeddedFontGeometryProvider.Session)operation.OpenSession(provider);
            var request = operation.Request(snapshot, "A中文 ", 4, 400, false);
            for (var i = 0; i < runs; i++)
            {
                var result = session.Measure(request); Assert.True(result.IsSupported, result.UnsupportedReason);
                Assert.Equal(4, result.Clusters.Count);
                AssertBox(result.Clusters[0].FontBoundsMillimeters!, 0.16, -2.8, 1.84, 0);
                Assert.Null(result.Clusters[3].FontBoundsMillimeters); Assert.Equal(2, result.Clusters[3].AdvanceXMillimeters);
            }
            Assert.Equal(1, operation.HashOperations); Assert.Equal(original.Length, operation.CopiedFontBytes);
            Assert.Equal(1, session.Statistics.FaceParses); Assert.Equal(4 * runs, session.Statistics.GlyphRequests);
            Assert.Equal(4, session.BackendGlyphCalls); Assert.Equal(0, session.Statistics.OutlinePointVisits);
            operation.ValidateSourceFontsUnchanged(); operation.Dispose();
            Assert.Equal(1, session.StreamsOpened); Assert.Equal(1, session.StreamsDisposed); Assert.Equal(0, session.BoundFaces);
            Assert.Throws<ObjectDisposedException>(() => session.Measure(request));
            Assert.Equal(Hash(original), snapshot.ContentSha256);
        }
    }

    [Fact]
    public void SameFamilyNamesCannotMixDistinctActualResourceFaces()
    {
        var narrow = Fixture("narrow"); var wide = Fixture("wide");
        using var operation = new OfdKeywordOperation();
        var first = operation.CaptureFont(Resource("A", narrow), "A"); var second = operation.CaptureFont(Resource("B", wide), "B");
        var session = operation.OpenSession(new OfdTrustedEmbeddedFontGeometryProvider([Identity(narrow), Identity(wide)]));
        var a = session.Measure(operation.Request(first, "A", 4, 400, false));
        var b = session.Measure(operation.Request(second, "A", 4, 400, false));
        var again = session.Measure(operation.Request(first, "A", 4, 400, false));
        Assert.True(a.IsSupported, a.UnsupportedReason); Assert.True(b.IsSupported, b.UnsupportedReason);
        Assert.Equal(2, a.Clusters[0].AdvanceXMillimeters); Assert.Equal(3.4, b.Clusters[0].AdvanceXMillimeters);
        AssertBox(again.Clusters[0].FontBoundsMillimeters!, 0.16, -2.8, 1.84, 0);
        Assert.Equal(2, session.Statistics.FaceParses); Assert.Equal(2, operation.HashOperations);
    }

    [Fact]
    public void ConstructorFreezesConfigurationAndRejectsConflictingOrMalformedIdentities()
    {
        var bytes = Fixture("narrow"); var profiles = new List<OfdTrustedEmbeddedFontIdentity> { Identity(bytes) };
        var provider = new OfdTrustedEmbeddedFontGeometryProvider(profiles); profiles.Clear();
        using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(Resource("10", bytes), "10");
        Assert.True(operation.OpenSession(provider).Measure(operation.Request(snapshot, "A", 4, 400, false)).IsSupported);
        Assert.Throws<ArgumentException>(() => new OfdTrustedEmbeddedFontGeometryProvider([Identity(bytes), new(Hash(bytes), bytes.Length, bold: true)]));
        Assert.Throws<ArgumentException>(() => new OfdTrustedEmbeddedFontIdentity(new string('z', 64), bytes.Length));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdTrustedEmbeddedFontGeometryProvider(Enumerable.Repeat(Identity(bytes), 65).ToArray()));
    }

    [Fact]
    public void ChangingConfigurationCountCannotExpandTheAcceptedCopy()
    {
        var bytes = Fixture("narrow"); var configuration = new GrowingConfiguration(Identity(bytes));
        var provider = new OfdTrustedEmbeddedFontGeometryProvider(configuration);
        Assert.Equal(1, configuration.EntriesRead);
        using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(Resource("A", bytes), "A");
        Assert.True(operation.OpenSession(provider).Measure(operation.Request(snapshot, "A", 4, 400, false)).IsSupported);
    }

    [Theory]
    [InlineData("unknown")]
    [InlineData("length")]
    [InlineData("profile-bold")]
    [InlineData("resource-bold")]
    [InlineData("resource-italic")]
    [InlineData("italic")]
    [InlineData("weight100")]
    [InlineData("weight500")]
    [InlineData("weight700")]
    public void UnknownIdentityOrStyleNeverCallsBackendAdd(string refusal)
    {
        var bytes = Fixture("narrow"); var font = Resource("10", bytes);
        var profiles = new[] { refusal == "unknown" ? Identity(Fixture("wide")) :
            new OfdTrustedEmbeddedFontIdentity(Hash(bytes), bytes.Length + (refusal == "length" ? 1 : 0), bold: refusal == "profile-bold") };
        font.Bold = refusal == "resource-bold"; font.Italic = refusal == "resource-italic";
        using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(font, "10");
        var session = (OfdTrustedEmbeddedFontGeometryProvider.Session)operation.OpenSession(new OfdTrustedEmbeddedFontGeometryProvider(profiles));
        var weight = refusal.StartsWith("weight", StringComparison.Ordinal) ? int.Parse(refusal[6..]) : 400;
        Assert.False(session.Measure(operation.Request(snapshot, "A", 4, weight, refusal == "italic")).IsSupported);
        Assert.Equal(0, session.BackendAddAttempts); Assert.Equal(0, session.StreamsOpened);
    }

    [Fact]
    public void ForeignReaderScopeIsRejectedBeforeAnyBackendBinding()
    {
        var bytes = Fixture("narrow"); using var first = new OfdKeywordOperation(); using var second = new OfdKeywordOperation();
        var snapshot = first.CaptureFont(Resource("A", bytes), "A");
        var session = (OfdTrustedEmbeddedFontGeometryProvider.Session)second.OpenSession(new OfdTrustedEmbeddedFontGeometryProvider([Identity(bytes)]));
        Assert.Throws<ArgumentException>(() => second.Request(snapshot, "A", 4, 400, false));
        Assert.Equal(0, session.BackendAddAttempts);
    }

    [Fact]
    public void MissingGlyphAndComplexTextAreExplicitlyUnsupported()
    {
        var bytes = Fixture("narrow"); using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(Resource("A", bytes), "A");
        var session = operation.OpenSession(new OfdTrustedEmbeddedFontGeometryProvider([Identity(bytes)]));
        Assert.Equal("MissingGlyph", session.Measure(operation.Request(snapshot, "Ā", 4, 400, false)).UnsupportedReason);
        Assert.Equal("ContextualOrComplexGlyphMappingIsUnsupported", session.Measure(operation.Request(snapshot, "𝟘", 4, 400, false)).UnsupportedReason);
        Assert.Equal("ContextualOrComplexGlyphMappingIsUnsupported", session.Measure(operation.Request(snapshot, "A\u0301", 4, 400, false)).UnsupportedReason);
        Assert.Equal("ContextualOrComplexGlyphMappingIsUnsupported", session.Measure(operation.Request(snapshot, "ا", 4, 400, false)).UnsupportedReason);
        Assert.Throws<InvalidDataException>(() => session.Measure(operation.Request(snapshot, new string('\uD800', 1), 4, 400, false)));
    }

    [Theory]
    [InlineData("ttc")]
    [InlineData("cff")]
    [InlineData("variable")]
    [InlineData("actual-weight")]
    [InlineData("actual-oblique")]
    public void UnsupportedTrustedFormatsOrActualStylesAreRejectedBeforeAdd(string kind)
    {
        var bytes = kind == "ttc" ? File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "style-collection.ttc")) : Fixture("narrow");
        if (kind == "cff") System.Text.Encoding.ASCII.GetBytes("OTTO").CopyTo(bytes, 0);
        if (kind is "variable" or "actual-weight" or "actual-oblique")
        {
            var count = (bytes[4] << 8) | bytes[5];
            for (var i = 0; i < count; i++)
            {
                var record = 12 + i * 16; var tag = System.Text.Encoding.ASCII.GetString(bytes, record, 4);
                if (kind == "variable" && tag == "glyf") System.Text.Encoding.ASCII.GetBytes("fvar").CopyTo(bytes, record);
                if (kind == "actual-weight" && tag == "OS/2")
                {
                    var offset = (int)System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 8, 4));
                    bytes[offset + 4] = 2; bytes[offset + 5] = 188; // 700, not the admitted actual Regular400 face.
                }
                if (kind == "actual-oblique" && tag == "OS/2")
                {
                    var offset = (int)System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 8, 4));
                    bytes[offset] = 0; bytes[offset + 1] = 4;
                    bytes[offset + 62] = 2; bytes[offset + 63] = 0; // v4 OBLIQUE alone; ITALIC/BOLD/macStyle remain clear.
                }
            }
        }
        using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(Resource("A", bytes), "A");
        var session = (OfdTrustedEmbeddedFontGeometryProvider.Session)operation.OpenSession(new OfdTrustedEmbeddedFontGeometryProvider([Identity(bytes)]));
        Assert.False(session.Measure(operation.Request(snapshot, "A", 4, 400, false)).IsSupported);
        Assert.False(session.Measure(operation.Request(snapshot, "A", 4, 400, false)).IsSupported);
        Assert.Equal(0, session.BackendAddAttempts); Assert.Equal(1, session.StreamsOpened); Assert.Equal(1, session.StreamsDisposed);
    }

    [Fact]
    public void CallerFaceLimitAndHardFourFaceCeilingAreBothEnforcedAndReleaseFaces()
    {
        var bytes = Fixture("narrow"); var provider = new OfdTrustedEmbeddedFontGeometryProvider([Identity(bytes)]);
        foreach (var cap in new[] { 1, 4 })
        {
            using var operation = new OfdKeywordOperation();
            using var generic = provider.OpenSession(new(maxFontFaces: cap == 1 ? 1 : 64));
            var session = (OfdTrustedEmbeddedFontGeometryProvider.Session)generic;
            for (var i = 0; i < cap; i++)
            {
                var snapshot = operation.CaptureFont(Resource("F" + i, bytes), "F" + i);
                Assert.True(session.Measure(operation.Request(snapshot, "A", 4, 400, false)).IsSupported);
            }
            var excess = operation.CaptureFont(Resource("excess", bytes), "excess");
            Assert.Throws<InvalidDataException>(() => session.Measure(operation.Request(excess, "A", 4, 400, false)));
            Assert.Equal(cap, session.BackendAddAttempts); Assert.Equal(cap, session.StreamsDisposed); Assert.Equal(0, session.BoundFaces);
        }
    }

    [Fact]
    public void SourceByteAndGlyphWorkLimitsFailBeforePublicationAndReleaseResources()
    {
        var bytes = Fixture("narrow"); using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(Resource("A", bytes), "A");
        var provider = new OfdTrustedEmbeddedFontGeometryProvider([Identity(bytes)]);
        using (var small = provider.OpenSession(new(maxFontBytes: bytes.Length - 1)))
        {
            Assert.Throws<InvalidDataException>(() => small.Measure(operation.Request(snapshot, "A", 4, 400, false)));
            Assert.Equal(0, small.Statistics.FaceParses);
        }
        using var limited = provider.OpenSession(new(maxGlyphOperations: 1));
        Assert.True(limited.Measure(operation.Request(snapshot, "A", 4, 400, false)).IsSupported);
        Assert.Throws<InvalidDataException>(() => limited.Measure(operation.Request(snapshot, "B", 4, 400, false)));
        var concrete = (OfdTrustedEmbeddedFontGeometryProvider.Session)limited;
        Assert.Equal(1, concrete.StreamsDisposed); Assert.Equal(0, concrete.CachedGlyphs);
    }

    [Theory]
    [InlineData("before-add")]
    [InlineData("after-add")]
    [InlineData("failure-after-add")]
    [InlineData("during-next-measure")]
    public void CancellationAndFaultsReleaseActualSnapshotStreams(string failure)
    {
        var bytes = Fixture("narrow"); using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(Resource("A", bytes), "A");
        using var cancellation = new CancellationTokenSource();
        var session = (OfdTrustedEmbeddedFontGeometryProvider.Session)operation.OpenSession(new OfdTrustedEmbeddedFontGeometryProvider([Identity(bytes)]));
        if (failure == "before-add") session.BeforeBackendAdd = cancellation.Cancel;
        if (failure == "after-add") session.AfterBackendAdd = cancellation.Cancel;
        if (failure == "failure-after-add") session.AfterBackendAdd = () => throw new IOException("Injected application failure after Add.");
        if (failure == "during-next-measure") { Assert.True(session.Measure(operation.Request(snapshot, "A", 4, 400, false)).IsSupported); cancellation.Cancel(); }
        var error = Record.Exception(() => session.Measure(operation.Request(snapshot, "A", 4, 400, false), cancellation.Token));
        if (failure == "failure-after-add") Assert.IsType<IOException>(error); else Assert.IsType<OperationCanceledException>(error);
        Assert.Equal(failure == "before-add" ? 0 : 1, session.BackendAddAttempts);
        Assert.Equal(1, session.StreamsOpened); Assert.Equal(1, session.StreamsDisposed); Assert.Equal(0, session.BoundFaces);
        Assert.Throws<ObjectDisposedException>(() => session.Measure(operation.Request(snapshot, "A", 4, 400, false)));
    }

    [Fact]
    public void FailureAfterSecondAddReleasesBothExistingFaceAndUnboundNewStream()
    {
        var bytes = Fixture("narrow"); using var operation = new OfdKeywordOperation();
        var first = operation.CaptureFont(Resource("A", bytes), "A"); var second = operation.CaptureFont(Resource("B", bytes), "B");
        var session = (OfdTrustedEmbeddedFontGeometryProvider.Session)operation.OpenSession(new OfdTrustedEmbeddedFontGeometryProvider([Identity(bytes)]));
        Assert.True(session.Measure(operation.Request(first, "A", 4, 400, false)).IsSupported);
        session.AfterBackendAdd = () => throw new IOException("Second Add completed; application fault before registration.");
        Assert.Throws<IOException>(() => session.Measure(operation.Request(second, "A", 4, 400, false)));
        Assert.Equal(2, session.BackendAddAttempts); Assert.Equal(2, session.StreamsOpened); Assert.Equal(2, session.StreamsDisposed);
        Assert.Equal(0, session.BoundFaces); Assert.Equal(0, session.CachedGlyphs);
    }

    [RealTrustedFontsTheory]
    [InlineData("cjk", 1)]
    [InlineData("cjk", 10)]
    [InlineData("cjk", 100)]
    [InlineData("latin", 1)]
    [InlineData("latin", 10)]
    [InlineData("latin", 100)]
    public void RealOriginalOflNotoBytesPassReaderSnapshotAndIndependentMetricExpectations(string face, int runs)
    {
        var path = Environment.GetEnvironmentVariable(face == "cjk" ? "OFDRW_TEST_TRUSTED_CJK_FONT" : "OFDRW_TEST_TRUSTED_LATIN_FONT")!;
        var bytes = File.ReadAllBytes(path);
        var expected = face == "cjk" ? "3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa" : "b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5";
        Assert.Equal(expected, Hash(bytes));
        // Fixed OFL sources and independent raw cmap/glyf/hmtx oracle are recorded in Issue06 evidence.
        // CJK: Noto revision f8d157... fixed Regular fixture; Latin: noto-fonts ffebf8... original Regular.
        using var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(Resource("actual-font", bytes), "actual-font");
        var session = (OfdTrustedEmbeddedFontGeometryProvider.Session)operation.OpenSession(new OfdTrustedEmbeddedFontGeometryProvider([new(expected, bytes.Length)]));
        OfdGlyphGeometry? measured = null;
        for (var i = 0; i < runs; i++)
        {
            measured = session.Measure(operation.Request(snapshot, face == "cjk" ? "AiW中文𠀋 " : "AiW ", 4, 400, false));
            Assert.True(measured.IsSupported, measured.UnsupportedReason);
        }
        var a = measured!.Clusters[0]; var space = measured.Clusters[^1];
        Assert.Equal((uint)(face == "cjk" ? 34 : 36), a.GlyphId);
        AssertBox(a.FontBoundsMillimeters!, face == "cjk" ? 0.016 : 0, face == "cjk" ? -2.932 : -2.868,
            face == "cjk" ? 2.416 : 2.552, 0);
        Assert.Equal(face == "cjk" ? 2.432 : 2.556, a.AdvanceXMillimeters, 9);
        Assert.Null(space.FontBoundsMillimeters); Assert.Equal(face == "cjk" ? 0.896 : 1.04, space.AdvanceXMillimeters, 9);
        if (face == "cjk")
        {
            Assert.Equal(2, measured.Clusters[5].Utf16Length); Assert.Equal((uint)59477, measured.Clusters[5].GlyphId);
            AssertBox(measured.Clusters[3].FontBoundsMillimeters!, 0.384, -3.36, 3.608, 0.316);
            AssertBox(measured.Clusters[4].FontBoundsMillimeters!, 0.144, -3.388, 3.868, 0.312);
        }
        else
        {
            Assert.Equal("UsedCompositeGlyphIsUnsupported", session.Measure(operation.Request(snapshot, "é", 4, 400, false)).UnsupportedReason);
            Assert.Equal("MissingGlyph", session.Measure(operation.Request(snapshot, "中", 4, 400, false)).UnsupportedReason);
        }
        Assert.Equal(1, operation.HashOperations); Assert.Equal(1, session.Statistics.FaceParses);
        Assert.Equal(face == "cjk" ? 7 : 4, session.BackendGlyphCalls);
        Assert.Equal(face == "cjk" ? 7 * runs : 4 * runs + 2, session.Statistics.GlyphRequests);
        operation.ValidateSourceFontsUnchanged(); operation.Dispose(); Assert.Equal(1, session.StreamsDisposed);
        output.WriteLine(System.Text.Json.JsonSerializer.Serialize(new { face, runs,
            operation.HashOperations, operation.CopiedFontBytes, operation.HashedFontBytes, operation.FontComparisonBytes,
            session.BackendAddAttempts, session.BackendGlyphCalls, session.ReadCalls, session.ReadBytes, session.SeekCalls,
            session.StreamsOpened, session.StreamsDisposed, statistics = session.Statistics,
            note = "Actual product counters; metadata totals are not observed outline visits. No timing/RSS guarantee." }));
        Assert.Equal(expected, Hash(File.ReadAllBytes(path)));
    }

    private static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", name + ".ttf"));
    private static OfdFontResource Resource(string id, byte[] bytes) => new() { Id = id, FontName = "SameName", Data = bytes };
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static OfdTrustedEmbeddedFontIdentity Identity(byte[] bytes) => new(Hash(bytes), bytes.Length);
    private static void AssertBox(OfdGlyphBox actual, double left, double top, double right, double bottom)
    { Assert.Equal(left, actual.Left, 9); Assert.Equal(top, actual.Top, 9); Assert.Equal(right, actual.Right, 9); Assert.Equal(bottom, actual.Bottom, 9); }
    private sealed class GrowingConfiguration(OfdTrustedEmbeddedFontIdentity identity) : IReadOnlyList<OfdTrustedEmbeddedFontIdentity>
    {
        private int countReads;
        internal int EntriesRead;
        public int Count => countReads++ == 0 ? 1 : 65;
        public OfdTrustedEmbeddedFontIdentity this[int index] { get { EntriesRead++; return identity; } }
        public IEnumerator<OfdTrustedEmbeddedFontIdentity> GetEnumerator() => throw new NotSupportedException();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}

public sealed class RealTrustedFontsTheoryAttribute : TheoryAttribute
{
    public RealTrustedFontsTheoryAttribute()
    {
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFDRW_TEST_TRUSTED_CJK_FONT")) ||
            string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFDRW_TEST_TRUSTED_LATIN_FONT")))
            Skip = "Set explicit OFDRW_TEST_TRUSTED_CJK_FONT and OFDRW_TEST_TRUSTED_LATIN_FONT to the pinned licensed original fixtures for the real-font gate.";
    }
}
