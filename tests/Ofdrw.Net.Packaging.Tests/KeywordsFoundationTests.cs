using System.Security.Cryptography;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class KeywordsFoundationTests
{
    [Fact]
    public void SnapshotPreservesOriginalBytesAndDetectsChangedSourceWithoutRehashing()
    {
        var source = Font("10", [1, 2, 3, 4]); source.Bold = true;
        using var operation = new OfdKeywordOperation();
        var snapshot = operation.CaptureFont(source, "10");
        for (var i = 0; i < 100; i++) Assert.Same(snapshot, operation.CaptureFont(source, "10"));
        Assert.Equal(Convert.ToHexString(SHA256.HashData(source.Data)).ToLowerInvariant(), snapshot.ContentSha256);
        Assert.True(snapshot.Bold);
        Assert.Equal(1, operation.HashOperations); Assert.Equal(4, operation.CopiedFontBytes); Assert.Equal(4, operation.HashedFontBytes);
        using var read = snapshot.OpenRead();
        Assert.False(read.CanWrite);
        Assert.False(read is MemoryStream);
        source.Data[0] = 9;
        Assert.Equal(1, read.ReadByte());
        Assert.Throws<InvalidDataException>(operation.ValidateSourceFontsUnchanged);
        Assert.Equal(1, operation.HashOperations);
    }

    [Fact]
    public async Task HostileCopyDestinationCannotMutateTheOwnedFontSnapshot()
    {
        using var operation = new OfdKeywordOperation();
        var font = Font("A", [1, 2, 3, 4]); var snapshot = operation.CaptureFont(font, "A");
        using (var source = snapshot.OpenRead()) using (var destination = new BufferMutatingDestination()) source.CopyTo(destination);
        using (var source = snapshot.OpenRead()) using (var destination = new BufferMutatingDestination()) await source.CopyToAsync(destination);
        using var reread = snapshot.OpenRead();
        var actual = new byte[4]; Assert.Equal(4, reread.Read(actual, 0, 4));
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, actual);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(actual)).ToLowerInvariant(), snapshot.ContentSha256);
        operation.ValidateSourceFontsUnchanged(); Assert.Equal(1, operation.HashOperations);
    }

    [Theory]
    [InlineData("id")]
    [InlineData("bold")]
    [InlineData("italic")]
    public void SnapshotChecksActualResourceIdentityAndStyle(string mutation)
    {
        var font = Font("A", [1]);
        using var operation = new OfdKeywordOperation(); operation.CaptureFont(font, "A");
        if (mutation == "id") font.Id = "a";
        if (mutation == "bold") font.Bold = true;
        if (mutation == "italic") font.Italic = true;
        Assert.Throws<InvalidDataException>(operation.ValidateSourceFontsUnchanged);
    }

    [Fact]
    public void IdenticalReplacementArrayRetainsIdentityButDifferentResourceCannotAliasId()
    {
        var font = Font("A", [1, 2]);
        using var operation = new OfdKeywordOperation(); operation.CaptureFont(font, "A");
        font.Data = [1, 2]; operation.ValidateSourceFontsUnchanged();
        Assert.Equal(2, operation.FontComparisonBytes); Assert.Equal(1, operation.HashOperations);
        Assert.Throws<InvalidDataException>(() => operation.CaptureFont(Font("A", [1, 2]), "A"));
        Assert.Throws<InvalidDataException>(() => operation.CaptureFont(font, "a"));
        Assert.Throws<NotSupportedException>(() => operation.CaptureFont(Font("B", []), "B"));
    }

    [Fact]
    public void AggregateCapacityIsCheckedBeforeCopyingOrHashingAnotherFont()
    {
        using var operation = new OfdKeywordOperation(new(geometryLimits: new(maxFontBytes: 4, maxTotalFontBytes: 4)));
        operation.CaptureFont(Font("A", [1, 2, 3]), "A");
        Assert.Throws<InvalidDataException>(() => operation.CaptureFont(Font("B", [4, 5]), "B"));
        Assert.Equal(3, operation.CopiedFontBytes); Assert.Equal(1, operation.HashOperations);
    }

    [Fact]
    public void OperationClosesUnclosedStreamsAndBoundsCumulativeReopens()
    {
        var operation = new OfdKeywordOperation(new(limits: new(maxFontReadStreams: 1)));
        var snapshot = operation.CaptureFont(Font("A", [1]), "A");
        var stream = snapshot.OpenRead();
        Assert.Throws<InvalidDataException>(() => snapshot.OpenRead());
        operation.Dispose(); operation.Dispose();
        Assert.False(stream.CanRead);
        Assert.Throws<ObjectDisposedException>(() => snapshot.OpenRead());
        Assert.Throws<ObjectDisposedException>(() => operation.CaptureFont(Font("B", [1]), "B"));
    }

    [Fact]
    public void SessionIsOpenedOnceAndProviderDisposeFailureStillReleasesSnapshots()
    {
        var provider = new TestProvider { ThrowOnDispose = true };
        var operation = new OfdKeywordOperation(); var snapshot = operation.CaptureFont(Font("A", [1]), "A");
        var stream = snapshot.OpenRead();
        var session = operation.OpenSession(provider);
        Assert.Same(session, operation.OpenSession(provider)); Assert.Equal(1, provider.OpenCalls);
        Assert.Throws<InvalidOperationException>(() => operation.OpenSession(new TestProvider()));
        Assert.Throws<IOException>(operation.Dispose);
        Assert.False(stream.CanRead); Assert.Equal(1, provider.DisposeCalls);
        operation.Dispose(); Assert.Equal(1, provider.DisposeCalls);
    }

    [Fact]
    public void CancellationDuringProviderOpenStillHasAReleasableOwner()
    {
        using var cancel = new CancellationTokenSource();
        var provider = new TestProvider { OnOpen = cancel.Cancel };
        var operation = new OfdKeywordOperation(cancellationToken: cancel.Token);
        var snapshot = operation.CaptureFont(Font("A", [1]), "A"); var stream = snapshot.OpenRead();
        try { Assert.Throws<OperationCanceledException>(() => operation.OpenSession(provider)); }
        finally { operation.Dispose(); }
        Assert.False(stream.CanRead); Assert.Equal(1, provider.DisposeCalls);
        Assert.Throws<ObjectDisposedException>(() => snapshot.OpenRead());
    }

    [Fact]
    public void CancelledWorkCannotOpenStreamsAndSnapshotsCannotCrossOperationScopes()
    {
        using var cancel = new CancellationTokenSource();
        using var first = new OfdKeywordOperation(cancellationToken: cancel.Token);
        using var second = new OfdKeywordOperation();
        var font = first.CaptureFont(Font("A", [1]), "A");
        Assert.Throws<ArgumentException>(() => second.Request(font, "A", 4, 400, false));
        cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => font.OpenRead());
        Assert.Throws<OperationCanceledException>(first.ValidateSourceFontsUnchanged);
    }

    [Fact]
    public void OrdinalMatcherReturnsOverlapsInOrderAndDoesNotNormalizeOrIgnoreCase()
    {
        using var operation = new OfdKeywordOperation();
        var matches = new OfdOrdinalMatcher("aa", operation).Find("aaaa", Scalars("aaaa"));
        Assert.Equal([0, 1, 2], matches.Select(m => m.Utf16Start));
        Assert.All(matches, m => { Assert.Equal(2, m.Utf16Length); Assert.Equal(2, m.ClusterCount); });
        Assert.Empty(new OfdOrdinalMatcher("A", operation).Find("a", Scalars("a")));
        Assert.Empty(new OfdOrdinalMatcher("é", operation).Find("e\u0301", [Cluster(0, 2)]));
        Assert.Empty(new OfdOrdinalMatcher("absent", operation).Find("present", Scalars("present")));
    }

    [Fact]
    public void MatchesRespectVerifiedClustersAndSupplementaryScalarOffsets()
    {
        using var operation = new OfdKeywordOperation();
        var text = "A\u0301𝟘B";
        OfdGlyphClusterGeometry[] clusters = [Cluster(0, 2), Cluster(2, 2), Cluster(4, 1)];
        Assert.Empty(new OfdOrdinalMatcher("A", operation).Find(text, clusters));
        Assert.Empty(new OfdOrdinalMatcher("\u0301", operation).Find(text, clusters));
        var hit = Assert.Single(new OfdOrdinalMatcher("𝟘B", operation).Find(text, clusters));
        Assert.Equal(2, hit.Utf16Start); Assert.Equal(3, hit.Utf16Length); Assert.Equal(1, hit.ClusterStart); Assert.Equal(2, hit.ClusterCount);
        Assert.Throws<InvalidDataException>(() => new OfdOrdinalMatcher("B", operation).Find("𝟘B", [Cluster(0, 1), Cluster(1, 2)]));
    }

    [Theory]
    [InlineData(0xD800)]
    [InlineData(0xDC00)]
    public void InvalidUnicodeNeverBecomesAnEmptySuccessfulMatch(int codeUnit)
    {
        // Attribute serialization replaces isolated surrogates; construct them at runtime.
        var bad = new string((char)codeUnit, 1);
        using var operation = new OfdKeywordOperation();
        Assert.Throws<InvalidDataException>(() => new OfdOrdinalMatcher(bad, operation));
        Assert.Throws<InvalidDataException>(() => new OfdOrdinalMatcher("A", operation).Find(bad, [Cluster(0, 1)]));
    }

    [Fact]
    public void EmptyKeywordAndIncompleteOrUnorderedMappingsAreRejected()
    {
        using var operation = new OfdKeywordOperation();
        Assert.Throws<ArgumentException>(() => new OfdOrdinalMatcher("", operation));
        var matcher = new OfdOrdinalMatcher("A", operation);
        Assert.Throws<InvalidDataException>(() => matcher.Find("AB", [Cluster(0, 1)]));
        Assert.Throws<InvalidDataException>(() => matcher.Find("AB", [Cluster(1, 1), Cluster(0, 1)]));
        Assert.Empty(matcher.Find("", []));
    }

    [Fact]
    public void WorkLimitsAreCumulativeAcrossRunsAndNeverReturnTruncatedHitSuccess()
    {
        using var scan = new OfdKeywordOperation(new(limits: new(maxTextUtf16Units: 3)));
        var matcher = new OfdOrdinalMatcher("Z", scan);
        Assert.Empty(matcher.Find("AB", Scalars("AB")));
        Assert.Throws<InvalidDataException>(() => matcher.Find("CD", Scalars("CD")));
        using var hits = new OfdKeywordOperation(new(limits: new(maxHits: 1)));
        Assert.Throws<InvalidDataException>(() => new OfdOrdinalMatcher("aa", hits).Find("aaa", Scalars("aaa")));
        using var work = new OfdKeywordOperation(new(limits: new(maxMatchComparisons: 2)));
        Assert.Throws<InvalidDataException>(() => new OfdOrdinalMatcher("Z", work).Find("ABCD", Scalars("ABCD")));
    }

    [Fact]
    public void RepetitiveSearchHasBoundedLinearComparisonsAndObservesCancellation()
    {
        using var cancel = new CancellationTokenSource();
        using var operation = new OfdKeywordOperation(cancellationToken: cancel.Token);
        var matcher = new OfdOrdinalMatcher(new string('a', 128) + "b", operation);
        var source = new string('a', 4096);
        Assert.Empty(matcher.Find(source, Scalars(source)));
        Assert.True(operation.MatchComparisons < 4 * (source.Length + 129));
        cancel.Cancel(); Assert.Throws<OperationCanceledException>(() => matcher.Find("a", Scalars("a")));
    }

    [Fact]
    public void SupplementaryScalarsCannotSkipPeriodicCancellationCheckpoints()
    {
        using var operation = new OfdKeywordOperation();
        var text = "A" + string.Concat(Enumerable.Repeat("𝟘", 1024));
        var before = operation.CancellationChecks;
        OfdOrdinalMatcher.ValidateUnicode(text, operation);
        Assert.True(operation.CancellationChecks - before >= text.Length / 256);
    }

    [Fact]
    public void GeometryCopiesMappingsAndKeepsWhitespaceAdvanceWithoutInventingInk()
    {
        OfdGlyphClusterGeometry[] mapped = [new(0, 1, 1, null, 2)];
        var geometry = new OfdGlyphGeometry(mapped); mapped[0] = Cluster(0, 1);
        Assert.Null(geometry.Clusters[0].FontBoundsMillimeters); Assert.Equal(2, geometry.Clusters[0].AdvanceXMillimeters);
        Assert.False(new OfdGlyphGeometry([], "MissingGlyph").IsSupported);
        Assert.Throws<ArgumentException>(() => new OfdGlyphGeometry(mapped, "MissingGlyph"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdGlyphClusterGeometry(0, 1, 0, null, 1));
        Assert.Throws<ArgumentException>(() => new OfdGlyphBox(0, double.NaN, 1, 1));
    }

    private static OfdFontResource Font(string id, byte[] bytes) => new() { Id = id, FontName = "SameName", Data = bytes };
    private static OfdGlyphClusterGeometry Cluster(int start, int length) => new(start, length, 1, new(0, -1, 1, 0), 1);
    private static OfdGlyphClusterGeometry[] Scalars(string text)
    {
        var result = new List<OfdGlyphClusterGeometry>(); var start = 0;
        foreach (var rune in text.EnumerateRunes()) { result.Add(Cluster(start, rune.Utf16SequenceLength)); start += rune.Utf16SequenceLength; }
        return result.ToArray();
    }
    private sealed class TestProvider : IOfdGlyphGeometryProvider, IOfdGlyphGeometrySession
    {
        internal int OpenCalls; internal int DisposeCalls; internal bool ThrowOnDispose; internal Action? OnOpen;
        public IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits limits, CancellationToken cancellationToken = default)
        { OpenCalls++; OnOpen?.Invoke(); return this; }
        public OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public OfdGlyphGeometryStatistics Statistics => new();
        public void Dispose() { DisposeCalls++; if (ThrowOnDispose) throw new IOException("Probe disposal failure."); }
    }
    private sealed class BufferMutatingDestination : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => Array.Fill(buffer, (byte)99, offset, count);
    }
}
