using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Reader.Keywords;

/// <summary>Opens independent glyph-measurement sessions without global font state or fallback.</summary>
public interface IOfdGlyphGeometryProvider
{
    /// <summary>Opens one operation-local session. The caller disposes it on success, failure or cancellation.</summary>
    IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits limits, CancellationToken cancellationToken = default);
}

/// <summary>Owns operation-local font faces. Implementations must honor limits and report unsupported mappings.</summary>
public interface IOfdGlyphGeometrySession : IDisposable
{
    /// <summary>Measures the exact snapshot font. Equal immutable requests within one session must produce consistent geometry;
    /// editing may reuse their validated immutable result. Never substitute host fonts or synthetic styles.</summary>
    OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken cancellationToken = default);
    /// <summary>Gets observed provider work, rather than inferred source counts.</summary>
    OfdGlyphGeometryStatistics Statistics { get; }
}

/// <summary>Identifies the range represented by glyph bounds; neither kind implies final painted-pixel bounds.</summary>
public enum OfdGlyphBoundsKind
{
    /// <summary>Baseline-relative bounds from the actual original font glyph, with Y down.</summary>
    FontGlyphBounds,
    /// <summary>A transformed font glyph box. Its AABB need not be a tight transformed outline.</summary>
    TransformedFontGlyphBounds
}

/// <summary>A finite, axis-aligned millimeter box.</summary>
public sealed class OfdGlyphBox
{
    /// <summary>Creates a box; whitespace should use a null box rather than a fabricated rectangle.</summary>
    public OfdGlyphBox(double left, double top, double right, double bottom)
    {
        if (!Finite(left) || !Finite(top) || !Finite(right) || !Finite(bottom) || right < left || bottom < top)
            throw new ArgumentException("Glyph box must be finite and ordered.");
        Left = left; Top = top; Right = right; Bottom = bottom;
    }
    /// <summary>Left coordinate in millimeters.</summary>
    public double Left { get; }
    /// <summary>Top coordinate in millimeters.</summary>
    public double Top { get; }
    /// <summary>Right coordinate in millimeters.</summary>
    public double Right { get; }
    /// <summary>Bottom coordinate in millimeters.</summary>
    public double Bottom { get; }
    internal static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>A provider-proven cluster-to-glyph mapping. Source offsets use UTF-16 units.</summary>
public sealed class OfdGlyphClusterGeometry
{
    /// <summary>Creates a mapped cluster. Null bounds indicate no ink; advance is still retained.</summary>
    public OfdGlyphClusterGeometry(int utf16Start, int utf16Length, uint glyphId,
        OfdGlyphBox? fontBoundsMillimeters, double advanceXMillimeters, double advanceYMillimeters = 0)
    {
        if (utf16Start < 0 || utf16Length <= 0 || utf16Length > int.MaxValue - utf16Start)
            throw new ArgumentOutOfRangeException(nameof(utf16Length));
        if (glyphId == 0) throw new ArgumentOutOfRangeException(nameof(glyphId), "Missing glyphs are unsupported, not verified mappings.");
        if (!OfdGlyphBox.Finite(advanceXMillimeters) || !OfdGlyphBox.Finite(advanceYMillimeters)) throw new ArgumentException("Advance must be finite.");
        Utf16Start = utf16Start; Utf16Length = utf16Length; GlyphId = glyphId;
        FontBoundsMillimeters = fontBoundsMillimeters; AdvanceXMillimeters = advanceXMillimeters; AdvanceYMillimeters = advanceYMillimeters;
    }
    /// <summary>Source start in UTF-16 units.</summary>
    public int Utf16Start { get; }
    /// <summary>Source length in UTF-16 units.</summary>
    public int Utf16Length { get; }
    /// <summary>Verified glyph ID in the exact snapshot font.</summary>
    public uint GlyphId { get; }
    /// <summary>Original font glyph bounds relative to its baseline; null means no ink.</summary>
    public OfdGlyphBox? FontBoundsMillimeters { get; }
    /// <summary>Horizontal advance in millimeters.</summary>
    public double AdvanceXMillimeters { get; }
    /// <summary>Vertical advance in millimeters.</summary>
    public double AdvanceYMillimeters { get; }
}

/// <summary>An immutable provider result. Unsupported results cannot contain purported verified geometry.</summary>
public sealed class OfdGlyphGeometry
{
    /// <summary>Copies verified mappings or returns an unsupported reason. The operation validates mapping coverage against source text.</summary>
    public OfdGlyphGeometry(IReadOnlyList<OfdGlyphClusterGeometry> clusters, string? unsupportedReason = null)
    {
        if (clusters is null) throw new ArgumentNullException(nameof(clusters));
        var count = clusters.Count;
        if (count < 0 || count > 16_000_000) throw new ArgumentOutOfRangeException(nameof(clusters));
        if (unsupportedReason is not null && (string.IsNullOrWhiteSpace(unsupportedReason) || count != 0))
            throw new ArgumentException("Unsupported geometry requires a reason and no verified mappings.");
        var copy = new OfdGlyphClusterGeometry[count];
        for (var i = 0; i < copy.Length; i++) copy[i] = clusters[i] ?? throw new ArgumentException("Null cluster.", nameof(clusters));
        Clusters = new ReadOnlyCollection<OfdGlyphClusterGeometry>(copy); UnsupportedReason = unsupportedReason;
    }
    /// <summary>Gets the immutable mappings.</summary>
    public IReadOnlyList<OfdGlyphClusterGeometry> Clusters { get; }
    /// <summary>Gets the reason the exact source font/mapping cannot be measured.</summary>
    public string? UnsupportedReason { get; }
    /// <summary>Whether the provider returned supported geometry; source coverage still requires validation.</summary>
    public bool IsSupported => UnsupportedReason is null;
    /// <summary>Gets the baseline-relative bounds kind.</summary>
    public OfdGlyphBoundsKind BoundsKind => OfdGlyphBoundsKind.FontGlyphBounds;
}

/// <summary>Observed session counters. These counters do not establish a hard CPU or process-memory bound.</summary>
public sealed class OfdGlyphGeometryStatistics
{
    /// <summary>Creates a statistics snapshot from actual measured operations.</summary>
    public OfdGlyphGeometryStatistics(long faceParses = 0, long fontBytesRead = 0, long glyphRequests = 0, long outlinePointVisits = 0)
    {
        if (faceParses < 0 || fontBytesRead < 0 || glyphRequests < 0 || outlinePointVisits < 0) throw new ArgumentOutOfRangeException(nameof(faceParses));
        FaceParses = faceParses; FontBytesRead = fontBytesRead; GlyphRequests = glyphRequests; OutlinePointVisits = outlinePointVisits;
    }
    /// <summary>Actual parsed font faces.</summary>
    public long FaceParses { get; }
    /// <summary>Actual bytes read by the provider, including rereads.</summary>
    public long FontBytesRead { get; }
    /// <summary>Actual glyph requests.</summary>
    public long GlyphRequests { get; }
    /// <summary>Actual outline point/curve visits.</summary>
    public long OutlinePointVisits { get; }
}

/// <summary>A library-created measurement request using an operation-owned original-font snapshot.</summary>
public sealed class OfdGlyphGeometryRequest
{
    internal OfdGlyphGeometryRequest(OfdGlyphFontSnapshot font, string text, double sizeMillimeters, int weight, bool italic)
    {
        Font = font ?? throw new ArgumentNullException(nameof(font)); Text = text ?? throw new ArgumentNullException(nameof(text));
        if (!OfdGlyphBox.Finite(sizeMillimeters) || sizeMillimeters <= 0) throw new ArgumentOutOfRangeException(nameof(sizeMillimeters));
        if (weight < 100 || weight > 900) throw new ArgumentOutOfRangeException(nameof(weight));
        SizeMillimeters = sizeMillimeters; Weight = weight; Italic = italic;
    }
    /// <summary>The exact original font, identified by internally computed content identity.</summary>
    public OfdGlyphFontSnapshot Font { get; }
    /// <summary>Source text without normalization or injected whitespace.</summary>
    public string Text { get; }
    /// <summary>Requested size in millimeters.</summary>
    public double SizeMillimeters { get; }
    /// <summary>Source weight. A provider must reject unsupported synthetic rendering.</summary>
    public int Weight { get; }
    /// <summary>Source italic flag. A provider must reject unsupported synthetic rendering.</summary>
    public bool Italic { get; }
}

/// <summary>An operation-owned, immutable copy of original embedded bytes. It cannot be caller-constructed from an unverified digest.</summary>
public sealed class OfdGlyphFontSnapshot
{
    private byte[]? data;
    private readonly List<Stream> streams = new();
    private readonly Action chargeStream;
    private readonly Action<int>? chargeRead;
    private readonly Action? chargeSeek;
    internal OfdGlyphFontSnapshot(string resourceId, bool bold, bool italic, byte[] ownedData, string sha256, Action chargeStream,
        Action<int>? chargeRead = null, Action? chargeSeek = null)
    {
        ResourceId = resourceId; Bold = bold; Italic = italic; data = ownedData;
        ByteCount = ownedData.Length; ContentSha256 = sha256; this.chargeStream = chargeStream; this.chargeRead = chargeRead; this.chargeSeek = chargeSeek;
    }
    /// <summary>Exact, case-sensitive resource ID.</summary>
    public string ResourceId { get; }
    /// <summary>Actual resource bold flag; not permission to synthesize bold.</summary>
    public bool Bold { get; }
    /// <summary>Actual resource italic flag; not permission to synthesize italic.</summary>
    public bool Italic { get; }
    /// <summary>Original byte length.</summary>
    public int ByteCount { get; }
    /// <summary>SHA256 computed internally over the private original-byte copy.</summary>
    public string ContentSha256 { get; }
    /// <summary>Opens a bounded, read-only stream. Its buffer is hidden; operation disposal closes even unclosed streams.</summary>
    public Stream OpenRead()
    {
        if (data is null) throw new ObjectDisposedException(nameof(OfdGlyphFontSnapshot));
        chargeStream();
        var stream = new SnapshotReadStream(data, chargeRead, chargeSeek);
        streams.Add(stream); return stream;
    }
    internal byte[] Data => data ?? throw new ObjectDisposedException(nameof(OfdGlyphFontSnapshot));
    internal void Release()
    {
        try { foreach (var stream in streams) stream.Dispose(); }
        finally { streams.Clear(); data = null; }
    }

    // MemoryStream.CopyTo/WriteTo can pass the backing array to another Stream.
    // Base Stream copy methods instead consume this copying Read implementation.
    private sealed class SnapshotReadStream : Stream
    {
        private byte[]? bytes;
        private long position;
        private readonly Action<int>? chargeRead;
        private readonly Action? chargeSeek;
        internal SnapshotReadStream(byte[] bytes, Action<int>? chargeRead, Action? chargeSeek)
        { this.bytes = bytes; this.chargeRead = chargeRead; this.chargeSeek = chargeSeek; }
        public override bool CanRead => bytes is not null;
        public override bool CanSeek => bytes is not null;
        public override bool CanWrite => false;
        public override long Length => Buffer.Length;
        public override long Position
        {
            get { _ = Buffer; return position; }
            set { _ = Buffer; if (value < 0) throw new ArgumentOutOfRangeException(nameof(value)); chargeSeek?.Invoke(); position = value; }
        }
        private byte[] Buffer => bytes ?? throw new ObjectDisposedException(nameof(SnapshotReadStream));
        public override int Read(byte[] buffer, int offset, int count)
        {
            if (buffer is null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
            var source = Buffer;
            var length = (int)Math.Min(count, Math.Max(0, source.Length - position));
            if (length > 0) { chargeRead?.Invoke(length); System.Buffer.BlockCopy(source, (int)position, buffer, offset, length); }
            position += length; return length;
        }
        public override int ReadByte()
        {
            var source = Buffer;
            if (position >= source.Length) return -1;
            chargeRead?.Invoke(1); return source[(int)position++];
        }
        public override long Seek(long offset, SeekOrigin origin)
        {
            _ = Buffer; chargeSeek?.Invoke();
            var basis = origin switch { SeekOrigin.Begin => 0, SeekOrigin.Current => position, SeekOrigin.End => Length,
                _ => throw new ArgumentOutOfRangeException(nameof(origin)) };
            long next;
            try { next = checked(basis + offset); }
            catch (OverflowException exception) { throw new IOException("Font stream seek overflow.", exception); }
            if (next < 0) throw new IOException("Cannot seek before the beginning of the font.");
            return position = next;
        }
        public override void Flush() { _ = Buffer; }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { bytes = null; base.Dispose(disposing); }
    }
}
