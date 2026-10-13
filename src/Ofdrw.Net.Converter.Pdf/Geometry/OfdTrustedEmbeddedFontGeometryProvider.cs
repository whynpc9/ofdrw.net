using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Numerics;
using System.Threading;
using Ofdrw.Net.Reader.Keywords;
using SixLabors.Fonts;
using SixLabors.Fonts.Unicode;

namespace Ofdrw.Net.Converter.Pdf.Geometry;

/// <summary>Opt-in original-font geometry for explicitly vetted identities and ordinary Regular simple glyphs.</summary>
/// <remarks>No embedded font is trusted by default. Profiles authorize identities, never replacement bytes.
/// Each operation admits at most four faces and uses no host resolver, synthetic style or global cache.
/// Snapshot bytes, metadata counts and read quotas are not RSS or CPU deadlines: SixLabors 1.0.1 calls
/// cannot be interrupted between reads. Observed cold Noto CJK Add thread allocation was about 489 MB,
/// and later observations about 129.5 MB. Output is a font glyph design box, not tight painted-pixel ink.
/// Only weight 400/nonitalic actual Regular, simple glyphs and known empty glyphs are supported.
/// Sizes are limited to 0.01 through 1000 mm. Ordinary Latin, CJK ideographs/punctuation and fullwidth scalars
/// are mapped individually without contextual shaping. Profiles do not authorize complex scripts.
/// Used composites, contextual shaping, variable/color/CFF/collection fonts remain unsupported.</remarks>
public sealed class OfdTrustedEmbeddedFontGeometryProvider : IOfdGlyphGeometryProvider
{
    /// <summary>Maximum constructor-configured identities; this is not the number of simultaneously opened faces.</summary>
    public const int MaximumTrustedIdentities = 64;
    /// <summary>Hard maximum bound faces per session, further reduced by caller limits.</summary>
    public const int MaximumOpenedFaces = 4;
    /// <summary>Hard cumulative source scalar requests per session, further reduced by caller limits.</summary>
    public const int MaximumGlyphRequests = 100_000;
    /// <summary>Hard cumulative cached scalar/size measurements per session.</summary>
    public const int MaximumCachedGlyphs = 8_192;
    /// <summary>Hard cumulative font read operations, including metadata and backend reads.</summary>
    public const long MaximumReadOperations = 100_000_000;
    /// <summary>Hard cumulative actual font bytes read, including rereads.</summary>
    public const long MaximumReadBytes = 512L * 1024 * 1024;
    /// <summary>Hard cumulative glyph metadata records admitted before their arrays are allocated.</summary>
    public const int MaximumMetadataGlyphRecords = 4 * 65_535;
    /// <summary>Hard cumulative whole-face simple points used for admission; these are metadata, not observed outline visits.</summary>
    public const long MaximumAdmittedSimplePoints = 16_000_000;

    private readonly Dictionary<string, OfdTrustedEmbeddedFontIdentity> trusted = new(StringComparer.Ordinal);

    /// <summary>Freezes an explicit application-vetted identity list. Empty lists trust nothing.</summary>
    public OfdTrustedEmbeddedFontGeometryProvider(IReadOnlyList<OfdTrustedEmbeddedFontIdentity> trustedFonts)
    {
        if (trustedFonts is null) throw new ArgumentNullException(nameof(trustedFonts));
        var count = trustedFonts.Count;
        if (count < 0 || count > MaximumTrustedIdentities) throw new ArgumentOutOfRangeException(nameof(trustedFonts));
        for (var i = 0; i < count; i++)
        {
            var identity = trustedFonts[i] ?? throw new ArgumentException("Null trusted font identity.", nameof(trustedFonts));
            if (trusted.TryGetValue(identity.Sha256, out var previous))
            {
                if (previous.ByteCount != identity.ByteCount || previous.Bold != identity.Bold || previous.Italic != identity.Italic)
                    throw new ArgumentException("Conflicting trusted identities for the same digest.", nameof(trustedFonts));
            }
            else trusted.Add(identity.Sha256, identity);
        }
    }

    /// <summary>Opens an independent, nonconcurrent session; dispose it on every outcome.</summary>
    public IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits limits, CancellationToken cancellationToken = default)
    {
        if (limits is null) throw new ArgumentNullException(nameof(limits));
        cancellationToken.ThrowIfCancellationRequested(); return new Session(trusted, limits);
    }

    internal sealed class Session : IOfdGlyphGeometrySession
    {
        private readonly IReadOnlyDictionary<string, OfdTrustedEmbeddedFontIdentity> trusted;
        private readonly OfdGlyphGeometryLimits limits;
        private readonly Dictionary<OfdGlyphFontSnapshot, Face> faces = new();
        private bool disposed; private int busy;
        private long admittedBytes, admittedPoints, parses, glyphRequests;
        private int metadataGlyphs, cachedGlyphs, faceReservations;
        private readonly Dictionary<OfdGlyphFontSnapshot, string> unsupportedFaces = new();
        internal long ReadCalls, ReadBytes, SeekCalls, BackendGlyphCalls;
        internal long BackendAddAttempts;
        internal int StreamsOpened, StreamsDisposed;
        internal int BoundFaces => faces.Count;
        internal int CachedGlyphs => cachedGlyphs;
        internal Action? BeforeBackendAdd; // deterministic cancellation/fault seams for friend tests only.
        internal Action? AfterBackendAdd;

        internal Session(IReadOnlyDictionary<string, OfdTrustedEmbeddedFontIdentity> trusted, OfdGlyphGeometryLimits limits)
        { this.trusted = trusted; this.limits = limits; }
        public OfdGlyphGeometryStatistics Statistics => new(parses, ReadBytes, glyphRequests, 0);

        public OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken cancellationToken = default)
        {
            if (disposed) throw new ObjectDisposedException(nameof(Session));
            if (Interlocked.CompareExchange(ref busy, 1, 0) != 0) throw new InvalidOperationException("Glyph sessions do not support reentry or concurrent use.");
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (request is null) throw new ArgumentNullException(nameof(request));
                if (request.Weight != 400 || request.Italic || request.Font.Bold || request.Font.Italic)
                    return Unsupported("OnlyActualRegularWeight400IsSupported");
                if (request.SizeMillimeters < 0.01 || request.SizeMillimeters > 1000)
                    return Unsupported("SizeOutsideSupportedMillimeterRange");
                if (!trusted.TryGetValue(request.Font.ContentSha256, out var identity) || identity.ByteCount != request.Font.ByteCount ||
                    identity.Bold != request.Font.Bold || identity.Italic != request.Font.Italic)
                    return Unsupported("UntrustedOriginalFontIdentityOrStyle");
                if (request.Text.Length > MaximumGlyphRequests * 2) throw new InvalidDataException("Trusted geometry request text capacity exceeded.");
                var scalars = new List<(int Scalar, int Start, int Length)>();
                var nextCheck = 0;
                for (var i = 0; i < request.Text.Length;)
                {
                    if (i >= nextCheck) { cancellationToken.ThrowIfCancellationRequested(); nextCheck = i + 256; }
                    var length = char.IsHighSurrogate(request.Text[i]) ? 2 : 1;
                    if (char.IsLowSurrogate(request.Text[i]) || length == 2 && (i + 1 >= request.Text.Length || !char.IsLowSurrogate(request.Text[i + 1])))
                        throw new InvalidDataException("Unpaired Unicode surrogate in glyph request.");
                    if (glyphRequests >= Math.Min(limits.MaxGlyphOperations, MaximumGlyphRequests)) throw new InvalidDataException("Trusted geometry glyph request capacity exceeded.");
                    glyphRequests++;
                    var scalar = char.ConvertToUtf32(request.Text, i);
                    var category = CharUnicodeInfo.GetUnicodeCategory(request.Text, i);
                    if (!OrdinaryScalar(scalar) || category is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or
                        UnicodeCategory.EnclosingMark or UnicodeCategory.Format or UnicodeCategory.Control)
                        return Unsupported("ContextualOrComplexGlyphMappingIsUnsupported");
                    scalars.Add((scalar, i, length)); i += length;
                }
                cancellationToken.ThrowIfCancellationRequested();
                if (scalars.Count == 0) return new OfdGlyphGeometry(Array.Empty<OfdGlyphClusterGeometry>());
                Face face;
                try { face = GetFace(request.Font, cancellationToken); }
                catch (NotSupportedException exception) { return Unsupported(exception.Message); }
                var result = new OfdGlyphClusterGeometry[scalars.Count];
                for (var i = 0; i < scalars.Count; i++)
                {
                    if ((i & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                    var scalar = scalars[i]; var key = (scalar.Scalar, request.SizeMillimeters);
                    if (!face.Cache.TryGetValue(key, out var measured))
                    {
                        if (cachedGlyphs >= MaximumCachedGlyphs) throw new InvalidDataException("Trusted geometry glyph cache capacity exceeded.");
                        measured = MeasureGlyph(face, scalar.Scalar, request.SizeMillimeters, cancellationToken);
                        face.Cache.Add(key, measured); cachedGlyphs++;
                    }
                    if (measured.Reason is not null) return Unsupported(measured.Reason);
                    result[i] = new OfdGlyphClusterGeometry(scalar.Start, scalar.Length, measured.GlyphId, measured.Box, measured.Advance);
                }
                cancellationToken.ThrowIfCancellationRequested(); return new OfdGlyphGeometry(result);
            }
            catch { Release(); throw; }
            finally { Volatile.Write(ref busy, 0); }
        }

        private Face GetFace(OfdGlyphFontSnapshot snapshot, CancellationToken token)
        {
            if (faces.TryGetValue(snapshot, out var existing)) return existing;
            if (unsupportedFaces.TryGetValue(snapshot, out var reason)) throw new NotSupportedException(reason);
            if (faceReservations >= Math.Min(limits.MaxFontFaces, MaximumOpenedFaces) || snapshot.ByteCount > limits.MaxFontBytes ||
                snapshot.ByteCount > limits.MaxTotalFontBytes - admittedBytes)
                throw new InvalidDataException("Trusted geometry face or original-byte admission capacity exceeded.");
            // Reserve before opening/parsing. Exceptions release the entire session; no half-bound face remains.
            admittedBytes += snapshot.ByteCount;
            faceReservations++;
            MeterStream? stream = null;
            try
            {
                stream = new MeterStream(snapshot.OpenRead(), this, token); StreamsOpened++;
                var metadata = TrustedFontMetadata.Read(stream, count =>
                {
                    if (count > MaximumMetadataGlyphRecords - metadataGlyphs) throw new InvalidDataException("Trusted metadata record capacity exceeded.");
                    metadataGlyphs += count;
                }, token);
                if (metadata.SimplePoints > MaximumAdmittedSimplePoints - admittedPoints)
                    throw new InvalidDataException("Trusted whole-face point admission capacity exceeded.");
                admittedPoints += metadata.SimplePoints;
                stream.Position = 0;
                var collection = new FontCollection();
                BeforeBackendAdd?.Invoke(); token.ThrowIfCancellationRequested();
                BackendAddAttempts++;
                var family = collection.Add(stream, out var description);
                parses++;
                token.ThrowIfCancellationRequested(); AfterBackendAdd?.Invoke(); token.ThrowIfCancellationRequested();
                if (description.Style != FontStyle.Regular) throw new NotSupportedException("Backend actual face style is not Regular.");
                var face = new Face(stream, collection, family, metadata);
                faces.Add(snapshot, face); stream = null; return face;
            }
            catch (NotSupportedException exception) { unsupportedFaces.Add(snapshot, exception.Message); throw; }
            finally { stream?.Dispose(); }
        }

        private CachedGlyph MeasureGlyph(Face face, int scalar, double size, CancellationToken token)
        {
            var gid = face.Metadata.Map(scalar);
            if (gid == 0) return new CachedGlyph("MissingGlyph");
            if (gid >= face.Metadata.GlyphCount) throw new InvalidDataException("Trusted glyph mapping is out of range.");
            var raw = face.Metadata[gid];
            if (raw.Contours < 0) return new CachedGlyph("UsedCompositeGlyphIsUnsupported");
            if (raw.Points > 4096) return new CachedGlyph("UsedGlyphPointCapacityIsUnsupported");
            token.ThrowIfCancellationRequested();
            var font = face.Family.CreateFont((float)(size * 72 / 25.4), FontStyle.Regular);
            BackendGlyphCalls++;
            var found = font.TryGetGlyphs(new CodePoint(scalar), ColorFontSupport.None, out var glyphs);
            token.ThrowIfCancellationRequested();
            if (!found || glyphs is null || glyphs.Count != 1 || glyphs[0].GlyphMetrics.GlyphId != gid ||
                glyphs[0].GlyphMetrics.UnitsPerEm != face.Metadata.UnitsPerEm || glyphs[0].GlyphMetrics.AdvanceWidth != raw.Advance)
                return new CachedGlyph("BackendOriginalGlyphMappingMismatch");
            var scale = size / face.Metadata.UnitsPerEm;
            OfdGlyphBox? box = raw.Empty ? null : new OfdGlyphBox(raw.XMin * scale, -raw.YMax * scale, raw.XMax * scale, -raw.YMin * scale);
            if (box is not null)
            {
                var observed = glyphs[0].BoundingBox(GlyphLayoutMode.Horizontal, Vector2.Zero, 72);
                var unit = 25.4 / 72; var tolerance = Math.Max(0.0000001, size * 0.0000005);
                if (!Same(observed.Left * unit, box.Left, tolerance) || !Same(observed.Top * unit, box.Top, tolerance) ||
                    !Same(observed.Right * unit, box.Right, tolerance) || !Same(observed.Bottom * unit, box.Bottom, tolerance))
                    return new CachedGlyph("BackendOriginalGlyphBoundsMismatch");
            }
            token.ThrowIfCancellationRequested(); return new CachedGlyph((uint)gid, box, raw.Advance * scale);
        }

        private static bool Same(double observed, double expected, double tolerance) =>
            !double.IsNaN(observed) && !double.IsInfinity(observed) && Math.Abs(observed - expected) <= tolerance;

        private static bool OrdinaryScalar(int cp) => cp is >= 0x20 and <= 0x24f or >= 0x2000 and <= 0x206f or
            >= 0x3000 and <= 0x303f or >= 0x3400 and <= 0x4dbf or >= 0x4e00 and <= 0x9fff or
            >= 0xf900 and <= 0xfaff or >= 0xff00 and <= 0xffef or >= 0x20000 and <= 0x2ffff;
        private static OfdGlyphGeometry Unsupported(string reason) => new(Array.Empty<OfdGlyphClusterGeometry>(), reason);
        public void Dispose()
        {
            if (disposed) return;
            if (Volatile.Read(ref busy) != 0) throw new InvalidOperationException("Cannot concurrently dispose a measuring glyph session.");
            Release();
        }
        private void Release()
        {
            if (disposed) return;
            disposed = true;
            try { foreach (var face in faces.Values) face.Dispose(); }
            finally { faces.Clear(); unsupportedFaces.Clear(); cachedGlyphs = 0; BeforeBackendAdd = null; AfterBackendAdd = null; }
        }
        private sealed class Face(MeterStream stream, FontCollection collection, FontFamily family, TrustedFontMetadata metadata) : IDisposable
        {
            internal readonly FontFamily Family = family;
            internal readonly TrustedFontMetadata Metadata = metadata;
            internal readonly Dictionary<(int Scalar, double Size), CachedGlyph> Cache = new();
            // Keep the collection alive with its face until disposal; no global registry is touched.
            private FontCollection? owner = collection;
            public void Dispose() { Cache.Clear(); owner = null; stream.Dispose(); }
        }
        private sealed class CachedGlyph
        {
            internal CachedGlyph(string reason) => Reason = reason;
            internal CachedGlyph(uint glyphId, OfdGlyphBox? box, double advance) { GlyphId = glyphId; Box = box; Advance = advance; }
            internal readonly string? Reason;
            internal readonly uint GlyphId;
            internal readonly OfdGlyphBox? Box;
            internal readonly double Advance;
        }
        private sealed class MeterStream(Stream source, Session owner, CancellationToken token) : Stream
        {
            private bool disposed;
            public override bool CanRead => !disposed && source.CanRead;
            public override bool CanSeek => !disposed && source.CanSeek;
            public override bool CanWrite => false;
            public override long Length => source.Length;
            public override long Position { get => source.Position; set { owner.SeekCalls++; source.Position = value; } }
            private void ChargeRead()
            {
                if (disposed) throw new ObjectDisposedException(nameof(MeterStream));
                if ((owner.ReadCalls & 4095) == 0) token.ThrowIfCancellationRequested();
                if (owner.ReadCalls >= MaximumReadOperations || owner.ReadBytes >= MaximumReadBytes)
                    throw new InvalidDataException("Trusted geometry font read-work capacity exceeded.");
                owner.ReadCalls++;
            }
            public override int Read(byte[] buffer, int offset, int count)
            {
                ChargeRead(); var n = source.Read(buffer, offset, (int)Math.Min(count, MaximumReadBytes - owner.ReadBytes)); owner.ReadBytes += n; return n;
            }
            public override int ReadByte() { ChargeRead(); var n = source.ReadByte(); if (n >= 0) owner.ReadBytes++; return n; }
            public override long Seek(long offset, SeekOrigin origin) { owner.SeekCalls++; return source.Seek(offset, origin); }
            public override void Flush() { }
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            protected override void Dispose(bool disposing)
            { if (disposing && !disposed) { disposed = true; owner.StreamsDisposed++; source.Dispose(); } base.Dispose(disposing); }
        }
    }
}
