using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using SkiaSharp;

namespace Ofdrw.Net.Graphics.SkiaSharp;

public sealed partial class SkiaDrawEvent
{
    // Private construction and digest ownership prevent an ordinary producer
    // from supplying a claimed identity. IVT is encapsulation, not a sandbox.
    private sealed class FontIdentity
    {
        internal readonly byte[] Hash;
        internal readonly bool Bold, Italic;
        internal FontIdentity(byte[] hash, bool bold, bool italic)
        { Hash = hash; Bold = bold; Italic = italic; }
    }

    private static void ValidateFixedFace(SKTypeface face)
    {
        if (face.GetTableSize(0x66766172u) != 0)
            throw new NotSupportedException("Variable font faces are not supported by this fixed-face profile.");
        using var style = face.FontStyle;
        if (style.Width != (int)SKFontStyleWidth.Normal)
            throw new NotSupportedException("Condensed/expanded font faces are not supported.");
    }

    private static (bool Bold, bool Italic) FaceStyle(SKTypeface face)
    {
        var selection = FontWord(face, 0x4f532f32u, 62);
        var macStyle = FontWord(face, 0x68656164u, 44);
        var bold = (selection & 0x20) != 0; var italic = (selection & 1) != 0;
        if (((macStyle & 1) != 0) != bold || ((macStyle & 2) != 0) != italic)
            throw new NotSupportedException("OS/2 and head font style flags must agree; ambiguous faces require a different font profile.");
        return (bold, italic);
    }

    // Native font backends may repair a short table before GetTableSize reads it.
    // Check the original sfnt directory and style bytes independently.
    private static (bool Bold, bool Italic) RawStyle(int length, Func<int, int, byte[]> read)
    {
        if (length < 12) throw new NotSupportedException("Original font directory is truncated.");
        var header = read(0, 12);
        uint version = U32(header, 0);
        if (version is not (0x00010000 or 0x4f54544f or 0x74727565))
            throw new NotSupportedException("Original font must be a standalone TrueType/OpenType face.");
        var count = U16(header, 4);
        if (count == 0 || count > 4095 || 12L + count * 16L > length)
            throw new NotSupportedException("Original font directory is invalid.");
        var directory = read(12, count * 16);
        var os2Offset = -1; var headOffset = -1;
        for (var index = 0; index < count; index++)
        {
            var record = index * 16;
            var tag = U32(directory, record);
            if (tag is not (0x4f532f32 or 0x68656164)) continue;
            var offset = U32(directory, record + 8);
            var size = U32(directory, record + 12);
            var minimum = tag == 0x4f532f32 ? 64 : 46;
            if (size < minimum || size > 1024 || offset > length || size > length - offset)
                throw new NotSupportedException("Original font style table is missing or truncated.");
            if (tag == 0x4f532f32)
            {
                if (os2Offset >= 0) throw new NotSupportedException("Duplicate original font style table.");
                os2Offset = checked((int)offset);
            }
            else
            {
                if (headOffset >= 0) throw new NotSupportedException("Duplicate original font style table.");
                headOffset = checked((int)offset);
            }
        }
        if (os2Offset < 0 || headOffset < 0)
            throw new NotSupportedException("Original font style table is missing.");
        var selection = U16(read(os2Offset + 62, 2), 0);
        var macStyle = U16(read(headOffset + 44, 2), 0);
        var bold = (selection & 0x20) != 0; var italic = (selection & 1) != 0;
        if (((macStyle & 1) != 0) != bold || ((macStyle & 2) != 0) != italic)
            throw new NotSupportedException("Original OS/2 and head style flags disagree.");
        return (bold, italic);
    }

    private static ushort U16(byte[] bytes, int offset) =>
        (ushort)((bytes[offset] << 8) | bytes[offset + 1]);

    private static uint U32(byte[] bytes, int offset) =>
        ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) |
        ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];

    private static byte[] ReadRaw(SKStream stream, int offset, int count)
    {
        if (!stream.Seek(offset)) throw new NotSupportedException("Original font stream is not seekable.");
        var bytes = new byte[count]; var read = 0;
        while (read < count)
        {
            var chunk = new byte[count - read];
            var n = stream.Read(chunk, chunk.Length);
            if (n <= 0) throw new NotSupportedException("Original font stream is truncated.");
            Buffer.BlockCopy(chunk, 0, bytes, read, n);
            read += n;
        }
        return bytes;
    }

    private static FontIdentity ReadIdentity(SKTypeface face, int maxFontBytes, byte[]? original, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var style = FaceStyle(face);
        using var stream = face.OpenStream(out var faceIndex);
        if (stream is null || !stream.HasLength || stream.Length <= 0 || stream.Length > maxFontBytes)
            throw new NotSupportedException("Typeface must expose a bounded original font stream.");
        if (faceIndex != 0) throw new NotSupportedException("Collection font faces are not supported.");
        if (original is not null && original.Length != stream.Length)
            throw new InvalidDataException("Typeface stream does not match its original font snapshot.");
        using var hash = SHA256.Create();
        var buffer = new byte[8192]; var remaining = stream.Length; var offset = 0;
        // Detect a collection even if an unusual native stream returns short reads.
        var signature = new byte[4];
        while (remaining > 0)
        {
            token.ThrowIfCancellationRequested();
            var read = stream.Read(buffer, Math.Min(remaining, buffer.Length));
            if (read <= 0) throw new InvalidDataException("Incomplete typeface stream.");
            // Public Text has no original snapshot to compare: inspect only
            // the four-byte collection signature, retaining its streaming SHA.
            var inspect = original is null ? Math.Min(read, Math.Max(0, 4 - offset)) : read;
            for (var i = 0; i < inspect; i++)
            {
                if (offset + i < 4) signature[offset + i] = buffer[i];
                if (original is not null && buffer[i] != original[offset + i])
                    throw new InvalidDataException("Typeface stream does not match its original font snapshot.");
            }
            if (offset < 4 && offset + read >= 4 && signature[0] == 't' && signature[1] == 't' && signature[2] == 'c' && signature[3] == 'f')
                throw new NotSupportedException("Collection fonts are not supported.");
            hash.TransformBlock(buffer, 0, read, buffer, 0); remaining -= read; offset += read;
        }
        token.ThrowIfCancellationRequested();
        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        var rawStyle = RawStyle(stream.Length, (offset, count) => ReadRaw(stream, offset, count));
        if (rawStyle != style)
            throw new NotSupportedException("Native and original font style flags disagree.");
        return new FontIdentity(hash.Hash!, style.Bold, style.Italic);
    }

    /// <summary>Owns one selected original fixed font for an internal producer's page lifetime.</summary>
    internal sealed class FixedFontBinding : IDisposable
    {
        private readonly object _gate = new();
        private readonly byte[] _original;
        private readonly SKTypeface _face;
        private readonly string _resourceId;
        private readonly int _maxFontBytes;
        private FontIdentity? _identity;
        private bool _disposed;
        internal string FamilyName { get; }
        internal bool Bold { get; }
        internal bool Italic { get; }

        private FixedFontBinding(byte[] original, string resourceId, int maxFontBytes)
        {
            if (original is null) throw new ArgumentNullException(nameof(original));
            if (maxFontBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxFontBytes));
            if (original.Length == 0 || original.Length > maxFontBytes)
                throw new InvalidDataException("Bounded original font payload is required.");
            if (string.IsNullOrWhiteSpace(resourceId)) throw new ArgumentException("A destination font resource ID is required.", nameof(resourceId));
            _original = (byte[])original.Clone(); _resourceId = resourceId; _maxFontBytes = maxFontBytes;
            using var data = SKData.CreateCopy(_original);
            _face = SKTypeface.FromData(data) ?? throw new NotSupportedException("PDFV_FONT_FACE");
            try
            {
                // Preserve the producer's early bounded style-table admission.
                if (_face.GetTableSize(0x4f532f32u) > 1024 || _face.GetTableSize(0x68656164u) > 1024)
                    throw new NotSupportedException("PDFV_FONT_STYLE_TABLE");
                var style = FaceStyle(_face);
                var rawStyle = RawStyle(_original.Length, (offset, count) =>
                {
                    var bytes = new byte[count]; Buffer.BlockCopy(_original, offset, bytes, 0, count); return bytes;
                });
                if (rawStyle != style)
                    throw new NotSupportedException("Native and original font style flags disagree.");
                Bold = style.Bold; Italic = style.Italic; FamilyName = _face.FamilyName;
            }
            catch { _face.Dispose(); throw; }
        }

        internal static FixedFontBinding FromOriginalBytes(byte[] original, string resourceId, int maxFontBytes)
            => new(original, resourceId, maxFontBytes);

        internal ushort GetGlyph(int originalUnicode)
        {
            lock (_gate) { ThrowIfDisposed(); return _face.GetGlyph(originalUnicode); }
        }

        internal SkiaDrawEvent SnapshotText(string text, SKPoint baseline, float size, SKPaint paint,
            IReadOnlyList<float> advances, SKMatrix matrix, int maxTextCharacters, CancellationToken token)
        {
            lock (_gate)
            {
                ThrowIfDisposed(); token.ThrowIfCancellationRequested();
                if (maxTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxTextCharacters));
                using var font = new SKFont(_face, size);
                if (font.Size != size) throw new NotSupportedException("TEXT_FLOAT_PRECISION: font changed validated size.");
                var snapshot = new SkiaDrawEvent(text, baseline, font, paint, _resourceId, advances, matrix,
                    maxTextCharacters, _maxFontBytes, () => Identity(token), token);
                token.ThrowIfCancellationRequested();
                return snapshot;
            }
        }

        // Called only while SnapshotText holds _gate and after common per-run
        // validation. A cancelled/failed pass never publishes a reusable digest.
        private FontIdentity Identity(CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (_identity is null)
            {
                ValidateFixedFace(_face);
                var identity = ReadIdentity(_face, _maxFontBytes, _original, token);
                token.ThrowIfCancellationRequested(); _identity = identity;
            }
            return _identity;
        }

        private void ThrowIfDisposed()
        { if (_disposed) throw new ObjectDisposedException(nameof(FixedFontBinding)); }

        public void Dispose()
        {
            lock (_gate)
            {
                if (_disposed) return;
                _disposed = true; _face.Dispose();
            }
        }
    }
}
