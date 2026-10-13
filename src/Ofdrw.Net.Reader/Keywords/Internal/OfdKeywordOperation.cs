using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Reader.Keywords.Internal;

// One operation owns snapshots, work counters and a single cooperating provider session.
// Source models must not be concurrently mutated. Final source comparison catches callback mutations.
internal sealed class OfdKeywordOperation : IDisposable
{
    private readonly Dictionary<OfdFontResource, OfdGlyphFontSnapshot> fonts = new();
    private readonly Dictionary<string, OfdFontResource> identities = new(StringComparer.Ordinal);
    private IOfdGlyphGeometryProvider? provider;
    private IOfdGlyphGeometrySession? session;
    private bool disposed;
    internal OfdModelMaterialization? Material { get; }

    internal OfdKeywordOperation(OfdKeywordSearchOptions? options = null, CancellationToken cancellationToken = default,
        OfdModelMaterialization? material = null)
    {
        Options = options ?? new OfdKeywordSearchOptions(); Token = cancellationToken; Material = material;
        Check();
    }
    internal OfdKeywordSearchOptions Options { get; }
    internal CancellationToken Token { get; }
    internal long CopiedFontBytes { get; private set; }
    internal long HashedFontBytes { get; private set; }
    internal int HashOperations { get; private set; }
    internal long FontComparisonBytes { get; private set; }
    internal long FontReadBytes { get; private set; }
    internal long FontSeekOperations { get; private set; }
    internal long ReservedFontReadBytes { get; private set; }
    private long fontReadCredit;
    private long workCredit;
    internal long LogicalValidationWorkUnits { get; private set; }
    internal long ReservedLogicalValidationWorkUnits { get; private set; }
    internal int FontStreamOpens { get; private set; }
    internal int TextUtf16Units { get; private set; }
    internal int ClusterVisits { get; private set; }
    internal int Hits { get; private set; }
    internal long MatchComparisons { get; private set; }
    internal long CancellationChecks { get; private set; }

    internal OfdGlyphFontSnapshot CaptureFont(OfdFontResource source, string expectedResourceId)
    {
        Check();
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (string.IsNullOrWhiteSpace(expectedResourceId) || !string.Equals(source.Id, expectedResourceId, StringComparison.Ordinal))
            throw new InvalidDataException("Glyph geometry requires the exact explicit font resource ID.");
        if (fonts.TryGetValue(source, out var existing))
        {
            if (existing.ResourceId != expectedResourceId) throw new InvalidDataException("Font identity changed during the operation.");
            return existing;
        }
        if (identities.ContainsKey(expectedResourceId)) throw new InvalidDataException("Ambiguous font resource ID.");
        var bytes = source.Data;
        if (bytes is null || bytes.Length == 0) throw new NotSupportedException("Name-only fonts have no exact embedded geometry identity.");
        var limits = Options.GeometryLimits;
        if (fonts.Count >= limits.MaxFontFaces || bytes.Length > limits.MaxFontBytes ||
            bytes.Length > limits.MaxTotalFontBytes - CopiedFontBytes)
            throw new InvalidDataException("Keyword font snapshot capacity exceeded.");
        var bold = source.Bold; var italic = source.Italic;
        Material?.ChargeCopy(bytes.LongLength + 512, Token);
        Material?.ChargeWork(bytes.LongLength, Token);
        var copy = new byte[bytes.Length];
        using var hash = SHA256.Create();
        HashOperations++;
        for (var start = 0; start < bytes.Length;)
        {
            Check();
            var count = Math.Min(81920, bytes.Length - start);
            Buffer.BlockCopy(bytes, start, copy, start, count);
            CopiedFontBytes += count;
            hash.TransformBlock(copy, start, count, copy, start);
            HashedFontBytes += count; start += count;
        }
        Check(); hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        var digest = BitConverter.ToString(hash.Hash!).Replace("-", string.Empty).ToLowerInvariant();
        Material?.StoreArray(copy);
        var snapshot = new OfdGlyphFontSnapshot(expectedResourceId, bold, italic, copy, digest, ChargeFontStream, ChargeFontRead, ChargeFontSeek);
        fonts.Add(source, snapshot); identities.Add(expectedResourceId, source);
        return snapshot;
    }

    internal IOfdGlyphGeometrySession OpenSession(IOfdGlyphGeometryProvider value)
    {
        Check();
        if (value is null) throw new ArgumentNullException(nameof(value));
        if (session is not null)
        {
            if (!ReferenceEquals(provider, value)) throw new InvalidOperationException("An operation cannot switch glyph providers.");
            return session;
        }
        // Assign ownership before checking cancellation again: a provider may cancel while opening.
        session = value.OpenSession(Options.GeometryLimits, Token)
            ?? throw new InvalidOperationException("The glyph provider returned no session.");
        provider = value;
        Check(); return session;
    }

    internal OfdGlyphGeometryRequest Request(OfdGlyphFontSnapshot font, string text, double size, int weight, bool italic)
    {
        Check();
        if (font is null || !identities.TryGetValue(font.ResourceId, out var source) || !ReferenceEquals(fonts[source], font))
            throw new ArgumentException("Font snapshot belongs to another operation.", nameof(font));
        return new OfdGlyphGeometryRequest(font, text, size, weight, italic);
    }

    internal void ValidateSourceFontsUnchanged()
    {
        Check();
        foreach (var pair in fonts)
        {
            var source = pair.Key; var snapshot = pair.Value; var bytes = source.Data; var original = snapshot.Data;
            if (source.Id != snapshot.ResourceId || source.Bold != snapshot.Bold || source.Italic != snapshot.Italic ||
                bytes is null || bytes.Length != original.Length)
                throw new InvalidDataException("Font source changed after snapshot creation.");
            Material?.ChargeWork(bytes.LongLength, Token);
            for (var index = 0; index < bytes.Length; index++)
            {
                if ((index & 255) == 0) Check();
                FontComparisonBytes++;
                if (bytes[index] != original[index]) throw new InvalidDataException("Font bytes changed after snapshot creation.");
            }
        }
        Check();
    }

    internal void ChargeText(int units, int clusters)
    {
        Check();
        if (units < 0 || units > Options.Limits.MaxTextUtf16Units - TextUtf16Units ||
            clusters < 0 || clusters > Options.Limits.MaxClusterVisits - ClusterVisits)
            throw new InvalidDataException("Keyword source scan capacity exceeded.");
        TextUtf16Units += units; ClusterVisits += clusters;
        Work(checked(4L * units + 32L * clusters));
    }
    internal void ChargeComparison()
    {
        if ((MatchComparisons & 255) == 0) Check();
        if (MatchComparisons >= Options.Limits.MaxMatchComparisons) throw new InvalidDataException("Keyword matching work capacity exceeded.");
        Work(2);
        MatchComparisons++;
    }
    internal void Work(long bytes)
    {
        Check();
        if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
        if (Material is not null && bytes > workCredit)
        {
            var reservation = checked(((bytes - workCredit + 65535) / 65536) * 65536);
            Material.ChargeWork(reservation, Token);
            ReservedLogicalValidationWorkUnits = checked(ReservedLogicalValidationWorkUnits + reservation); workCredit += reservation;
        }
        if (Material is not null) workCredit -= bytes;
        LogicalValidationWorkUnits = checked(LogicalValidationWorkUnits + bytes);
    }
    internal void ChargeHit()
    {
        Check();
        if (Hits >= Options.Limits.MaxHits) throw new InvalidDataException("Keyword hit capacity exceeded.");
        Hits++;
    }
    private void ChargeFontStream()
    {
        Check();
        if (FontStreamOpens >= Options.Limits.MaxFontReadStreams) throw new InvalidDataException("Keyword font stream capacity exceeded.");
        FontStreamOpens++;
        Material?.ChargeCopy(128, Token);
    }
    private void ChargeFontRead(int count)
    {
        Check();
        if (Material is not null && count > fontReadCredit)
        {
            // Font parsers often read individual bytes. Prepaid byte blocks
            // avoid treating every byte as a model reference. No read crosses
            // its prepaid credit, and unused credit is never refunded/reset.
            var reservation = checked(((count - fontReadCredit + 65535) / 65536) * 65536);
            Material.ChargeCopy(reservation, Token); Material.ChargeWork(reservation, Token);
            ReservedFontReadBytes = checked(ReservedFontReadBytes + reservation); fontReadCredit += reservation;
        }
        if (Material is not null) fontReadCredit -= count;
        FontReadBytes = checked(FontReadBytes + count);
    }
    private void ChargeFontSeek() { Work(16); FontSeekOperations = checked(FontSeekOperations + 1); }
    internal void Check()
    {
        if (disposed) throw new ObjectDisposedException(nameof(OfdKeywordOperation));
        CancellationChecks++;
        Token.ThrowIfCancellationRequested();
    }
    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        try { session?.Dispose(); }
        finally
        {
            session = null; provider = null;
            foreach (var font in fonts.Values) font.Release();
            fonts.Clear(); identities.Clear();
        }
    }
}
