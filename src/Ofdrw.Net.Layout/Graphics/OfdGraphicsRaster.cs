using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Graphics;

// A deliberately explicit, bounded raster fallback for the rectangle-only
// gradient subset. It produces ordinary RGBA PNG rather than unmodeled OFD
// shading nodes, so every existing image exporter follows the same contract.
internal static class OfdGraphicsRaster
{
    private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
    internal static byte[] LinearGradient(OfdColor start, OfdColor end, int samples, bool horizontal, CancellationToken token)
    {
        if (samples < 2 || samples > 4096) throw new ArgumentOutOfRangeException(nameof(samples));
        var width = horizontal ? samples : 1;
        var height = horizontal ? 1 : samples;
        var scan = new byte[checked(height * (1 + 4 * width))];
        for (var position = 0; position < samples; position++)
        {
            token.ThrowIfCancellationRequested();
            var t = position / (double)(samples - 1);
            var pixel = horizontal ? 1 + 4 * position : position * 5 + 1;
            scan[pixel] = Mix(start.Red, end.Red, t);
            scan[pixel + 1] = Mix(start.Green, end.Green, t);
            scan[pixel + 2] = Mix(start.Blue, end.Blue, t);
            scan[pixel + 3] = Mix(start.Alpha, end.Alpha, t);
        }
        using var compressed = new MemoryStream();
        using (var deflate = new DeflateStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
            deflate.Write(scan, 0, scan.Length);
        using var zlib = new MemoryStream();
        zlib.WriteByte(0x78); zlib.WriteByte(0x9c);
        var raw = compressed.ToArray(); zlib.Write(raw, 0, raw.Length);
        uint a = 1, b = 0;
        foreach (var value in scan) { a = (a + value) % 65521; b = (b + a) % 65521; }
        Write32(zlib, b << 16 | a);
        using var png = new MemoryStream();
        png.Write(Signature, 0, Signature.Length);
        using var header = new MemoryStream();
        Write32(header, (uint)width); Write32(header, (uint)height);
        header.WriteByte(8); header.WriteByte(6); header.WriteByte(0); header.WriteByte(0); header.WriteByte(0);
        Chunk(png, "IHDR", header.ToArray());
        Chunk(png, "IDAT", zlib.ToArray());
        Chunk(png, "IEND", Array.Empty<byte>());
        return png.ToArray();
    }

    private static byte Mix(int start, int end, double t) => (byte)Math.Round(start + (end - start) * t);
    private static void Chunk(Stream output, string name, byte[] payload)
    {
        Write32(output, (uint)payload.Length);
        uint crc = 0xffffffff;
        foreach (var character in name)
        {
            var value = (byte)character; output.WriteByte(value); crc = Crc(crc, value);
        }
        foreach (var value in payload) { output.WriteByte(value); crc = Crc(crc, value); }
        Write32(output, crc ^ 0xffffffff);
    }
    private static uint Crc(uint crc, byte value)
    {
        crc ^= value;
        for (var bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? crc >> 1 ^ 0xedb88320u : crc >> 1;
        return crc;
    }
    private static void Write32(Stream output, uint value)
    {
        output.WriteByte((byte)(value >> 24)); output.WriteByte((byte)(value >> 16));
        output.WriteByte((byte)(value >> 8)); output.WriteByte((byte)value);
    }
}
