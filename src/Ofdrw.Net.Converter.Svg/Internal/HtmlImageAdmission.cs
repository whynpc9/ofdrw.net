using System;
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace Ofdrw.Net.Converter.Svg.Internal;

// Encoded framing admission, not an image decoder or a claim of arbitrary image
// validity. Require bounded, complete PNG/JPEG structure and no trailing payload.
// Browser decoding remains fallible; other formats never reach it.
internal static class HtmlImageAdmission
{
    internal static void Validate(byte[] data, string mime, CancellationToken token, Action<long>? reservePixels = null)
    {
        if (mime == "image/png") Png(data, token, reservePixels);
        else if (mime == "image/jpeg") BaselineJpegAdmission.Validate(data, token, reservePixels);
        else throw new NotSupportedException("HTML preview admits only PNG and JPEG images.");
    }
    private static void Png(byte[] data, CancellationToken token, Action<long>? reservePixels)
    {
        byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        if (data.Length < 45) Bad();
        for (var i = 0; i < 8; i++) if (data[i] != signature[i]) Bad();
        var position = 8; var header = false; var pixels = false; var metadata = new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);
        while (position < data.Length)
        {
            token.ThrowIfCancellationRequested();
            if (data.Length - position < 12) Bad();
            var length = Read32(data, position);
            if (length > int.MaxValue || length > data.Length - position - 12) Bad();
            var count = (int)length; var type = position + 4;
            uint crc = 0xffffffff;
            for (var i = type; i < type + 4 + count; i++)
            {
                if ((i & 4095) == 0) token.ThrowIfCancellationRequested();
                crc ^= data[i]; for (var bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xedb88320u : crc >> 1;
            }
            if ((crc ^ 0xffffffff) != Read32(data, type + 4 + count)) Bad();
            var name = new string(new[] { (char)data[type], (char)data[type + 1], (char)data[type + 2], (char)data[type + 3] });
            if (!header && name != "IHDR") Bad();
            switch (name)
            {
                case "IHDR":
                    if (header || count != 13 || Read32(data, type + 4) == 0 || Read32(data, type + 8) == 0 ||
                        Read32(data, type + 4) > 100_000 || Read32(data, type + 8) > 100_000 ||
                        (long)Read32(data, type + 4) * Read32(data, type + 8) > 100_000_000 ||
                        data[type + 14] != 0 || data[type + 15] != 0 || data[type + 16] > 1) Bad();
                    reservePixels?.Invoke(checked((long)Read32(data, type + 4) * Read32(data, type + 8)));
                    token.ThrowIfCancellationRequested();
                    if (data[type + 13] == 3) throw new NotSupportedException("Indexed PNG is outside the HTML preview image profile.");
                    header = true; break;
                case "IDAT": pixels = true; break;
                case "IEND": if (!pixels || count != 0 || position + 12 != data.Length) Bad(); PngPixels(data, token); return;
                // Keep a narrow static profile: no animation, compressed metadata,
                // private critical chunks, embedded XML or arbitrary text payloads.
                case "PLTE": case "tRNS": case "sRGB": case "gAMA": case "cHRM": case "pHYs":
                    if (pixels || !metadata.Add(name)) Bad();
                    var color = data[25];
                    if (name == "PLTE" && (color is 0 or 4 || count == 0 || count > 768 || count % 3 != 0) ||
                        name == "tRNS" && (color == 0 ? count != 2 : color != 2 || count != 6) ||
                        name == "sRGB" && (count != 1 || data[type + 4] > 3) || name == "gAMA" && count != 4 ||
                        name == "cHRM" && count != 32 || name == "pHYs" && (count != 9 || data[type + 12] > 1)) Bad();
                    break;
                default: throw new NotSupportedException("PNG chunk is outside the static HTML image profile: " + name);
            }
            position += count + 12;
        }
        Bad();
    }
    private static void PngPixels(byte[] data, CancellationToken token)
    {
        var depth = data[24]; var color = data[25];
        var channels = color switch { 0 => 1, 2 => 3, 3 => 1, 4 => 2, 6 => 4, _ => 0 };
        if (channels == 0 || !(depth == 8 || depth == 16 || (color is 0 or 3 && depth is 1 or 2 or 4)) || color == 3 && depth == 16) Bad();
        if (data[28] != 0) throw new NotSupportedException("Interlaced PNG is outside the HTML preview image profile.");
        using var joined = new MemoryStream();
        for (var p = 8; p < data.Length;)
        {
            token.ThrowIfCancellationRequested(); var length = checked((int)Read32(data, p));
            if (data[p + 4] == 'I' && data[p + 5] == 'D' && data[p + 6] == 'A' && data[p + 7] == 'T') joined.Write(data, p + 8, length);
            p += length + 12;
        }
        var compressed = joined.ToArray();
        if (compressed.Length < 6 || (compressed[0] & 15) != 8 || compressed[0] >> 4 > 7 ||
            ((compressed[0] << 8) + compressed[1]) % 31 != 0 || (compressed[1] & 32) != 0) Bad();
        using var encoded = new MemoryStream(compressed, 2, compressed.Length - 6, false);
        using var inflate = new DeflateStream(encoded, CompressionMode.Decompress);
        var rowBytes = checked((long)Read32(data, 16) * channels * depth + 7) / 8;
        var height = Read32(data, 20); var buffer = new byte[8192]; uint a = 1, b = 0;
        for (uint row = 0; row < height; row++)
        {
            token.ThrowIfCancellationRequested(); var filter = inflate.ReadByte(); if (filter < 0 || filter > 4) Bad();
            a = (a + (uint)filter) % 65521; b = (b + a) % 65521;
            for (long left = rowBytes; left > 0;)
            {
                token.ThrowIfCancellationRequested(); var got = inflate.Read(buffer, 0, (int)Math.Min(left, buffer.Length)); if (got == 0) Bad();
                for (var i = 0; i < got; i++) { a = (a + buffer[i]) % 65521; b = (b + a) % 65521; } left -= got;
            }
        }
        if (inflate.ReadByte() != -1 || (b << 16 | a) != Read32(compressed, compressed.Length - 4)) Bad();
    }
    private static uint Read32(byte[] b, int offset) => ((uint)b[offset] << 24) | ((uint)b[offset + 1] << 16) | ((uint)b[offset + 2] << 8) | b[offset + 3];
    private static void Bad() => throw new InvalidDataException("Invalid or incomplete embedded preview image.");
}
