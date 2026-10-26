using System;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.IO;

/// <summary>
/// Admits only one-page, sequential, standalone JBIG2 files containing a page
/// information segment and one immediate generic region. Dictionaries, globals,
/// refinement, unknown page height and segment references are outside this profile.
/// This is an allocation preflight; the decoder must still be treated as fallible.
/// </summary>
internal static class RestrictedJbig2Admission
{
    private static readonly byte[] Signature = { 0x97, 0x4a, 0x42, 0x32, 0x0d, 0x0a, 0x1a, 0x0a };
    internal const int MaxEncodedBytes = 1_048_576;
    internal const int MaxPixels = 1_048_576;

    internal static bool HasSignature(byte[] data)
    {
        if (data is null || data.Length < Signature.Length) return false;
        for (var i = 0; i < Signature.Length; i++) if (data[i] != Signature[i]) return false;
        return true;
    }

    internal static (int Width, int Height) Validate(byte[] data, CancellationToken token)
    {
        if (!HasSignature(data) || data.Length > MaxEncodedBytes || data.Length < 13 || data[8] != 1 || Read32(data, 9) != 1)
            throw new NotSupportedException("JBIG2 is outside the bounded standalone generic-region profile.");
        var position = 13;
        var width = 0; var height = 0;
        for (var segment = 0; segment < 4; segment++)
        {
            token.ThrowIfCancellationRequested();
            if (data.Length - position < 11 || Read32(data, position) != (uint)segment || data[position + 5] != 0)
                throw new InvalidDataException("Invalid JBIG2 segment header.");
            var type = data[position + 4];
            var expected = segment == 0 ? 48 : segment == 1 ? 38 : segment == 2 ? 49 : 51;
            if (type != expected || data[position + 6] != (segment == 3 ? 0 : 1))
                throw new NotSupportedException("Unsupported JBIG2 segment arrangement.");
            var length = Read32(data, position + 7);
            if (length > data.Length - position - 11) throw new InvalidDataException("Truncated JBIG2 segment.");
            position += 11;
            if (segment == 0)
            {
                if (length != 19) throw new NotSupportedException("Unsupported JBIG2 page information.");
                width = Dimension(Read32(data, position)); height = Dimension(Read32(data, position + 4));
                if ((long)width * height > MaxPixels || data[position + 16] != 1)
                    throw new NotSupportedException("JBIG2 page exceeds the bounded one-page profile.");
            }
            else if (segment == 1)
            {
                if (length < 26 || Read32(data, position) != width || Read32(data, position + 4) != height ||
                    Read32(data, position + 8) != 0 || Read32(data, position + 12) != 0 || data[position + 16] != 0 ||
                    data[position + 17] != 0)
                    throw new NotSupportedException("Unsupported JBIG2 generic region.");
            }
            else if (length != 0) throw new InvalidDataException("Invalid JBIG2 ending segment.");
            position += (int)length;
        }
        if (position != data.Length) throw new InvalidDataException("Trailing JBIG2 payload.");
        return (width, height);
    }

    private static int Dimension(uint value)
    {
        if (value == 0 || value > MaxPixels) throw new NotSupportedException("Invalid JBIG2 dimensions.");
        return (int)value;
    }

    private static uint Read32(byte[] data, int offset) =>
        (uint)data[offset] << 24 | (uint)data[offset + 1] << 16 | (uint)data[offset + 2] << 8 | data[offset + 3];
}
