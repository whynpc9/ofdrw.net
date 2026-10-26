using System;
using System.IO;
using System.Threading;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Layout.Graphics;

internal static class OfdGraphicsImage
{
    internal static (string MediaType, int Width, int Height) Validate(byte[] data, long maximumPixels, CancellationToken token)
    {
        var format = EncodedImageAdmission.IdentifyFormat(data);
        if (format is not ("image/png" or "image/jpeg"))
            throw new NotSupportedException("Graphics DrawImage accepts only PNG and baseline JPEG.");
        EncodedImageAdmission.Validate(data, format, token, pixels =>
        {
            if (pixels > maximumPixels) throw new InvalidDataException("Graphics image pixel budget exceeded.");
        });
        if (format == "image/png")
            return (format, checked((int)Read32(data, 16)), checked((int)Read32(data, 20)));
        // The framing validator above has already checked the complete baseline
        // bitstream. Read its SOF0 dimensions without a second decoder.
        var width = 0; var height = 0;
        for (var cursor = 2; cursor < data.Length - 9;)
        {
            token.ThrowIfCancellationRequested();
            if (data[cursor++] != 255) break;
            while (cursor < data.Length && data[cursor] == 255) cursor++;
            var marker = data[cursor++];
            var length = data[cursor] * 256 + data[cursor + 1];
            if (marker is 0xe1 or 0xe2 or 0xee)
                throw new NotSupportedException("Graphics JPEG excludes EXIF, ICC and Adobe metadata so pixel axes and color remain stable across exports.");
            if (marker == 0xc0)
            {
                if (data[cursor + 7] is not (1 or 3))
                    throw new NotSupportedException("Graphics JPEG supports gray or three-component baseline data.");
                width = data[cursor + 5] * 256 + data[cursor + 6];
                height = data[cursor + 3] * 256 + data[cursor + 4];
            }
            if (marker == 0xda) break;
            cursor += length;
        }
        if (width > 0 && height > 0) return (format, width, height);
        throw new InvalidDataException("Validated JPEG has no baseline frame dimensions.");
    }

    private static uint Read32(byte[] bytes, int offset) =>
        (uint)bytes[offset] << 24 | (uint)bytes[offset + 1] << 16 | (uint)bytes[offset + 2] << 8 | bytes[offset + 3];
}
