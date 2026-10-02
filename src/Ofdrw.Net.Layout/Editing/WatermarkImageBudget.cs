using System;
using System.IO;

namespace Ofdrw.Net.Layout.Editing;

internal static class WatermarkImageBudget
{
    internal static void Validate(byte[] data, string mediaType, long maxPixels)
    {
        if (maxPixels <= 0) throw new ArgumentOutOfRangeException(nameof(maxPixels));
        long width = 0, height = 0;
        if (mediaType == "image/png" && data.Length >= 24 && data[0] == 137 && data[1] == 80 && data[2] == 78 && data[3] == 71 &&
            data[4] == 13 && data[5] == 10 && data[6] == 26 && data[7] == 10 && data[12] == 73 && data[13] == 72 && data[14] == 68 && data[15] == 82)
        {
            width = BigEndian(data, 16, 4); height = BigEndian(data, 20, 4);
        }
        else if (mediaType == "image/jpeg" && data.Length >= 4 && data[0] == 255 && data[1] == 216)
        {
            var index = 2;
            while (index < data.Length)
            {
                if (data[index++] != 255) break;
                while (index < data.Length && data[index] == 255) index++;
                if (index >= data.Length) break;
                var marker = data[index++];
                if (marker is 0xD9 or 0xDA) break;
                if (marker is 0x01 or >= 0xD0 and <= 0xD7) continue;
                if (index + 2 > data.Length) break;
                var length = (int)BigEndian(data, index, 2);
                if (length < 2 || index + length > data.Length) break;
                if (marker is >= 0xC0 and <= 0xCF && marker is not (0xC4 or 0xC8 or 0xCC) && length >= 8)
                {
                    height = BigEndian(data, index + 3, 2); width = BigEndian(data, index + 5, 2); break;
                }
                index += length;
            }
        }
        if (width <= 0 || height <= 0 || width > maxPixels / height)
            throw new InvalidDataException("Watermark image dimensions are invalid or exceed the decoded pixel budget.");
    }

    private static long BigEndian(byte[] data, int offset, int count)
    {
        long value = 0;
        for (var i = 0; i < count; i++) value = (value << 8) | data[offset + i];
        return value;
    }
}
