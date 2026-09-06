namespace Ofdrw.Net.Core.Models;

internal static class OfdFontStyle
{
    // OpenType's head.macStyle flags describe the embedded file's actual style,
    // independently of the appearance requested by the OFD resource.
    internal static (bool Bold, bool Italic) Read(byte[] data)
    {
        if (data.Length < 12) return (false, false);
        var count = (data[4] << 8) | data[5];
        for (var index = 0; index < count && 12L + (index + 1L) * 16 <= data.Length; index++)
        {
            var record = 12 + index * 16;
            if (data[record] != 'h' || data[record + 1] != 'e' || data[record + 2] != 'a' || data[record + 3] != 'd') continue;
            var offset = ((long)data[record + 8] << 24) | ((long)data[record + 9] << 16) | ((long)data[record + 10] << 8) | data[record + 11];
            if (offset + 46 > data.Length) return (false, false);
            var flags = (data[(int)offset + 44] << 8) | data[(int)offset + 45];
            return ((flags & 1) != 0, (flags & 2) != 0);
        }
        return (false, false);
    }
}
