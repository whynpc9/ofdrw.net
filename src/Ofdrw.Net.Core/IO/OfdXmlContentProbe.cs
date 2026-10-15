namespace Ofdrw.Net.Core.IO;

internal static class OfdXmlContentProbe
{
    internal static bool LooksLikeXml(byte[] bytes)
    {
        var index = 0; var width = 1; var little = true;
        if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0 && bytes[3] == 0) { index = 4; width = 4; }
        else if (bytes.Length >= 4 && bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0xFE && bytes[3] == 0xFF) { index = 4; width = 4; little = false; }
        else if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) { index = 2; width = 2; }
        else if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) { index = 2; width = 2; little = false; }
        else if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) index = 3;
        while (index + width <= bytes.Length)
        {
            uint value = 0; for (var part = 0; part < width; part++) value |= (uint)bytes[index + part] << (8 * (little ? part : width - part - 1));
            if (value is not (0x20 or 0x09 or 0x0A or 0x0D)) return value == '<';
            index += width;
        }
        return false;
    }
}
