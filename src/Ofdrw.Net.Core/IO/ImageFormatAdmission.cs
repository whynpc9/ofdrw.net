using System.IO;

namespace Ofdrw.Net.Core.IO;

// Inspect only the already bounded encoded buffer, without invoking a decoder or advancing a stream.
internal static class ImageFormatAdmission
{
    internal static void RequirePngOrJpeg(byte[] data, int length)
    {
        var png = length >= 8 && data[0] == 137 && data[1] == 80 && data[2] == 78 && data[3] == 71 &&
            data[4] == 13 && data[5] == 10 && data[6] == 26 && data[7] == 10;
        var jpeg = length >= 3 && data[0] == 255 && data[1] == 216 && data[2] == 255;
        if (!png && !jpeg) throw new InvalidDataException("Only PNG and JPEG encoded images are supported.");
    }

    internal static void RejectBigTiff(byte[] data)
    {
        // ImageSharp 2.1.13 can enter a non-progressing BigTIFF IFD loop even during Identify.
        // Reject both byte orders before any generic image operation; classic TIFF remains supported.
        if (data.Length >= 4 &&
            ((data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2b && data[3] == 0) ||
             (data[0] == 0x4d && data[1] == 0x4d && data[2] == 0 && data[3] == 0x2b)))
            throw new InvalidDataException("BigTIFF encoded images are not supported.");
    }
}
