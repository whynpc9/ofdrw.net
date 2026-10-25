using System;
using System.Threading;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Converter.Svg.Internal;

internal static class HtmlImageAdmission
{
    internal static void Validate(byte[] data, string mime, CancellationToken token, Action<long>? reservePixels = null)
        => EncodedImageAdmission.Validate(data, mime, token, reservePixels);
}
