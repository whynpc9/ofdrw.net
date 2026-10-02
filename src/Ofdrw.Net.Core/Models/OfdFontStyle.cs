using Ofdrw.Net.Core.Fonts;
namespace Ofdrw.Net.Core.Models;
internal static class OfdFontStyle
{
    internal static (bool Bold, bool Italic) Read(byte[] data) => new OpenTypeFace(data).Style;
}
