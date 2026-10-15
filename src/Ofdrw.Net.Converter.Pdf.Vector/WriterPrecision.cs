using Ofdrw.Net.Layout.Graphics;
using SkiaSharp;

namespace Ofdrw.Net.Converter.Pdf.Vector;

internal static class WriterPrecision
{
    internal const double Units = 25.4 / 72;
    // Use the same double multiplication and operand order as SkiaDrawEvent.Emit.
    internal static OfdMatrix Transform(SKMatrix value) => OfdMatrix.Scale(Units, Units).Multiply(
        new OfdMatrix(value.ScaleX, value.SkewY, value.SkewX, value.ScaleY, value.TransX, value.TransY));
}
