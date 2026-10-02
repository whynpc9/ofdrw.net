using System;
using System.Globalization;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal static class PdfOperandGeometry
{
    private static double Points(double millimeters) => millimeters * 72d / 25.4d;
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    internal static bool Origin(double x, double y) => Finite(Points(x)) && Finite(Points(y));

    internal static bool Box(double x, double y, double width, double height)
    {
        var px = Points(x); var py = Points(y); var pw = Points(width); var ph = Points(height);
        return Finite(px) && Finite(py) && PositiveExtent(pw) && PositiveExtent(ph) &&
            Finite(px + pw) && Finite(py + ph);
    }

    internal static bool PositiveExtent(double value)
        => PositiveSerializedExtent(value, "0.####");

    internal static bool PositiveFormExtent(double value)
        => PositiveSerializedExtent(value, "0.###");

    private static bool PositiveSerializedExtent(double value, string format)
    {
        // PdfSharpCore 1.3.67 writes cm operands with 0.#### and XForm BBox with 0.###.
        // A positive double that serializes as zero would silently erase an appearance.
        return Finite(value) && value > 0 && value.ToString(format, CultureInfo.InvariantCulture) != "0";
    }
}
