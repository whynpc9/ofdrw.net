using System;

namespace Ofdrw.Net.Converter.Pdf.Vector;

// Coordinates are PDF user-space points. The effective crop is strictly inside
// the media box; the display frame is the crop after the PDF quarter turn.
internal readonly struct VectorPageFrame
{
    internal readonly double[] Media;
    internal readonly double[] Crop;
    internal readonly int Rotation;
    internal readonly double Width;
    internal readonly double Height;

    internal VectorPageFrame(double[] media, double[] crop, int rotation)
    {
        Media = media; Crop = crop; Rotation = rotation;
        Width = rotation is 90 or 270 ? crop[3] - crop[1] : crop[2] - crop[0];
        Height = rotation is 90 or 270 ? crop[2] - crop[0] : crop[3] - crop[1];
    }

    internal bool IsLegacyFrame => Rotation == 0 && Media[0] == 0 && Media[1] == 0 &&
        Crop[0] == 0 && Crop[1] == 0 && Crop[2] == Media[2] && Crop[3] == Media[3];

    // The existing producer flips a synthetic page's Y at emission. This
    // initial PDF CTM maps source coordinates into that synthetic page, so its
    // established path/text precision checks still compare the actual values.
    internal double[] InitialTransform => Rotation switch
    {
        0 => new[] { 1d, 0d, 0d, 1d, -Crop[0], -Crop[1] },
        90 => new[] { 0d, -1d, 1d, 0d, -Crop[1], Crop[2] },
        180 => new[] { -1d, 0d, 0d, -1d, Crop[2], Crop[3] },
        270 => new[] { 0d, 1d, -1d, 0d, Crop[3], -Crop[0] },
        _ => throw new InvalidOperationException("Page rotation was not admitted.")
    };
}
