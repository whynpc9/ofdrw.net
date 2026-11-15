namespace Ofdrw.Net.Converter.Pdf;

/// <summary>What an exported page image shows where the page paints nothing. Immutable.</summary>
public sealed class OfdImageBackground
{
    private OfdImageBackground(byte red, byte green, byte blue, bool transparent)
    {
        Red = red;
        Green = green;
        Blue = blue;
        IsTransparent = transparent;
    }

    /// <summary>Opaque white, the default.</summary>
    public static OfdImageBackground White { get; } = new(255, 255, 255, false);

    /// <summary>Keeps the page's own transparency. The PNG output has an alpha channel.</summary>
    public static OfdImageBackground Transparent { get; } = new(0, 0, 0, true);

    /// <summary>Creates an opaque background color.</summary>
    public static OfdImageBackground Opaque(byte red, byte green, byte blue) => new(red, green, blue, false);

    /// <summary>True for <see cref="Transparent"/>.</summary>
    public bool IsTransparent { get; }

    /// <summary>Red component of an opaque background.</summary>
    public byte Red { get; }

    /// <summary>Green component of an opaque background.</summary>
    public byte Green { get; }

    /// <summary>Blue component of an opaque background.</summary>
    public byte Blue { get; }
}
