namespace Ofdrw.Net.Core.Models;

public enum OfdElementType
{
    Text,
    Image
}

public abstract class OfdElement
{
    public OfdElementType Type { get; protected set; }

    public double XMillimeters { get; set; }

    public double YMillimeters { get; set; }

    public double WidthMillimeters { get; set; }

    public double HeightMillimeters { get; set; }
}

public sealed class OfdTextElement : OfdElement
{
    public OfdTextElement()
    {
        Type = OfdElementType.Text;
    }

    public string Text { get; set; } = string.Empty;

    public string FontName { get; set; } = "SimSun";

    public double FontSizeMillimeters { get; set; } = 4d;
}

public sealed class OfdImageElement : OfdElement
{
    public OfdImageElement()
    {
        Type = OfdElementType.Image;
    }

    public string ResourceId { get; set; } = string.Empty;

    public string FileName { get; set; } = string.Empty;

    public string MediaType { get; set; } = "image/png";

    public byte[] Data { get; set; } = [];
}
