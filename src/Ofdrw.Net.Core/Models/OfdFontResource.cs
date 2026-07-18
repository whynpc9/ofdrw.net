namespace Ofdrw.Net.Core.Models;

/// <summary>
/// A font declared by an OFD public resource file.
/// </summary>
public sealed class OfdFontResource
{
    public string Id { get; set; } = string.Empty;

    public string FontName { get; set; } = string.Empty;

    public string? FamilyName { get; set; }

    public string? Charset { get; set; }

    public bool Bold { get; set; }

    public bool Italic { get; set; }

    public string? FileName { get; set; }

    public byte[] Data { get; set; } = [];
}
