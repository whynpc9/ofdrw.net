namespace Ofdrw.Net.Core.Models;

public sealed class OfdAttachment
{
    public string Name { get; set; } = string.Empty;

    public string MediaType { get; set; } = "application/octet-stream";

    public bool IsExternal { get; set; }

    public string? ExternalPath { get; set; }

    public byte[] Data { get; set; } = [];
}
