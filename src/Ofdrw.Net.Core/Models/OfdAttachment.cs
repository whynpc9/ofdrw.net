using System;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdAttachment
{
    public string? Id { get; set; }

    public string Name { get; set; } = string.Empty;

    public string MediaType { get; set; } = "application/octet-stream";

    public string? Format { get; set; }

    public DateTimeOffset? CreationDate { get; set; }

    public DateTimeOffset? ModificationDate { get; set; }

    public double? SizeKilobytes { get; set; }

    public bool? Visible { get; set; }

    public string? Usage { get; set; }

    public bool IsExternal { get; set; }

    public string? ExternalPath { get; set; }

    public byte[] Data { get; set; } = [];
}
