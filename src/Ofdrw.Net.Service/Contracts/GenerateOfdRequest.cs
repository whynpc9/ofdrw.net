using System.Collections.Generic;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Service.Contracts;

public sealed class GenerateOfdRequest
{
    public OfdDocumentOptions? Options { get; set; }

    public List<GeneratePageRequest> Pages { get; set; } = [];

    public List<GenerateAttachmentRequest> Attachments { get; set; } = [];

    public Dictionary<string, string> CustomTags { get; set; } = [];
}

public sealed class GeneratePageRequest
{
    public double WidthMillimeters { get; set; }

    public double HeightMillimeters { get; set; }

    public List<GenerateTextElementRequest> Texts { get; set; } = [];

    public List<GenerateImageElementRequest> Images { get; set; } = [];
}

public sealed class GenerateTextElementRequest
{
    public string Text { get; set; } = string.Empty;

    public string FontName { get; set; } = "SimSun";

    public double FontSizeMillimeters { get; set; } = 4d;

    public double XMillimeters { get; set; }

    public double YMillimeters { get; set; }

    public double WidthMillimeters { get; set; }

    public double HeightMillimeters { get; set; }
}

public sealed class GenerateImageElementRequest
{
    public string Base64Data { get; set; } = string.Empty;

    public string MediaType { get; set; } = "image/png";

    public string FileName { get; set; } = string.Empty;

    public double XMillimeters { get; set; }

    public double YMillimeters { get; set; }

    public double WidthMillimeters { get; set; }

    public double HeightMillimeters { get; set; }
}

public sealed class GenerateAttachmentRequest
{
    public string Name { get; set; } = string.Empty;

    public string MediaType { get; set; } = "application/octet-stream";

    public bool IsExternal { get; set; }

    public string? ExternalPath { get; set; }

    public string? Base64Data { get; set; }
}
