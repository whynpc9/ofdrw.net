namespace Ofdrw.Net.Layout;

/// <summary>Geometry, font-family request and finite input/output budgets for plain text.</summary>
public sealed class PlainTextOfdOptions
{
    /// <summary>Flow options, copied and validated when constructing the converter. Font names do not embed fonts or select real font metrics.</summary>
    public FlowDocumentOptions Layout { get; } = new();

    /// <summary>Maximum UTF-8 bytes read, including a leading BOM; defaults to 4,000,000. Does not apply to strings.</summary>
    public long MaxInputBytes { get; set; } = 4_000_000;
}
