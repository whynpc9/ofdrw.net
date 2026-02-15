namespace Ofdrw.Net.Packaging.Validation;

public sealed class OfdPackageStructureIssue
{
    public string Code { get; set; } = string.Empty;

    public string Message { get; set; } = string.Empty;

    public bool IsError { get; set; }
}
