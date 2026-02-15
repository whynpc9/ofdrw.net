using System.Collections.Generic;

namespace Ofdrw.Net.Core.Validation;

public enum ValidationSeverity
{
    Info,
    Warning,
    Error
}

public sealed class ValidationFinding
{
    public string RuleId { get; set; } = string.Empty;

    public string Clause { get; set; } = string.Empty;

    public ValidationSeverity Severity { get; set; }

    public string Message { get; set; } = string.Empty;

    public string Location { get; set; } = string.Empty;

    public string Recommendation { get; set; } = string.Empty;
}

public sealed class ValidationEvidence
{
    public string Key { get; set; } = string.Empty;

    public string Value { get; set; } = string.Empty;
}

public sealed class ValidationReport
{
    public string ProfileVersion { get; set; } = string.Empty;

    public bool Passed { get; set; }

    public List<ValidationFinding> Findings { get; set; } = [];

    public List<ValidationEvidence> Evidence { get; set; } = [];
}
