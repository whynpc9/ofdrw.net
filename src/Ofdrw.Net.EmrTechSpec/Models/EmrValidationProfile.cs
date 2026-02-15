using System.Collections.Generic;
using Ofdrw.Net.Core.Validation;

namespace Ofdrw.Net.EmrTechSpec.Models;

public sealed class EmrValidationProfile
{
    public string Id { get; set; } = string.Empty;

    public string Version { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    public List<EmrValidationRule> Rules { get; set; } = [];
}

public sealed class EmrValidationRule
{
    public string RuleId { get; set; } = string.Empty;

    public string Clause { get; set; } = string.Empty;

    public ValidationSeverity Severity { get; set; }

    public string Summary { get; set; } = string.Empty;

    public string Recommendation { get; set; } = string.Empty;
}
