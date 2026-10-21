using System;

namespace Ofdrw.Net.Signatures.Signing;

/// <summary>An OFD 2.0 signed extension property.</summary>
public sealed class OfdSignatureParameter
{
    /// <summary>Creates a named extension property with optional type and text value.</summary>
    public OfdSignatureParameter(string name, string value, string? type = null)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Type = type;
    }

    /// <summary>Unique parameter name within this signature.</summary>
    public string Name { get; }
    /// <summary>Text value stored in the signed Signature.xml.</summary>
    public string Value { get; }
    /// <summary>Optional application-defined value type.</summary>
    public string? Type { get; }
}
