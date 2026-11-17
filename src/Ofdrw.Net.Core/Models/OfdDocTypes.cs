using System;

namespace Ofdrw.Net.Core.Models;

/// <summary>
/// Root <c>DocType</c> declaration values of <c>OFD.xml</c>. A declaration is only a label: it is
/// not evidence that a package conforms to the subset it names.
/// </summary>
public static class OfdDocTypes
{
    /// <summary>The ordinary OFD declaration, <c>OFD</c>.</summary>
    public const string Ordinary = "OFD";

    /// <summary>The archival subset declaration of GB/T 42133, <c>OFD-A</c>.</summary>
    public const string Archival = "OFD-A";

    /// <summary>
    /// Returns whether a declared <c>DocType</c> value denotes the archival subset. The comparison is
    /// exact and ordinal; <see langword="null"/> (a missing attribute) is not archival. Pass the value
    /// actually declared, not a reader fallback substituted for a missing attribute.
    /// </summary>
    /// <param name="declaredDocType">The raw root attribute value, or <see langword="null"/> when absent.</param>
    /// <returns><see langword="true"/> only for the exact value <c>OFD-A</c>.</returns>
    public static bool IsArchival(string? declaredDocType) =>
        string.Equals(declaredDocType, Archival, StringComparison.Ordinal);
}

/// <summary>
/// Thrown when a content-changing operation receives an input whose <c>OFD.xml</c> declares
/// <c>DocType="OFD-A"</c>. The operation fails before producing output and leaves the original
/// unchanged. Create an ordinary OFD derivative first (SDK <c>OfdOrdinaryDerivation.CreateAsync</c> in
/// Ofdrw.Net.Packaging, or CLI <c>ofdrw derive-ofd</c>) and edit the derivative.
/// </summary>
public sealed class OfdArchivalOriginalException : InvalidOperationException
{
    /// <summary>Creates the exception for the named operation.</summary>
    /// <param name="operation">The rejected operation, for example <c>OfdDocumentEditor.RemovePages</c>.</param>
    public OfdArchivalOriginalException(string operation)
        : base(CreateMessage(operation))
    {
        Operation = operation ?? string.Empty;
    }

    /// <summary>The operation that refused the archival original.</summary>
    public string Operation { get; }

    /// <summary>The declaration that triggered the refusal; always <c>OFD-A</c>.</summary>
    public string DeclaredDocType => OfdDocTypes.Archival;

    private static string CreateMessage(string? operation) =>
        $"{(string.IsNullOrEmpty(operation) ? "This operation" : operation)} refused the input: OFD.xml declares DocType=\"OFD-A\" " +
        "(an archival original), and content-changing operations do not modify archival originals. No output was produced. " +
        "To edit, first create an ordinary OFD derivative in a new file with OfdOrdinaryDerivation.CreateAsync " +
        "or 'ofdrw derive-ofd <input.ofd> <output.ofd>', then edit the derivative. " +
        "The derivative does not inherit any archival conclusion or signature validity.";
}

/// <summary>Entry-point guard shared by content-changing operations in base packages.</summary>
internal static class OfdArchivalEditGuard
{
    /// <summary>Throws when the package's declared DocType is OFD-A; null packages are left to the caller's own argument checks.</summary>
    internal static void ThrowIfArchival(OfdDocumentPackage? package, string operation)
    {
        if (package is not null && OfdDocTypes.IsArchival(package.Options?.DocType))
            throw new OfdArchivalOriginalException(operation);
    }

    /// <summary>Throws when a raw root DocType attribute value is OFD-A.</summary>
    internal static void ThrowIfArchival(string? declaredDocType, string operation)
    {
        if (OfdDocTypes.IsArchival(declaredDocType)) throw new OfdArchivalOriginalException(operation);
    }
}
