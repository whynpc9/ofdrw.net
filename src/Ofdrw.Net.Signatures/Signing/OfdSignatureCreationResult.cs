using System.Collections.Generic;

namespace Ofdrw.Net.Signatures.Signing;

/// <summary>Exact package entry coverage of the newly generated signature.</summary>
public sealed class OfdSignatureCreationResult
{
    internal OfdSignatureCreationResult(
        IReadOnlyList<string> protectedFiles,
        IReadOnlyList<string> unprotectedFiles)
    {
        ProtectedFiles = protectedFiles;
        UnprotectedFiles = unprotectedFiles;
    }

    /// <summary>Absolute paths whose bytes were included in the new signature's References.</summary>
    public IReadOnlyList<string> ProtectedFiles { get; }

    /// <summary>Absolute paths in the final package omitted from the new signature's References.</summary>
    public IReadOnlyList<string> UnprotectedFiles { get; }

    /// <summary>True because the signer refuses a selection that omits required document entries.</summary>
    public bool RequiredFileCoverageValid => true;
}
