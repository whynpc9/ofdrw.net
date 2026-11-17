using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>
/// Resource-table <c>BaseLoc</c> semantics shared by the local preflight and the archival font checks: a relative
/// payload such as <c>FontFile</c> resolves against the table's BaseLoc directory, itself relative to the table entry;
/// a leading slash is the package root. The caller rejects external, ambiguous or namespaced values first.
/// </summary>
internal static class OfdResourceBase
{
    /// <summary>A synthetic containing entry inside the BaseLoc directory; throws InvalidDataException on root escape.</summary>
    internal static string Containing(string resourceEntry, string? baseLoc) =>
        OfdPackagePath.Resolve(resourceEntry, Probe(baseLoc));

    /// <summary>Resolves a payload reference declared in a resource table; false on root escape.</summary>
    internal static bool TryResolve(string resourceEntry, string? baseLoc, string reference, out string target)
    {
        target = string.Empty;
        return OfdPackagePath.TryResolve(resourceEntry, Probe(baseLoc), out var containing) &&
            OfdPackagePath.TryResolve(containing, reference, out target) && target.Length > 0;
    }

    private static string Probe(string? baseLoc) =>
        string.IsNullOrEmpty(baseLoc) ? "_payload_" : baseLoc!.TrimEnd('/') + "/_payload_";
}
