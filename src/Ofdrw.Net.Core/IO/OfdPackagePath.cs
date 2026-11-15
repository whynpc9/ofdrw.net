using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Ofdrw.Net.Core.IO;

internal static class OfdPackagePath
{
    internal static string Resolve(string containingEntry, string reference) =>
        TryResolve(containingEntry, reference, out var resolved)
            ? resolved
            : throw new InvalidDataException("Package reference escapes the OFD root.");

    internal static bool TryResolve(string containingEntry, string reference, out string resolved)
    {
        var value = reference.Replace('\\', '/');
        var combined = value.StartsWith("/", StringComparison.Ordinal)
            ? value : GetDirectory(containingEntry) + "/" + value;
        var parts = new List<string>();
        foreach (var part in combined.Split('/'))
        {
            if (part.Length == 0 || part == ".") continue;
            if (part == "..")
            {
                if (parts.Count == 0)
                {
                    resolved = string.Empty;
                    return false;
                }
                parts.RemoveAt(parts.Count - 1);
            }
            else parts.Add(part);
        }

        resolved = string.Join("/", parts);
        return true;
    }

    internal static string GetDirectory(string entry)
    {
        var index = entry.LastIndexOf('/');
        return index < 0 ? string.Empty : entry.Substring(0, index);
    }

    internal static string RelativeTo(string containingEntry, string target)
    {
        var sourceParts = GetDirectory(containingEntry).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var targetParts = target.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var common = 0;
        while (common < sourceParts.Length && common < targetParts.Length &&
            string.Equals(sourceParts[common], targetParts[common], StringComparison.Ordinal)) common++;
        return string.Join("/", Enumerable.Repeat("..", sourceParts.Length - common).Concat(targetParts.Skip(common)));
    }
}
