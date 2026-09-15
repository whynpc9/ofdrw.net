using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Ofdrw.Net.Core.IO;

internal static class OfdPackagePath
{
    internal static string Resolve(string containingEntry, string reference)
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
                if (parts.Count == 0) throw new InvalidDataException("Package reference escapes the OFD root.");
                parts.RemoveAt(parts.Count - 1);
            }
            else parts.Add(part);
        }

        return string.Join("/", parts);
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
