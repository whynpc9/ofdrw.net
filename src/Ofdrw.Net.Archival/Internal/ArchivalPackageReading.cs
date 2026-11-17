using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival.Internal;

/// <summary>Thrown when this package's own inspection exhausts a Preflight budget.</summary>
internal sealed class ArchivalBudgetExceededException : Exception
{
    internal ArchivalBudgetExceededException(string budget) : base(budget + " budget exceeded.") { Budget = budget; }
    internal string Budget { get; }
}

/// <summary>
/// Bounded XML access shared by the scope inspection and package-owned clause checks. It applies the caller's
/// ArchivalPreflightOptions independently of the Preflight services' own accounting, so both stay within budget.
/// </summary>
internal sealed class ArchivalXmlReader
{
    private readonly OfdPackageArchive archive;
    private readonly ArchivalPreflightOptions options;
    private readonly CancellationToken token;
    private readonly Dictionary<string, XDocument?> cache = new Dictionary<string, XDocument?>(StringComparer.OrdinalIgnoreCase);
    private long bytes;
    private int nodes, references;

    internal ArchivalXmlReader(OfdPackageArchive archive, ArchivalPreflightOptions options, CancellationToken token)
    { this.archive = archive; this.options = options; this.token = token; }

    internal OfdPackageArchive Archive => archive;

    /// <summary>Counts one reference visit against MaxReferences.</summary>
    internal void Visit()
    {
        token.ThrowIfCancellationRequested();
        if (++references > options.MaxReferences) throw new ArchivalBudgetExceededException("Reference");
    }

    /// <summary>Returns null when the entry is absent or not well-formed XML (DTD prohibited).</summary>
    internal XDocument? Read(string entry)
    {
        token.ThrowIfCancellationRequested();
        if (cache.TryGetValue(entry, out var cached)) return cached;
        if (!archive.TryGetBytes(entry, out var data)) { cache[entry] = null; return null; }
        if (data.Length > options.MaxSingleXmlBytes || data.Length > options.MaxXmlBytes - bytes)
            throw new ArchivalBudgetExceededException("XmlBytes");
        bytes += data.Length;
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = options.MaxSingleXmlBytes, MaxCharactersFromEntities = 1
        };
        XDocument? document;
        try
        {
            // Count before building the tree, as the Preflight checker does.
            using (var source = new MemoryStream(data, false))
            using (var reader = XmlReader.Create(source, settings))
                while (reader.Read())
                {
                    token.ThrowIfCancellationRequested();
                    if (++nodes > options.MaxXmlNodes || reader.Depth > options.MaxXmlDepth)
                        throw new ArchivalBudgetExceededException("XmlNodes");
                }
            using (var source = new MemoryStream(data, false))
            using (var reader = XmlReader.Create(source, settings)) document = XDocument.Load(reader);
        }
        catch (XmlException) { document = null; }
        cache[entry] = document;
        return document;
    }
}

internal static class ArchivalPackagePath
{
    internal const string OfdNamespace = "http://www.ofdspec.org";
    internal const string OfdNamespace2016 = "http://www.ofdspec.org/2016";

    internal static bool Supported(XNamespace ns) =>
        ns.NamespaceName == OfdNamespace || ns.NamespaceName == OfdNamespace2016;

    /// <summary>Rejects schemes, UNC, queries, fragments, escapes and backslashes, like the Preflight services.</summary>
    internal static bool Unsafe(string? path) => string.IsNullOrWhiteSpace(path) || path!.Contains(":") ||
        path.Contains("?") || path.Contains("#") || path.Contains("%") || path.Contains("\\") ||
        path.StartsWith("//", StringComparison.Ordinal) || path.Any(char.IsControl);

    /// <summary>Resolves a package-relative reference against its containing entry; a leading slash is package root.</summary>
    internal static bool TryResolve(string containingEntry, string reference, out string resolved)
    {
        var value = reference.Trim();
        var index = containingEntry.LastIndexOf('/');
        var directory = index < 0 ? string.Empty : containingEntry.Substring(0, index);
        var combined = value.StartsWith("/", StringComparison.Ordinal) ? value : directory + "/" + value;
        var parts = new List<string>();
        foreach (var part in combined.Split('/'))
        {
            if (part.Length == 0 || part == ".") continue;
            if (part == "..")
            {
                if (parts.Count == 0) { resolved = string.Empty; return false; }
                parts.RemoveAt(parts.Count - 1);
            }
            else parts.Add(part);
        }
        resolved = string.Join("/", parts);
        return resolved.Length > 0;
    }

    internal static string Selector(XElement node)
    {
        var parts = node.AncestorsAndSelf().Reverse().Select(item =>
            "{" + (Supported(item.Name.Namespace) ? item.Name.NamespaceName : "<unsupported-namespace>") + "}" +
            item.Name.LocalName + "[" + (item.ElementsBeforeSelf(item.Name).Count() + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
        var selector = "/" + string.Join("/", parts);
        return selector.Length <= 8192 ? selector : "<selector exceeds metadata limit>";
    }
}
