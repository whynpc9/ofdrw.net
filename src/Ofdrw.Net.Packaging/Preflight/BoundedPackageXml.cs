using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>Thrown when a bounded package read exhausts one of the caller's budgets.</summary>
internal sealed class PackageBudgetExceededException : Exception
{
    internal PackageBudgetExceededException(string budget) : base(budget + " budget exceeded.") { Budget = budget; }

    /// <summary>Stable budget name: <c>XmlBytes</c>, <c>XmlNodes</c> or <c>Reference</c>.</summary>
    internal string Budget { get; }
}

/// <summary>
/// The archive beneath a bounded package read. Two adapters exist: an already-loaded <see cref="OfdPackageArchive"/>
/// and the Medical package's own central-directory view, which can report an entry's size before inflating it.
/// </summary>
internal interface IBoundedEntrySource
{
    /// <summary>Declared uncompressed length, or -1 when the entry is absent.</summary>
    long Length(string entry);

    /// <summary>The whole entry. Called only after <see cref="Length"/> admitted it.</summary>
    byte[] ReadAll(string entry);
}

/// <summary>Adapter over an in-memory <see cref="OfdPackageArchive"/>.</summary>
internal sealed class ArchiveEntrySource : IBoundedEntrySource
{
    private readonly OfdPackageArchive archive;
    internal ArchiveEntrySource(OfdPackageArchive archive) { this.archive = archive; }
    public long Length(string entry) => archive.TryGetBytes(entry, out var bytes) ? bytes.Length : -1;
    public byte[] ReadAll(string entry) => archive.GetBytes(entry);
}

/// <summary>How XML nodes are charged against <c>MaxXmlNodes</c>.</summary>
internal enum XmlNodeCharging
{
    /// <summary>One unit per reader node.</summary>
    Nodes,
    /// <summary>One unit per reader node plus one per attribute.</summary>
    NodesAndAttributes
}

/// <summary>Why a bounded read produced no document.</summary>
internal enum BoundedXmlStatus { Parsed, Absent, Malformed, UnsupportedEncoding }

/// <summary>The outcome of one bounded XML read.</summary>
internal readonly struct BoundedXmlResult
{
    internal BoundedXmlResult(BoundedXmlStatus status, XDocument? document) { Status = status; Document = document; }
    internal BoundedXmlStatus Status { get; }
    internal XDocument? Document { get; }
}

/// <summary>
/// Bounded, DTD-free XML access to one package snapshot. It owns byte, node, depth and reference charging, the
/// per-entry cache (case-insensitive, the same identity as entry lookup) and the "charge each distinct entry's bytes
/// once" rule. Every consumer creates its own instance, so each keeps an independent count against the same limits.
/// Budget overruns throw <see cref="PackageBudgetExceededException"/>.
/// </summary>
internal sealed class BoundedPackageXml
{
    private readonly IBoundedEntrySource source;
    private readonly ArchivalPreflightOptions budgets;
    private readonly CancellationToken token;
    private readonly XmlNodeCharging charging;
    private readonly LoadOptions loadOptions;
    private readonly Dictionary<string, BoundedXmlResult> cache = new Dictionary<string, BoundedXmlResult>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> charged = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private long xmlBytes;
    private int nodes, references;

    internal BoundedPackageXml(IBoundedEntrySource source, ArchivalPreflightOptions budgets, CancellationToken token,
        XmlNodeCharging charging = XmlNodeCharging.Nodes, LoadOptions loadOptions = LoadOptions.None)
    {
        this.source = source ?? throw new ArgumentNullException(nameof(source));
        this.budgets = budgets ?? throw new ArgumentNullException(nameof(budgets));
        this.token = token; this.charging = charging; this.loadOptions = loadOptions;
    }

    internal BoundedPackageXml(OfdPackageArchive archive, ArchivalPreflightOptions budgets, CancellationToken token,
        XmlNodeCharging charging = XmlNodeCharging.Nodes, LoadOptions loadOptions = LoadOptions.None)
        : this(new ArchiveEntrySource(archive), budgets, token, charging, loadOptions) { }

    internal ArchivalPreflightOptions Budgets => budgets;

    /// <summary>True when the entry exists; nothing is charged.</summary>
    internal bool Exists(string entry) => source.Length(entry) >= 0;

    /// <summary>Counts one reference visit against <c>MaxReferences</c>; throws when exhausted.</summary>
    internal void Visit()
    {
        token.ThrowIfCancellationRequested();
        if (!TryVisit()) throw new PackageBudgetExceededException("Reference");
    }

    /// <summary>Counts one reference visit; false (and nothing thrown) when <c>MaxReferences</c> is exhausted.</summary>
    internal bool TryVisit() => ++references <= budgets.MaxReferences;

    /// <summary>
    /// Whole entry bytes within the single-entry budget; null when absent. A distinct entry's bytes are charged once,
    /// so a later lexical re-read is free.
    /// </summary>
    internal byte[]? Bytes(string entry)
    {
        token.ThrowIfCancellationRequested();
        var length = source.Length(entry);
        if (length < 0) return null;
        if (length > budgets.MaxSingleXmlBytes) throw new PackageBudgetExceededException("XmlBytes");
        if (!charged.Contains(entry))
        {
            if (length > budgets.MaxXmlBytes - xmlBytes) throw new PackageBudgetExceededException("XmlBytes");
            xmlBytes += length;
            charged.Add(entry);
        }
        return source.ReadAll(entry);
    }

    /// <summary>The parsed document, or null when the entry is absent or not well-formed.</summary>
    internal XDocument? Read(string entry) => Load(entry).Document;

    /// <summary>Parses an entry (cached), distinguishing absent, malformed and unsupported-encoding input.</summary>
    internal BoundedXmlResult Load(string entry)
    {
        token.ThrowIfCancellationRequested();
        if (cache.TryGetValue(entry, out var cached)) return cached;
        var data = Bytes(entry);
        BoundedXmlResult result;
        if (data == null) result = new BoundedXmlResult(BoundedXmlStatus.Absent, null);
        else
        {
            var settings = Settings();
            try
            {
                // Charge nodes and depth while streaming, before an untrusted tree is allocated.
                using (var stream = new MemoryStream(data, false))
                using (var reader = XmlReader.Create(stream, settings))
                    while (reader.Read())
                    {
                        token.ThrowIfCancellationRequested();
                        var cost = charging == XmlNodeCharging.NodesAndAttributes ? 1 + reader.AttributeCount : 1;
                        if (cost > budgets.MaxXmlNodes - nodes || reader.Depth > budgets.MaxXmlDepth)
                            throw new PackageBudgetExceededException("XmlNodes");
                        nodes += cost;
                    }
                using (var stream = new MemoryStream(data, false))
                using (var reader = XmlReader.Create(stream, settings))
                    result = new BoundedXmlResult(BoundedXmlStatus.Parsed, XDocument.Load(reader, loadOptions));
            }
            catch (XmlException ex)
            {
                var encoding = ex.InnerException is ArgumentException || ex.InnerException is NotSupportedException;
                result = new BoundedXmlResult(encoding ? BoundedXmlStatus.UnsupportedEncoding : BoundedXmlStatus.Malformed, null);
            }
        }
        cache[entry] = result;
        return result;
    }

    /// <summary>Reader settings for a lexical pass over bytes obtained from <see cref="Bytes"/>.</summary>
    internal XmlReaderSettings Settings() => new XmlReaderSettings
    {
        DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
        MaxCharactersInDocument = budgets.MaxSingleXmlBytes, MaxCharactersFromEntities = 1
    };
}

/// <summary>
/// The single "safe package reference" policy used by read-only assessment. A safe reference is package-relative:
/// no scheme or drive (<c>:</c>), query, fragment, percent escape, backslash, network path or control character.
/// </summary>
internal static class OfdPackageReference
{
    internal const string ShortNamespace = "http://www.ofdspec.org";
    internal const string Namespace2016 = "http://www.ofdspec.org/2016";

    /// <summary>True for the two OFD namespaces read-only assessment understands.</summary>
    internal static bool SupportedNamespace(XNamespace ns) =>
        ns.NamespaceName == ShortNamespace || ns.NamespaceName == Namespace2016;

    /// <summary>True when the value is empty or not a safe package-relative reference.</summary>
    internal static bool Unsafe(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference)) return true;
        var value = reference!.Trim();
        return value.IndexOfAny(new[] { ':', '?', '#', '%', '\\' }) >= 0 ||
            value.StartsWith("//", StringComparison.Ordinal) || value.Any(char.IsControl);
    }

    /// <summary>Resolves a safe reference against its containing entry; a leading slash is the package root.</summary>
    internal static bool TryResolve(string containingEntry, string? reference, out string resolved) =>
        Normalize(containingEntry, reference, false, out resolved);

    /// <summary>Resolves a safe directory reference such as a resource BaseLoc; the package root is the empty string.</summary>
    internal static bool TryResolveDirectory(string containingEntry, string? reference, out string directory) =>
        Normalize(containingEntry, reference, true, out directory);

    private static bool Normalize(string containingEntry, string? reference, bool allowRoot, out string resolved)
    {
        resolved = string.Empty;
        if (Unsafe(reference)) return false;
        if (!OfdPackagePath.TryResolve(containingEntry, reference!.Trim(), out resolved)) return false;
        return allowRoot || resolved.Length > 0;
    }

    /// <summary>A stable, bounded element selector for evidence locations.</summary>
    internal static string Selector(XElement node, Func<XElement, int>? ordinal = null)
    {
        var parts = node.AncestorsAndSelf().Reverse().Select(item =>
            "{" + (SupportedNamespace(item.Name.Namespace) ? item.Name.NamespaceName : "<unsupported-namespace>") + "}" +
            item.Name.LocalName + "[" + (ordinal?.Invoke(item) ?? item.ElementsBeforeSelf(item.Name).Count() + 1)
                .ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
        var selector = "/" + string.Join("/", parts);
        return selector.Length <= 8192 ? selector : "<selector exceeds metadata limit>";
    }
}
