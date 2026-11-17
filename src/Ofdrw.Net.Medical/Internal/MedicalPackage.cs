using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Medical.Internal;

internal sealed class MedicalBudgetExceededException : Exception
{
    internal MedicalBudgetExceededException(string budget) : base(budget + " budget exceeded.") { Budget = budget; }
    internal string Budget { get; }
}

internal sealed class MedicalInputRejectedException : Exception
{
    internal MedicalInputRejectedException(string reason) : base(reason) { }
}

/// <summary>One central-directory record, in central-directory order.</summary>
internal sealed class MedicalZipEntry
{
    internal string Name = string.Empty;
    internal int Method, VersionNeeded, Flags;
    internal long CompressedSize, UncompressedSize, LocalOffset;
    internal bool Encrypted => (Flags & 1) != 0;
}

/// <summary>
/// Bounded read-only view of an OFD snapshot: its own central-directory parse (compression method and physical order,
/// which System.IO.Compression does not expose), entry bytes and DTD-free XML.
/// </summary>
internal sealed class MedicalPackage : IDisposable
{
    internal const string Namespace2016 = "http://www.ofdspec.org/2016";
    internal const string ShortNamespace = "http://www.ofdspec.org";

    private readonly ZipArchive zip;
    private readonly Dictionary<string, ZipArchiveEntry> byName;
    private readonly ArchivalPreflightOptions budgets;
    private readonly CancellationToken token;
    private readonly Dictionary<string, XDocument?> cache = new Dictionary<string, XDocument?>(StringComparer.OrdinalIgnoreCase);
    private long xmlBytes;
    private readonly HashSet<string> charged = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // same identity as byName
    private int xmlNodes, references;

    internal MedicalPackage(byte[] bytes, ArchivalPreflightOptions budgets, CancellationToken token)
    {
        this.budgets = budgets; this.token = token;
        Directory = ReadCentralDirectory(bytes, budgets.LoadOptions.MaxEntryCount);
        // The same ZIP loading limits the embedded archival assessment applies, enforced before any entry is opened.
        var load = budgets.LoadOptions; long total = 0;
        foreach (var entry in Directory)
        {
            if (entry.UncompressedSize > load.MaxEntryUncompressedBytes) throw new MedicalBudgetExceededException("EntryUncompressedBytes");
            total += entry.UncompressedSize;
            if (total > load.MaxTotalUncompressedBytes) throw new MedicalBudgetExceededException("TotalUncompressedBytes");
            if (entry.UncompressedSize > 0 && entry.UncompressedSize > load.MaxCompressionRatio * Math.Max(1, entry.CompressedSize))
                throw new MedicalBudgetExceededException("CompressionRatio");
        }
        try { zip = new ZipArchive(new MemoryStream(bytes, false), ZipArchiveMode.Read); }
        catch (InvalidDataException) { throw new MedicalInputRejectedException("MalformedZip"); }
        byName = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries)
        {
            if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
            var name = entry.FullName.TrimStart('/');
            if (byName.ContainsKey(name) || name.Contains("\\") || name.Split('/').Any(x => x == ".."))
                throw new MedicalInputRejectedException("UnsafeOrDuplicateEntry");
            byName.Add(name, entry);
        }
    }

    /// <summary>Central-directory records of file entries (directories excluded).</summary>
    internal IReadOnlyList<MedicalZipEntry> Directory { get; }

    /// <summary>File entries ordered by local header offset: the physical order a streaming reader sees.</summary>
    internal IReadOnlyList<MedicalZipEntry> PhysicalOrder => Directory.OrderBy(x => x.LocalOffset).ToList();

    internal IEnumerable<string> EntryNames => byName.Keys;

    internal bool Exists(string name) => byName.ContainsKey(name);

    internal long Length(string name) => byName.TryGetValue(name, out var entry) ? entry.Length : -1;

    internal void Visit()
    {
        token.ThrowIfCancellationRequested();
        if (++references > budgets.MaxReferences) throw new MedicalBudgetExceededException("Reference");
    }

    /// <summary>Reads at most <paramref name="max"/> leading bytes of an entry; null when absent.</summary>
    internal byte[]? Head(string name, int max)
    {
        token.ThrowIfCancellationRequested();
        if (!byName.TryGetValue(name, out var entry)) return null;
        using var stream = entry.Open();
        var buffer = new byte[Math.Min(max, (int)Math.Min(int.MaxValue, Math.Max(0, entry.Length)))];
        var read = 0;
        while (read < buffer.Length)
        {
            var count = stream.Read(buffer, read, buffer.Length - read);
            if (count == 0) break;
            read += count;
        }
        return read == buffer.Length ? buffer : buffer.Take(read).ToArray();
    }

    /// <summary>Reads a whole entry within the single-XML budget; null when absent.</summary>
    internal byte[]? Bytes(string name)
    {
        if (!byName.TryGetValue(name, out var entry)) return null;
        if (entry.Length > budgets.MaxSingleXmlBytes) throw new MedicalBudgetExceededException("XmlBytes");
        // MaxXmlBytes bounds the distinct XML read; re-reading an entry (e.g. the MH-003 lexical pass) is not charged again.
        if (!charged.Contains(name))
        {
            if (entry.Length > budgets.MaxXmlBytes - xmlBytes) throw new MedicalBudgetExceededException("XmlBytes");
            xmlBytes += entry.Length;
            charged.Add(name);
        }
        var bytes = Head(name, (int)entry.Length) ?? Array.Empty<byte>();
        if (bytes.Length != entry.Length) throw new MedicalInputRejectedException("TruncatedEntry");
        return bytes;
    }

    /// <summary>Returns the parsed XML, or null when the entry is absent or not well-formed (DTD prohibited).</summary>
    internal XDocument? Xml(string name)
    {
        if (cache.TryGetValue(name, out var cached)) return cached;
        var data = Bytes(name);
        XDocument? document = null;
        if (data != null)
        {
            var settings = Settings();
            try
            {
                using (var reader = XmlReader.Create(new MemoryStream(data, false), settings))
                    while (reader.Read())
                    {
                        token.ThrowIfCancellationRequested();
                        if (++xmlNodes > budgets.MaxXmlNodes || reader.Depth > budgets.MaxXmlDepth)
                            throw new MedicalBudgetExceededException("XmlNodes");
                    }
                using (var reader = XmlReader.Create(new MemoryStream(data, false), settings))
                    document = XDocument.Load(reader);
            }
            catch (XmlException) { document = null; }
        }
        cache[name] = document;
        return document;
    }

    /// <summary>Lexical XML facts XLinq does not keep: OFD-namespace elements not written with the <c>ofd</c> prefix, and qualified attributes.</summary>
    internal (bool WellFormed, int NonOfdPrefixedElements, int QualifiedOfdAttributes) Lexical(string name, string ofdNamespace)
    {
        var data = Bytes(name);
        if (data == null) return (false, 0, 0);
        int unprefixed = 0, qualified = 0;
        try
        {
            using var reader = XmlReader.Create(new MemoryStream(data, false), Settings());
            while (reader.Read())
            {
                token.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.NamespaceURI == ofdNamespace && reader.Prefix != "ofd") unprefixed++;
                if (reader.MoveToFirstAttribute())
                {
                    do
                        if (reader.NamespaceURI == ofdNamespace && reader.Prefix.Length > 0) qualified++;
                    while (reader.MoveToNextAttribute());
                    reader.MoveToElement();
                }
            }
            return (true, unprefixed, qualified);
        }
        catch (XmlException) { return (false, 0, 0); }
    }

    private XmlReaderSettings Settings() => new XmlReaderSettings
    {
        DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
        MaxCharactersInDocument = budgets.MaxSingleXmlBytes, MaxCharactersFromEntities = 1
    };

    public void Dispose() => zip.Dispose();

    /// <summary>Resolves an OFD location against the entry that contains it; a leading slash is the package root.</summary>
    internal static bool TryResolve(string containingEntry, string? location, out string resolved) =>
        Normalize(containingEntry, location, false, out resolved);

    /// <summary>Resolves a directory location such as a resource BaseLoc; the package root is the empty string.</summary>
    internal static bool TryResolveDirectory(string containingEntry, string? location, out string directory) =>
        Normalize(containingEntry, location, true, out directory);

    private static bool Normalize(string containingEntry, string? location, bool allowRoot, out string resolved)
    {
        resolved = string.Empty;
        if (string.IsNullOrWhiteSpace(location)) return false;
        var value = location!.Trim();
        if (value.Contains(":") || value.Contains("\\") || value.Contains("?") || value.Contains("#") || value.Any(char.IsControl)) return false;
        var slash = containingEntry.LastIndexOf('/');
        var directory = slash < 0 ? string.Empty : containingEntry.Substring(0, slash);
        var combined = value.StartsWith("/", StringComparison.Ordinal) ? value : directory + "/" + value;
        var parts = new List<string>();
        foreach (var part in combined.Split('/'))
        {
            if (part.Length == 0 || part == ".") continue;
            if (part == "..") { if (parts.Count == 0) return false; parts.RemoveAt(parts.Count - 1); }
            else parts.Add(part);
        }
        resolved = string.Join("/", parts);
        return allowRoot || resolved.Length > 0;
    }

    private static IReadOnlyList<MedicalZipEntry> ReadCentralDirectory(byte[] data, int maxEntries)
    {
        var end = -1;
        for (var i = data.Length - 22; i >= Math.Max(0, data.Length - 22 - 65535); i--)
            if (U32(data, i) == 0x06054b50) { end = i; break; }
        if (end < 0) throw new MedicalInputRejectedException("MalformedZip");
        int count = U16(data, end + 10);
        long size = U32(data, end + 12), offset = U32(data, end + 16);
        if (count == 0xFFFF || size == 0xFFFFFFFF || offset == 0xFFFFFFFF) throw new MedicalInputRejectedException("Zip64Unsupported");
        if (count > maxEntries) throw new MedicalBudgetExceededException("EntryCount");
        if (offset + size > end) throw new MedicalInputRejectedException("MalformedZip");
        var result = new List<MedicalZipEntry>();
        var position = (int)offset;
        for (var i = 0; i < count; i++)
        {
            if (position + 46 > end || U32(data, position) != 0x02014b50) throw new MedicalInputRejectedException("MalformedZip");
            int nameLength = U16(data, position + 28), extra = U16(data, position + 30), comment = U16(data, position + 32);
            if (position + 46 + nameLength > end) throw new MedicalInputRejectedException("MalformedZip");
            var entry = new MedicalZipEntry
            {
                VersionNeeded = U16(data, position + 6), Flags = U16(data, position + 8), Method = U16(data, position + 10),
                CompressedSize = U32(data, position + 20), UncompressedSize = U32(data, position + 24),
                LocalOffset = U32(data, position + 42),
                Name = Encoding.UTF8.GetString(data, position + 46, nameLength).TrimStart('/')
            };
            if (!entry.Name.EndsWith("/", StringComparison.Ordinal)) result.Add(entry);
            position += 46 + nameLength + extra + comment;
        }
        return result;
    }

    private static int U16(byte[] data, int at) => data[at] | (data[at + 1] << 8);
    private static long U32(byte[] data, int at) => (uint)(data[at] | (data[at + 1] << 8) | (data[at + 2] << 16) | (data[at + 3] << 24));
}
