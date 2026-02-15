using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ofdrw.Net.Packaging.Archive;

public sealed class OfdPackageArchive
{
    private readonly Dictionary<string, byte[]> _entries;

    public OfdPackageArchive(Dictionary<string, byte[]> entries)
    {
        _entries = entries ?? throw new ArgumentNullException(nameof(entries));
    }

    public IReadOnlyCollection<string> EntryNames => _entries.Keys;

    public bool Contains(string entryName)
    {
        return _entries.ContainsKey(Normalize(entryName));
    }

    public bool TryGetBytes(string entryName, out byte[] bytes)
    {
        return _entries.TryGetValue(Normalize(entryName), out bytes!);
    }

    public byte[] GetBytes(string entryName)
    {
        if (!TryGetBytes(entryName, out var bytes))
        {
            throw new KeyNotFoundException($"Entry not found: {entryName}");
        }

        return bytes;
    }

    public string ReadUtf8Text(string entryName)
    {
        return Encoding.UTF8.GetString(GetBytes(entryName));
    }

    public IEnumerable<string> FindByPrefix(string prefix)
    {
        var normalized = Normalize(prefix).TrimEnd('/') + "/";
        return _entries.Keys.Where(x => x.StartsWith(normalized, StringComparison.OrdinalIgnoreCase));
    }

    private static string Normalize(string path)
    {
        return path.Replace('\\', '/').TrimStart('/');
    }
}
