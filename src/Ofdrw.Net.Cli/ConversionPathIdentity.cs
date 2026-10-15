internal static class ConversionPathIdentity
{
    private static readonly StringComparer PathComparer = OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
        ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;

    internal static void RequireDifferentPaths(string inputPath, string outputPath)
    {
        if (PathComparer.Equals(Resolve(inputPath), Resolve(outputPath)))
            throw new ArgumentException("Conversion input and output must use different paths.");
    }

    private static string Resolve(string path)
    {
        var absolute = Path.GetFullPath(path);
        var root = Path.GetPathRoot(absolute)!;
        return ResolveComponents(root, absolute[root.Length..], new HashSet<string>(PathComparer));
    }

    private static string ResolveComponents(string current, string relative, HashSet<string> activeLinks)
    {
        var components = relative.Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries);
        for (var index = 0; index < components.Length; index++)
        {
            if (components[index] == ".") continue;
            if (components[index] == "..")
            {
                current = Path.GetDirectoryName(current) ?? current;
                continue;
            }

            var parent = current;
            current = Path.Combine(parent, components[index]);
            FileSystemInfo entry = index == components.Length - 1 ? new FileInfo(current) : new DirectoryInfo(current);
            var target = entry.LinkTarget;
            if (target is null) continue; // Keep missing output leaves, resolving their existing ancestors.
            if (activeLinks.Count >= 64 || !activeLinks.Add(current))
                throw new IOException("Conversion path contains a link cycle or too many nested links.");
            var link = current;
            try
            {
                if (Path.IsPathRooted(target))
                {
                    var root = Path.GetPathRoot(target)!;
                    current = ResolveComponents(root, target[root.Length..], activeLinks);
                }
                else
                {
                    // Resolve each target component before '..'; lexical normalization would
                    // lose the identity of a preceding directory symlink.
                    current = ResolveComponents(parent, target, activeLinks);
                }
            }
            finally { activeLinks.Remove(link); }
        }
        return current;
    }
}
