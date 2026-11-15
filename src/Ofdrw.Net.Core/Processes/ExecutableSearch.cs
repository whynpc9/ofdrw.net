using System;
using System.IO;

namespace Ofdrw.Net.Core.Processes;

/// <summary>Resolves a bare tool name without consulting the working directory.</summary>
internal static class ExecutableSearch
{
    internal static string? Find(string name, string baseDirectory, string? pathVariable,
        bool windows, Func<string, bool> isExecutableFile)
    {
        if (Path.IsPathRooted(name)) return name;
        if (name.IndexOf(Path.DirectorySeparatorChar) >= 0 || name.IndexOf(Path.AltDirectorySeparatorChar) >= 0)
            return windows ? name : Path.GetFullPath(name);

        // The Windows launcher quotes the resolved path as the first command-line token.
        if (windows && name.IndexOf('"') >= 0) return null;
        var names = windows && !Path.HasExtension(name) ? new[] { name + ".exe", name + ".com" } : new[] { name };
        return FindInDirectory(baseDirectory) ?? FindInPath();

        string? FindInPath()
        {
            foreach (var entry in (pathVariable ?? string.Empty).Split(Path.PathSeparator))
            {
                var directory = entry;
                if (windows && directory.Length >= 2 && directory[0] == '"' && directory[directory.Length - 1] == '"')
                    directory = directory.Substring(1, directory.Length - 2);
                var executable = FindInDirectory(directory);
                if (executable != null) return executable;
            }
            return null;
        }

        string? FindInDirectory(string directory)
        {
            if (windows && directory.IndexOf('"') >= 0) return null;
            foreach (var fileName in names)
            {
                string candidate;
                try
                {
                    if (!IsAbsolute(directory)) return null;
                    candidate = Path.GetFullPath(Path.Combine(directory, fileName));
                }
                catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
                {
                    return null;
                }
                if (isExecutableFile(candidate)) return candidate;
            }
            return null;
        }
    }

    private static bool IsAbsolute(string path)
    {
        if (!Path.IsPathRooted(path)) return false;
        if (Path.DirectorySeparatorChar == '/') return true;

        // netstandard2.0 lacks IsPathFullyQualified. An explicit drive root or UNC prefix
        // is required on Windows because drive-relative and root-relative paths are rooted.
        return (path.Length >= 3 && path[1] == ':' && IsSeparator(path[2]))
            || (path.Length >= 2 && IsSeparator(path[0]) && IsSeparator(path[1]));
    }

    private static bool IsSeparator(char value) => value == Path.DirectorySeparatorChar || value == Path.AltDirectorySeparatorChar;
}
