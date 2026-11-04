using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Ofdrw.Net.Core.IO;

// Staging area for document copies: mkdtemp (mode 0700, atomic, never reuses a name) on Unix,
// a GUID directory under the per-user %TEMP% on Windows.
internal static class PrivateTemporaryDirectory
{
    /// Creates a new directory under the system temp path that only the current user can access.
    internal static string Create(string prefix)
    {
        if (prefix is null) throw new ArgumentNullException(nameof(prefix));
        if (prefix.Length == 0 || prefix.IndexOfAny(new[] { '/', '\\', '\0' }) >= 0)
            throw new ArgumentException("A temporary directory prefix must be a nonempty name without path separators.", nameof(prefix));
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var directory = Path.Combine(Path.GetTempPath(), prefix + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            return directory;
        }

        var template = Encoding.UTF8.GetBytes(Path.Combine(Path.GetTempPath(), prefix + "XXXXXX") + "\0");
        if (mkdtemp(template) == IntPtr.Zero)
            throw new IOException("Cannot create a private temporary directory.", new Win32Exception(Marshal.GetLastWin32Error()));
        return Encoding.UTF8.GetString(template, 0, Array.IndexOf(template, (byte)0));
    }

    /// Recursively deletes a directory from <see cref="Create"/>; cleanup failures never mask the caller's outcome.
    internal static void TryDelete(string? directory)
    {
        if (string.IsNullOrEmpty(directory)) return;
        try { Directory.Delete(directory, recursive: true); }
        catch (DirectoryNotFoundException) { }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    [DllImport("libc", SetLastError = true)]
    private static extern IntPtr mkdtemp([In, Out] byte[] template);
}
