using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace Ofdrw.Net.Core.Processes;

// Assets are source-bound and verified by the SDK build, then checked again
// before execution. Extraction never uses a shared writable executable cache.
internal static partial class NativeSupervisorAssets
{
    private sealed class Asset
    {
        internal readonly string Rid;
        internal readonly long Size;
        internal readonly string Hash;
        internal Asset(string rid, long size, string hash) { Rid = rid; Size = size; Hash = hash; }
    }

    internal static Extraction Extract()
    {
        var architecture = RuntimeInformation.ProcessArchitecture;
        var suffix = architecture == Architecture.X64 ? "x64" : architecture == Architecture.Arm64 ? "arm64" :
            throw new PlatformNotSupportedException("Owned Unix processes require x64 or ARM64 native supervisor assets.");
        var rid = (RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "osx-" : "linux-") + suffix;
        var asset = FindAsset(rid);
        using var resource = typeof(NativeSupervisorAssets).Assembly.GetManifestResourceStream("Ofdrw.ProcessSupervisor." + rid)
            ?? throw new IOException("The owned-process supervisor asset is missing: " + rid);
        if (resource.Length != asset.Size) throw new IOException("The owned-process supervisor asset has an invalid size: " + rid);
        using (var hash = SHA256.Create())
        {
            var actual = BitConverter.ToString(hash.ComputeHash(resource)).Replace("-", "");
            if (!actual.Equals(asset.Hash, StringComparison.OrdinalIgnoreCase))
                throw new IOException("The owned-process supervisor asset failed its SHA256 check: " + rid);
        }
        resource.Position = 0;
        var template = Encoding.UTF8.GetBytes(Path.Combine(Path.GetTempPath(), "ofdrw-supervisor-XXXXXX") + "\0");
        if (mkdtemp(template) == IntPtr.Zero) throw Failure("Cannot create a private owned-process supervisor directory.");
        var directory = Encoding.UTF8.GetString(template, 0, Array.IndexOf(template, (byte)0));
        var extraction = new Extraction(directory);
        try
        {
            using (var file = new FileStream(extraction.ExecutablePath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                resource.CopyTo(file);
                file.Flush(true);
            }
            // The writer is closed before the helper can be executed. mkdtemp's
            // private 0700 directory protects the path during this transition.
            if (chmod(extraction.ExecutablePath, 0x1c0) != 0) throw Failure("Cannot make the owned-process supervisor executable.");
            return extraction;
        }
        catch (Exception failure)
        {
            try { extraction.Dispose(); }
            catch (Exception cleanup) { failure.Data["OwnedProcessCleanup.extraction"] = cleanup.ToString(); }
            throw;
        }
    }

    internal sealed class Extraction : IDisposable
    {
        private readonly string _directory;
        private bool _disposed;
        internal string ExecutablePath { get; }
        internal Extraction(string directory) { _directory = directory; ExecutablePath = Path.Combine(directory, "ofdrw-supervisor"); }
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            File.Delete(ExecutablePath);
            Directory.Delete(_directory);
        }
    }
    private static IOException Failure(string message) => new IOException(message, new Win32Exception(Marshal.GetLastWin32Error()));
    [DllImport("libc", SetLastError = true)] private static extern IntPtr mkdtemp([In, Out] byte[] template);
    [DllImport("libc", SetLastError = true)] private static extern int chmod(string path, int mode);
}
