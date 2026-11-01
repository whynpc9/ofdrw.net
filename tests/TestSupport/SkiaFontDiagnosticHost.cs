using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using SkiaSharp;

namespace Ofdrw.Net.TestSupport;

// Test-host collection only. The product never loads a second native library
// or probes a suspect face to complete diagnostic evidence.
internal static class SkiaFontDiagnosticHost
{
    private const string Key = "Ofdrw.Skia.FontReadFailure.v1";
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
    private static readonly ReportBudget Budget = new();
    [ThreadStatic] private static bool _collecting;

    [ModuleInitializer]
    internal static void Initialize()
    {
        var directory = Environment.GetEnvironmentVariable("OFDRW_SKIA_DIAGNOSTICS_DIRECTORY");
        if (string.IsNullOrWhiteSpace(directory)) return;
        AppDomain.CurrentDomain.FirstChanceException += (_, args) => Capture(args.Exception, directory, true);
    }

    internal static string? Capture(Exception error, string directory, bool atThrow = false)
    {
        if (_collecting) return null;
        _collecting = true;
        try
        {
            if (!error.Data.Contains(Key) || error.Data[Key] is not { } evidence) return null;
            var detail = JsonSerializer.SerializeToElement(evidence, evidence.GetType(), Options);
            var classification = detail.GetProperty("classification").GetString()!;
            if (!classification.StartsWith("native-", StringComparison.Ordinal)) return null;
            var generation = detail.GetProperty("generation").GetInt64();
            var injected = detail.GetProperty("injected").GetBoolean();
            var identity = directory + ":" + generation + ":" + classification;
            if (!Budget.TryReserve(identity, injected)) return null;
            var modules = LoadedModules();
            var host = new
            {
                schemaVersion = 1, capturedAtUtc = DateTime.UtcNow, processId = Environment.ProcessId,
                scope = injected ? "controlled-injection" : detail.GetProperty("rawStylePreflightPassed").GetBoolean()
                    ? "uninjected-guard-after-raw-style-preflight" : "uninjected-guard-without-raw-style-preflight",
                observation = !atThrow ? "explicit-post-catch-sample" :
                    detail.GetProperty("decoratedAfterOriginalThrow").GetBoolean() ? "decorated-propagation-event-not-original-first-chance" : "initial-guard-first-chance",
                errorType = error.GetType().FullName,
                architecture = RuntimeInformation.ProcessArchitecture.ToString(), osArchitecture = RuntimeInformation.OSArchitecture.ToString(),
                os = RuntimeInformation.OSDescription, osBuild = Environment.OSVersion.VersionString,
                runtime = RuntimeInformation.FrameworkDescription, rid = RuntimeInformation.RuntimeIdentifier,
                reportLimitPerProcessPerScope = 16,
                sourceSha = SourceSha(), managedSkia = AssemblyIdentity(typeof(SKTypeface).Assembly),
                adapter = AssemblyIdentity(typeof(Ofdrw.Net.Graphics.SkiaSharp.SkiaDrawEvent).Assembly),
                native = modules, failure = detail,
                limitations = "Mapped path hashes are disk files at capture time, not memory hashes. Module inventory does not independently attest each native call address. No font table reread, face call or library reload is performed."
            };
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, $"font-read-{Environment.ProcessId}-{generation}-{(injected ? "controlled" : "observed")}.json");
            File.WriteAllText(path, JsonSerializer.Serialize(host, Options) + "\n");
            return path;
        }
        catch (Exception captureError)
        {
            try { error.Data["Ofdrw.Skia.FontReadHostStatus"] = "capture-failed:" + captureError.GetType().Name; } catch { }
            return null;
        }
        finally { _collecting = false; }
    }

    internal sealed class ReportBudget
    {
        private readonly object _gate = new();
        private readonly HashSet<string> _seen = new(StringComparer.Ordinal);
        private int _controlled, _observed;
        internal bool TryReserve(string identity, bool controlled)
        {
            lock (_gate)
            {
                if (_seen.Contains(identity) || (controlled ? _controlled : _observed) >= 16) return false;
                _seen.Add(identity); // At most32 entries, including failed sinks.
                if (controlled) _controlled++; else _observed++;
                return true;
            }
        }
        internal (int Controlled, int Observed, int Deduplication) Counts
        { get { lock (_gate) return (_controlled, _observed, _seen.Count); } }
    }

    private static string? SourceSha()
    {
        var value = Environment.GetEnvironmentVariable("OFDRW_SOURCE_SHA") ?? Environment.GetEnvironmentVariable("GITHUB_SHA");
        return value is { Length: 40 } && value.All(Uri.IsHexDigit) ? value : null;
    }
    private static object AssemblyIdentity(Assembly assembly) => new
    { name = assembly.FullName, mvid = assembly.ManifestModule.ModuleVersionId, file = FileIdentity(assembly.Location) };
    private static object FileIdentity(string path)
    {
        try
        {
            using var stream = File.OpenRead(path);
            return new { path, bytes = stream.Length, sha256 = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(), status = "hashed-mapped-path" };
        }
        catch (Exception error) { return new { path, status = "hash-unavailable", reason = error.GetType().Name }; }
    }
    private static bool Relevant(string path)
    {
        var name = Path.GetFileName(path);
        return name.StartsWith("libSkiaSharp", StringComparison.OrdinalIgnoreCase) ||
            name.StartsWith("libHarfBuzzSharp", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("pdfium", StringComparison.OrdinalIgnoreCase);
    }
    private static object LoadedModules()
    {
        var paths = new HashSet<string>(StringComparer.Ordinal); var methods = new List<string>(); var errors = new List<string>();
        try
        {
            using var process = Process.GetCurrentProcess(); process.Refresh();
            foreach (ProcessModule module in process.Modules) if (Relevant(module.FileName)) paths.Add(module.FileName);
            methods.Add("current-process Process.Modules");
        }
        catch (Exception error) { errors.Add("Process.Modules:" + error.GetType().Name); }
        if (OperatingSystem.IsMacOS())
        {
            // dyld's image-array traversal is not thread safe. NOLOAD obtains
            // only an existing loader reference; dladdr is thread safe. The
            // symbol is resolved but never invoked, and our reference closes.
            foreach (var name in new[] { "libSkiaSharp.dylib", "libSkiaSharp" })
            {
                var handle = IntPtr.Zero;
                try
                {
                    handle = Dlopen(name, 0x1 | 0x10); // SDK RTLD_LAZY | RTLD_NOLOAD
                    if (handle == IntPtr.Zero) continue;
                    var symbol = Dlsym(handle, "sk_typeface_get_table_size");
                    if (symbol != IntPtr.Zero && Dladdr(symbol, out var info) != 0 && Marshal.PtrToStringUTF8(info.FileName) is { } path)
                        paths.Add(path);
                    methods.Add("existing-loader RTLD_NOLOAD + symbol dladdr; no font function invoked");
                }
                catch (Exception error) { errors.Add("NOLOAD/dladdr:" + error.GetType().Name); }
                finally { if (handle != IntPtr.Zero) Dlclose(handle); }
            }
        }
        return new
        {
            status = paths.Any(path => Path.GetFileName(path).StartsWith("libSkiaSharp", StringComparison.OrdinalIgnoreCase)) ? "loaded-skia-module-captured" : "loaded-skia-module-unavailable",
            methods, errors, modules = paths.Order().Select(FileIdentity).ToArray(),
            attribution = "Actual current-process loader inventory/symbol mapping only. Additional private copies or individual call targets are not proven absent."
        };
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct DlInfo { internal IntPtr FileName, Base, SymbolName, Symbol; }
    [DllImport("/usr/lib/libSystem.B.dylib", EntryPoint = "dlopen")]
    private static extern IntPtr Dlopen(string path, int mode);
    [DllImport("/usr/lib/libSystem.B.dylib", EntryPoint = "dlsym")]
    private static extern IntPtr Dlsym(IntPtr handle, string symbol);
    [DllImport("/usr/lib/libSystem.B.dylib", EntryPoint = "dladdr")]
    private static extern int Dladdr(IntPtr symbol, out DlInfo info);
    [DllImport("/usr/lib/libSystem.B.dylib", EntryPoint = "dlclose")]
    private static extern int Dlclose(IntPtr handle);
}
