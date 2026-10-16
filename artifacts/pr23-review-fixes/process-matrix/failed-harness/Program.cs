using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using Ofdrw.Net.Core.Models;

if (args.FirstOrDefault() == "fixture")
{
    Console.Write(JsonSerializer.Serialize(new { arguments = args.Skip(1).ToArray(), cwd = Environment.CurrentDirectory, marker = Environment.GetEnvironmentVariable("OFDRW_PROCESS_MARKER") }));
    return 0;
}
var mode = args.FirstOrDefault() ?? "default";
var output = Path.GetFullPath(args.ElementAtOrDefault(1) ?? "process-consumer-output");
Directory.CreateDirectory(output);
var cases = new List<object>();
var assembly = args.Length > 2 ? Assembly.LoadFile(Path.GetFullPath(args[2])) : typeof(OfdAttachment).Assembly;
var method = assembly.GetType("Ofdrw.Net.Core.Processes.ExternalProcessRunner", true)!.GetMethod("RunAsync", BindingFlags.Static | BindingFlags.NonPublic)!;
var trueExecutable = File.Exists("/usr/bin/true") ? "/usr/bin/true" : "/bin/true";
using (var initialized = Process.Start(new ProcessStartInfo(trueExecutable) { UseShellExecute = false }))
    initialized!.WaitForExit();
if (mode == "ignore" && signal(17, (IntPtr)1) == (IntPtr)(-1)) throw new Win32Exception(Marshal.GetLastPInvokeError());
if (mode == "pid1" && Environment.ProcessId != 1) throw new InvalidOperationException("PID1 proof requires actual pid1");
async Task<object> Run(ProcessStartInfo info, TimeSpan timeout, CancellationToken cancellation = default)
{
    Task task;
    try { task = (Task)method.Invoke(null, new object?[] { info, timeout, cancellation, null })!; }
    catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException!).Throw(); throw; }
    await task;
    return task.GetType().GetProperty("Result")!.GetValue(task)!;
}
int Code(object result) => (int)result.GetType().GetProperty("ExitCode", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(result)!;
string Text(object result) => (string)result.GetType().GetProperty("Output", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(result)!;
async Task Check(string name, Func<Task> test)
{
    var watch = Stopwatch.StartNew();
    try { await test(); cases.Add(new { name, outcome = "Passed", milliseconds = watch.ElapsedMilliseconds }); }
    catch (Exception error) { cases.Add(new { name, outcome = "Failed", milliseconds = watch.ElapsedMilliseconds, error = error.ToString() }); throw; }
    finally { File.WriteAllText(Path.Combine(output, "results.json"), JsonSerializer.Serialize(new { mode, pid = Environment.ProcessId, runtime = RuntimeInformation.FrameworkDescription, os = RuntimeInformation.OSDescription, architecture = RuntimeInformation.ProcessArchitecture.ToString(), corePath = assembly.Location, coreSha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(assembly.Location))), cases }, new JsonSerializerOptions { WriteIndented = true })); }
}
void Require(bool condition, string why) { if (!condition) throw new InvalidOperationException(why); }
ProcessStartInfo Shell(string command) { var info = new ProcessStartInfo("/bin/sh"); info.ArgumentList.Add("-c"); info.ArgumentList.Add(command); return info; }
var limit = TimeSpan.FromSeconds(10);
await Check("all-four-embedded-assets-and-private-extraction", () =>
{
    var expectedRids = new[] { "linux-x64", "linux-arm64", "osx-x64", "osx-arm64" };
    using var manifestStream = assembly.GetManifestResourceStream("Ofdrw.ProcessSupervisor.manifest.json") ?? throw new Exception("missing embedded manifest");
    using var manifest = JsonDocument.Parse(manifestStream);
    var files = manifest.RootElement.GetProperty("files").EnumerateArray().ToArray();
    Require(files.Select(x => x.GetProperty("rid").GetString()).Order().SequenceEqual(expectedRids.Order()), "incomplete RID inventory");
    foreach (var entry in files)
    {
        using var stream = assembly.GetManifestResourceStream("Ofdrw.ProcessSupervisor." + entry.GetProperty("rid").GetString()) ?? throw new Exception("missing RID resource");
        Require(stream.Length == entry.GetProperty("size").GetInt64(), "resource size mismatch");
        Require(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream)).Equals(entry.GetProperty("sha256").GetString(), StringComparison.OrdinalIgnoreCase), "resource hash mismatch");
    }
    var type = assembly.GetType("Ofdrw.Net.Core.Processes.NativeSupervisorAssets", true)!;
    var extraction = (IDisposable)type.GetMethod("Extract", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, null)!;
    var path = (string)extraction.GetType().GetProperty("ExecutablePath", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(extraction)!;
    var directory = Path.GetDirectoryName(path)!;
    try
    {
        var required = UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute;
        Require(File.GetUnixFileMode(path) == required && File.GetUnixFileMode(directory) == required, "extraction must be private0700");
        var selected = (OperatingSystem.IsMacOS() ? "osx-" : "linux-") + (RuntimeInformation.ProcessArchitecture == Architecture.X64 ? "x64" : "arm64");
        var entry = files.Single(x => x.GetProperty("rid").GetString() == selected);
        Require(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path))).Equals(entry.GetProperty("sha256").GetString(), StringComparison.OrdinalIgnoreCase), "wrong extracted native bytes");
    }
    finally { extraction.Dispose(); }
    Require(!Directory.Exists(directory), "private extraction not cleaned");
    return Task.CompletedTask;
});
await Check("real-zero-no-shell", async () => Require(Code(await Run(new ProcessStartInfo(trueExecutable), limit)) == 0, "real0"));
await Check("real-seven", async () => Require(Code(await Run(Shell("exit 7"), limit)) == 7, "real7"));
await Check("argv-cwd-environment", async () =>
{
    var cwd = Path.Combine(output, "中文 directory"); Directory.CreateDirectory(cwd);
    var info = new ProcessStartInfo(Environment.ProcessPath!) { WorkingDirectory = cwd };
    info.ArgumentList.Add(Assembly.GetExecutingAssembly().Location); info.ArgumentList.Add("fixture");
    var expected = new[] { "", "two words", "中文-é", "single'quote", "double\"quote", "back\\slash", "$literal;`notexec`" };
    foreach (var value in expected) info.ArgumentList.Add(value);
    info.Environment["OFDRW_PROCESS_MARKER"] = "中-env";
    using var json = JsonDocument.Parse(Text(await Run(info, limit)));
    Require(json.RootElement.GetProperty("arguments").EnumerateArray().Select(x => x.GetString()).SequenceEqual(expected), "argv changed");
    Require(json.RootElement.GetProperty("cwd").GetString() == cwd, "cwd changed");
    Require(json.RootElement.GetProperty("marker").GetString() == "中-env", "env changed");
});
await Check("missing-executable-native-error", async () =>
{
    try { await Run(new ProcessStartInfo(Path.Combine(output, "missing-" + Guid.NewGuid())), limit); throw new Exception("missing unexpectedly succeeded"); }
    catch (Win32Exception error) { Require(error.NativeErrorCode == 2, "not ENOENT"); }
});
await Check("non-executable-native-error", async () =>
{
    var path = Path.Combine(output, "not-executable"); File.WriteAllText(path, "not executable"); File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite);
    try { await Run(new ProcessStartInfo(path), limit); throw new Exception("permission unexpectedly succeeded"); }
    catch (Win32Exception error) { Require(error.NativeErrorCode == 13, "not EACCES"); }
});
await Check("bounded-timeout", async () =>
{
    var info = new ProcessStartInfo("/bin/sleep"); info.ArgumentList.Add("30");
    try { await Run(info, TimeSpan.FromMilliseconds(350)); throw new Exception("timeout unexpectedly succeeded"); }
    catch (TimeoutException) { }
});
await Check("bounded-cancellation", async () =>
{
    using var token = new CancellationTokenSource(TimeSpan.FromMilliseconds(350)); var info = new ProcessStartInfo("/bin/sleep"); info.ArgumentList.Add("30");
    try { await Run(info, limit, token.Token); throw new Exception("cancel unexpectedly succeeded"); }
    catch (OperationCanceledException) { Require(token.IsCancellationRequested, "foreign cancellation"); }
});
await Check("parent-first-descendant-cleanup", async () =>
{
    var marker = Path.Combine(output, "descendant.pid"); var escaped = "'" + marker.Replace("'", "'\\''") + "'";
    try { await Run(Shell("sleep 30 & echo $! > " + escaped + "; exit 0"), TimeSpan.FromSeconds(1)); throw new Exception("inherited pipes unexpectedly completed"); }
    catch (TimeoutException) { }
    Require(File.Exists(marker), "descendant never published"); var pid = int.Parse(File.ReadAllText(marker).Trim());
    // kill(pid,0) observes zombies as present; Linux state is separately recorded.
    var proc = "/proc/" + pid + "/stat"; var zombie = File.Exists(proc) && File.ReadAllText(proc).Split(')')[1].TrimStart().StartsWith("Z ");
    Require(kill(pid, 0) < 0 || zombie, "descendant still executing after cleanup");
});
Console.WriteLine("PACKAGE_PROCESS_CONSUMER_PASS " + cases.Count);
return 0;

partial class Program
{
    [DllImport("libc", SetLastError = true)] static extern IntPtr signal(int number, IntPtr handler);
    [DllImport("libc", SetLastError = true)] static extern int kill(int pid, int signal);
}
