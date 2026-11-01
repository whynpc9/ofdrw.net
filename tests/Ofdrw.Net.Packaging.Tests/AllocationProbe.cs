using System.Diagnostics;
using System.Reflection;
using System.Runtime;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging.Tests;

// The test SDK normally supplies an empty executable entry point. VSTest still
// loads this assembly normally; only an explicit diagnostic invocation runs here.
internal static class AllocationProbeProgram
{
    private static int Main(string[] args)
    {
        if (args.Length == 0) return 0;
        if (args.Length != 3 || args[0] is not ("--audit025-budget" or "--audit025-budget-regression-control" or "--audit025-pressure"))
        {
            Console.Error.WriteLine("Expected --audit025-budget|--audit025-budget-regression-control|--audit025-pressure FONT RECEIPT.");
            return 2;
        }
        return AllocationProbe.Write(args[0], args[1], args[2]);
    }
}

internal static class AllocationProbe
{
    private const string FontSha = "b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5";
    private const long RegionBytes = 8L * 1024 * 1024;
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    internal static async Task<AllocationReceipt> RunBudgetChild(string font, Xunit.Abstractions.ITestOutputHelper output)
        => await RunChild("--audit025-budget", 0, font, output);

    internal static async Task<AllocationReceipt> RunRegressionChild(string font, Xunit.Abstractions.ITestOutputHelper output)
        => await RunChild("--audit025-budget-regression-control", 1, font, output);

    private static async Task<AllocationReceipt> RunChild(string mode, int expectedExit, string font,
        Xunit.Abstractions.ITestOutputHelper output)
    {
        var folder = Environment.GetEnvironmentVariable("OFDRW_AUDIT025_DIAGNOSTICS_DIRECTORY")
            ?? Path.Combine(Path.GetTempPath(), "ofdrw-audit025");
        Directory.CreateDirectory(folder);
        var path = Path.GetFullPath(Path.Combine(folder, $"budget-{Environment.ProcessId}-{Guid.NewGuid():N}.json"));
        var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet")
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (var argument in new[] { typeof(AllocationProbe).Assembly.Location, mode, font, path })
            start.ArgumentList.Add(argument);
        using var process = Process.Start(start) ?? throw new InvalidOperationException("Allocation probe did not start.");
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        try
        {
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(20));
        }
        finally
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            }
        }
        output.WriteLine($"Allocation receipt: {path}");
        output.WriteLine(await stdout.WaitAsync(TimeSpan.FromSeconds(5)));
        var errors = await stderr.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(process.ExitCode == expectedExit, $"Allocation probe exited {process.ExitCode}, expected {expectedExit}: {errors}; receipt: {path}");
        var receipt = JsonSerializer.Deserialize<AllocationReceipt>(File.ReadAllText(path), Json)
            ?? throw new InvalidDataException("Missing allocation receipt.");
        Assert.Equal(mode, receipt.Mode);
        Assert.Equal(FontSha, receipt.FixtureSha256);
        Assert.Equal(Hash(File.ReadAllBytes(typeof(OfdTextElement).Assembly.Location)), receipt.Core.Sha256);
        Assert.Equal(typeof(OfdTextElement).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
            receipt.Core.InformationalVersion);
        Assert.All(receipt.Cases, item => Assert.All(item.Samples, sample =>
            Assert.True(sample.Before >= 0 && sample.After >= sample.Before,
                $"Invalid raw allocation counters: before={sample.Before}, after={sample.After}.")));
        return receipt;
    }

    internal static int Write(string mode, string font, string path)
    {
        // Reserve the destination before any probe work; never overwrite an
        // earlier receipt, including a failed run's first observations.
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            using var destination = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
            var result = new AllocationReceipt { Mode = mode, Core = Identity(typeof(OfdTextElement).Assembly) };
            try
            {
                var bytes = File.ReadAllBytes(font);
                result.FixtureSha256 = Hash(bytes);
                if (result.FixtureSha256 != FontSha) throw new InvalidDataException("Unexpected allocation fixture.");
                var cmap = new OpenTypeCmap(new OpenTypeFace(bytes));
                if (mode.StartsWith("--audit025-budget", StringComparison.Ordinal)) Budget(result, cmap);
                else Pressure(result, cmap);
            }
            catch (Exception error) { result.Error = error.ToString(); }
            result.Passed = result.Error is null && (!mode.StartsWith("--audit025-budget", StringComparison.Ordinal) || result.Cases.All(Accept));
            using var writer = new StreamWriter(destination);
            writer.Write(JsonSerializer.Serialize(result, Json));
            writer.WriteLine();
            return result.Passed ? 0 : 1;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 2;
        }
    }

    private static void Budget(AllocationReceipt result, OpenTypeCmap cmap)
    {
        foreach (var (role, delta) in new (string, string?)[] { ("null", null), ("empty", ""), ("whitespace", " \t\r\n") })
            foreach (var length in new[] { 0, 32, 4096, 8192 })
                result.Cases.Add(Interval("target", role, length, delta, cmap,
                    regression: result.Mode == "--audit025-budget-regression-control"));
        result.Cases.Add(Interval("empty", "none", 0, null, cmap));
        result.Cases.Add(Interval("byte8192", "none", 0, null, cmap, 8192));
        result.Cases.Add(Interval("old-per-grapheme", "null", 4096, null, cmap, 4096));
        result.Cases.Add(Interval("old-per-grapheme", "null", 8192, null, cmap, 8192));
    }

    private static AllocationCase Interval(string operation, string role, int length, string? delta,
        OpenTypeCmap cmap, long minimum = 0, bool regression = false)
    {
        var text = Element(length, delta);
        var result = new AllocationCase { Operation = operation, DeltaRole = role, Length = length,
            MinimumExpectedAllocation = minimum, Samples = new AllocationSample[9],
            MeasuredOperation = regression ? "old-per-grapheme" : operation };
        var recorded = 0;
        var started = false;
        try
        {
            // Inputs, result storage and font parsing precede this interval.
            started = GC.TryStartNoGCRegion(RegionBytes, disallowFullBlockingGC: true);
            if (!started) throw new InvalidOperationException("No-GC interval could not start; no fallback.");
            result.CollectionsBefore = Counts();
            for (var index = 0; index < result.Samples.Length; index++)
            {
                var thread = Environment.CurrentManagedThreadId;
                var before = GC.GetAllocatedBytesForCurrentThread();
                Execute(result.MeasuredOperation, text, cmap, index == 0 ? "warmup" : "measured");
                var after = GC.GetAllocatedBytesForCurrentThread();
                result.Samples[index] = new(index + 1, index == 0, before, after, thread);
                recorded++;
            }
            result.CollectionsAfter = Counts();
            var sameThread = true;
            foreach (var sample in result.Samples) sameThread &= sample.Thread == Environment.CurrentManagedThreadId;
            result.ValidNoGcInterval = GCSettings.LatencyMode == GCLatencyMode.NoGCRegion &&
                result.CollectionsBefore == result.CollectionsAfter && sameThread;
            if (!result.ValidNoGcInterval) throw new InvalidOperationException("No-GC interval was invalidated; no fallback.");
        }
        catch (Exception error) { result.Error = error.ToString(); result.ValidNoGcInterval = false; }
        finally
        {
            if (started)
            {
                try { GC.EndNoGCRegion(); }
                catch (Exception error) { result.Error = error.ToString(); result.ValidNoGcInterval = false; }
            }
        }
        result.SamplesRecorded = recorded;
        return result;
    }

    internal static bool Accept(AllocationCase item) => item.ValidNoGcInterval && item.Error is null &&
        item.SamplesRecorded == 9 && item.Samples.All(sample => sample.Before >= 0 && sample.After >= sample.Before) &&
        item.Samples.Skip(1).All(sample => item.Operation switch
        {
            "target" => sample.Allocated >= 0 && sample.Allocated < 4096,
            "empty" => sample.Allocated == 0,
            _ => sample.Allocated >= item.MinimumExpectedAllocation
        });

    // This explicitly requested diagnostic keeps every raw interval. It does
    // not retry/relax the budget test or attribute an earlier CI failure.
    private static void Pressure(AllocationReceipt result, OpenTypeCmap cmap)
    {
        var operations = new[] { "empty-spin", "target", "byte8192", "old-per-grapheme" };
        var text = Element(4096, " \t\r\n");
        foreach (var operation in operations) Execute(operation, text, cmap, "warmup");
        var samples = new PressureSample[128];
        using var stop = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        using var ready = new ManualResetEventSlim();
        Exception? workerError = null;
        var requests = 0;
        var watch = Stopwatch.StartNew();
        result.PressureCollectionsBefore = Counts();
        var worker = new Thread(() =>
        {
            try
            {
                var ring = new byte[16][];
                for (var index = 0; index < ring.Length; index++) ring[index] = Allocate(1024 * 1024);
                var roots = new object[262144];
                for (var index = 0; index < roots.Length; index++) roots[index] = new object();
                GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: false);
                ready.Set();
                for (var index = 0; index < 256 && !stop.IsCancellationRequested; index++)
                {
                    ring[index % ring.Length] = Allocate(1024 * 1024);
                    Interlocked.Increment(ref requests);
                    GC.Collect(2, GCCollectionMode.Forced, blocking: false, compacting: false);
                    Thread.Sleep(4);
                }
                GC.KeepAlive(roots); GC.KeepAlive(ring);
            }
            catch (Exception error) { workerError = error; ready.Set(); }
        }) { IsBackground = true, Name = "Bounded allocation counter diagnostic" };
        worker.Start();
        try
        {
            if (!ready.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException("GC worker startup exceeded five seconds.");
            if (workerError is not null) throw new InvalidOperationException("GC worker failed.", workerError);
            for (var index = 0; index < samples.Length; index++)
            {
                if (watch.Elapsed > TimeSpan.FromSeconds(20)) throw new TimeoutException("GC diagnostic exceeded twenty seconds.");
                var anchor = Allocate(128);
                var operation = operations[index % operations.Length];
                var counts = Counts();
                var before = GC.GetAllocatedBytesForCurrentThread();
                Execute(operation, text, cmap, "measured");
                var after = GC.GetAllocatedBytesForCurrentThread();
                samples[index] = new(operation, index / operations.Length + 1, before, after, counts, Counts(), Environment.CurrentManagedThreadId);
                GC.KeepAlive(anchor);
                result.PressureSamplesRecorded++;
            }
        }
        finally
        {
            stop.Cancel();
            if (!worker.Join(TimeSpan.FromSeconds(5))) throw new TimeoutException("GC worker shutdown exceeded five seconds.");
            result.Pressure = samples;
            result.BackgroundCollectionRequests = requests;
            result.PressureCollectionsAfter = Counts();
        }
        if (workerError is not null) throw new InvalidOperationException("GC worker failed.", workerError);
        if (requests == 0) throw new InvalidOperationException("No background collection request ran; diagnostic inconclusive.");
        if (samples.Any(sample => sample.Operation == "byte8192" && sample.Allocated < 8192 ||
            sample.Operation == "old-per-grapheme" && sample.Allocated < 4096))
            throw new InvalidOperationException("Allocation diagnostic missed a positive/regression control.");
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void Execute(string operation, OfdTextElement text, OpenTypeCmap cmap, string name)
    {
        switch (operation)
        {
            case "target": EmbeddedFontCoverage.Validate(text, cmap, name); break;
            case "empty": break;
            case "empty-spin": Thread.SpinWait(100000); break;
            case "byte8192": GC.KeepAlive(Allocate(8192)); break;
            case "old-per-grapheme":
                foreach (var glyph in OfdTextGeometry.EnumerateGlyphs(text.Text, default))
                    EmbeddedFontCoverage.Validate(glyph, cmap, name);
                break;
            default: throw new InvalidOperationException(operation);
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)] private static byte[] Allocate(int size) => new byte[size];
    private static OfdTextElement Element(int length, string? delta)
    { var text = new string('A', length); return new() { Text = text, Runs = { new OfdTextRun { Text = text, DeltaX = delta } } }; }
    private static CollectionCounts Counts() => new(GC.CollectionCount(0), GC.CollectionCount(1), GC.CollectionCount(2));
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static AssemblyIdentity Identity(Assembly assembly) => new(assembly.Location,
        Hash(File.ReadAllBytes(assembly.Location)), assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
        assembly.GetCustomAttribute<TargetFrameworkAttribute>()?.FrameworkName);
}

internal sealed class AllocationReceipt
{
    public string Mode { get; set; } = "";
    public bool Passed { get; set; }
    public string? Error { get; set; }
    public string FixtureSha256 { get; set; } = "";
    public AssemblyIdentity Core { get; set; } = new("", "", null, null);
    public string Runtime { get; set; } = RuntimeInformation.FrameworkDescription;
    public string OperatingSystem { get; set; } = RuntimeInformation.OSDescription;
    public string Rid { get; set; } = RuntimeInformation.RuntimeIdentifier;
    public string Architecture { get; set; } = RuntimeInformation.ProcessArchitecture.ToString();
    public bool ServerGc { get; set; } = GCSettings.IsServerGC;
    public string? TieredCompilation { get; set; } = Environment.GetEnvironmentVariable("DOTNET_TieredCompilation");
    public string? TieredPgo { get; set; } = Environment.GetEnvironmentVariable("DOTNET_TieredPGO");
    public string? ReadyToRun { get; set; } = Environment.GetEnvironmentVariable("DOTNET_ReadyToRun");
    public List<AllocationCase> Cases { get; set; } = new();
    public PressureSample[] Pressure { get; set; } = Array.Empty<PressureSample>();
    public int PressureSamplesRecorded { get; set; }
    public int BackgroundCollectionRequests { get; set; }
    public CollectionCounts PressureCollectionsBefore { get; set; }
    public CollectionCounts PressureCollectionsAfter { get; set; }
}
internal sealed class AllocationCase
{
    public string Operation { get; set; } = "";
    public string MeasuredOperation { get; set; } = "";
    public string DeltaRole { get; set; } = "";
    public int Length { get; set; }
    public long MinimumExpectedAllocation { get; set; }
    public bool ValidNoGcInterval { get; set; }
    public string? Error { get; set; }
    public int SamplesRecorded { get; set; }
    public AllocationSample[] Samples { get; set; } = Array.Empty<AllocationSample>();
    public CollectionCounts CollectionsBefore { get; set; }
    public CollectionCounts CollectionsAfter { get; set; }
}
internal sealed record AssemblyIdentity(string Path, string Sha256, string? InformationalVersion, string? TargetFramework);
internal readonly record struct CollectionCounts(int Gen0, int Gen1, int Gen2);
internal readonly record struct AllocationSample(int Ordinal, bool Warmup, long Before, long After, int Thread)
{ public long Allocated => After - Before; }
internal readonly record struct PressureSample(string Operation, int Ordinal, long Before, long After,
    CollectionCounts CollectionsBefore, CollectionCounts CollectionsAfter, int Thread)
{ public long Allocated => After - Before; }
