namespace Ofdrw.Net.Packaging.Archive;

public sealed class OfdPackageLoadOptions
{
    /// <summary>Maximum compressed input size, enforced before or during staging.</summary>
    public long MaxInputBytes { get; set; } = 512L * 1024 * 1024;

    public int MaxEntryCount { get; set; } = 10_000;

    /// <summary>Maximum page references materialized by OfdReader.</summary>
    public int MaxPageCount { get; set; } = 10_000;

    public long MaxEntryUncompressedBytes { get; set; } = 128L * 1024 * 1024;

    public long MaxTotalUncompressedBytes { get; set; } = 512L * 1024 * 1024;

    public double MaxCompressionRatio { get; set; } = 1_000d;
}
