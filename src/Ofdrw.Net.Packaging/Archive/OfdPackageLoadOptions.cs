namespace Ofdrw.Net.Packaging.Archive;

public sealed class OfdPackageLoadOptions
{
    public int MaxEntryCount { get; set; } = 10_000;

    public long MaxEntryUncompressedBytes { get; set; } = 128L * 1024 * 1024;

    public long MaxTotalUncompressedBytes { get; set; } = 512L * 1024 * 1024;

    public double MaxCompressionRatio { get; set; } = 1_000d;
}
