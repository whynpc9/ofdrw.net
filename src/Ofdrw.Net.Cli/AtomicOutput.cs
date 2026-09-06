/// <summary>Stages output beside its destination and publishes only a completed file.</summary>
internal sealed class AtomicOutput : IAsyncDisposable
{
    private readonly string _destination;
    private readonly string _temporaryPath;
    private bool _committed;

    internal AtomicOutput(string destination)
    {
        _destination = Path.GetFullPath(destination);
        var directory = Path.GetDirectoryName(_destination)!;
        Directory.CreateDirectory(directory);
        _temporaryPath = Path.Combine(directory, $".ofdrw-{Guid.NewGuid():N}.tmp");
        Stream = new FileStream(_temporaryPath, FileMode.CreateNew, FileAccess.Write,
            FileShare.None, 81920, FileOptions.Asynchronous);
    }

    internal FileStream Stream { get; }

    internal async Task CommitAsync(CancellationToken cancellationToken)
    {
        await Stream.FlushAsync(cancellationToken).ConfigureAwait(false);
        await Stream.DisposeAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        // Both paths are on the same filesystem. A failed conversion never opens
        // the destination for writing, and replacement exposes a complete file.
        File.Move(_temporaryPath, _destination, overwrite: true);
        _committed = true;
    }

    public async ValueTask DisposeAsync()
    {
        await Stream.DisposeAsync().ConfigureAwait(false);
        if (!_committed)
        {
            File.Delete(_temporaryPath);
        }
    }
}
