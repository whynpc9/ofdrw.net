/// <summary>
/// Stages output beside its destination and publishes only a completed file,
/// preserving Unix permission bits when replacing an existing file.
/// </summary>
internal sealed class AtomicOutput : IAsyncDisposable
{
    private readonly string _destination;
    private readonly string _temporaryPath;
    private readonly UnixFileMode? _destinationMode;
    private bool _committed;

    internal AtomicOutput(string destination)
    {
        _destination = Path.GetFullPath(destination);
        var directory = Path.GetDirectoryName(_destination)!;
        Directory.CreateDirectory(directory);
        _temporaryPath = Path.Combine(directory, $".ofdrw-{Guid.NewGuid():N}.tmp");
        if (!OperatingSystem.IsWindows() && File.Exists(_destination))
        {
            try
            {
                _destinationMode = File.GetUnixFileMode(_destination);
            }
            catch (IOException)
            {
                // The destination may have disappeared since the existence check.
            }
            catch (UnauthorizedAccessException)
            {
                // Fall back to the default mode if metadata cannot be read.
            }

            if (_destinationMode.HasValue)
            {
                Stream = new FileStream(_temporaryPath, new FileStreamOptions
                {
                    Mode = FileMode.CreateNew,
                    Access = FileAccess.Write,
                    Share = FileShare.None,
                    BufferSize = 81920,
                    Options = FileOptions.Asynchronous,
                    UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite
                });
                return;
            }
        }

        Stream = new FileStream(_temporaryPath, FileMode.CreateNew, FileAccess.Write,
            FileShare.None, 81920, FileOptions.Asynchronous);
    }

    internal FileStream Stream { get; }

    internal async Task CommitAsync(CancellationToken cancellationToken)
    {
        await Stream.FlushAsync(cancellationToken).ConfigureAwait(false);
        // Use the opened file so a replaced temporary path cannot redirect chmod.
        if (!OperatingSystem.IsWindows() && _destinationMode is { } mode)
            File.SetUnixFileMode(Stream.SafeFileHandle, mode);
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
