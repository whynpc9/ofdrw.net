using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Core.IO;

internal static class BoundedStreamCopy
{
    internal static long Copy(Stream input, Stream output, long maximumBytes, string description, CancellationToken cancellationToken)
    {
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        cancellationToken.ThrowIfCancellationRequested();
        if (input.CanSeek && input.Length - input.Position > maximumBytes)
            throw new InvalidDataException($"{description} exceeds the configured {maximumBytes} byte limit.");
        var buffer = new byte[81920];
        long total = 0;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var remaining = maximumBytes - total;
            var count = remaining >= buffer.Length ? buffer.Length : (int)remaining + 1;
            var read = input.Read(buffer, 0, count);
            if (read == 0) return total;
            if (read > remaining) throw new InvalidDataException($"{description} exceeds the configured {maximumBytes} byte limit.");
            output.Write(buffer, 0, read);
            total += read;
        }
    }

    internal static async Task<long> CopyAsync(
        Stream input,
        Stream output,
        long maximumBytes,
        string description,
        CancellationToken cancellationToken)
    {
        if (maximumBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (input.CanSeek && input.Length - input.Position > maximumBytes)
        {
            throw new InvalidDataException($"{description} exceeds the configured {maximumBytes} byte limit.");
        }

        var buffer = new byte[81920];
        long total = 0;
        while (true)
        {
            // Read at most one byte beyond the limit to distinguish EOF from an
            // oversized non-seekable stream, without staging the excess input.
            var remaining = maximumBytes - total;
            var count = remaining >= buffer.Length ? buffer.Length : (int)remaining + 1;
            var read = await input.ReadAsync(buffer, 0, count, cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                return total;
            }

            if (read > remaining)
            {
                throw new InvalidDataException($"{description} exceeds the configured {maximumBytes} byte limit.");
            }

            await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
            total += read;
        }
    }
}
