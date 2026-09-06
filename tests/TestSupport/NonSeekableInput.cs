namespace Ofdrw.Net.TestSupport;

internal sealed class NonSeekableInput(byte[] bytes) : MemoryStream(bytes)
{
    public override bool CanSeek => false;
    internal long BytesRead { get; private set; }

    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        var read = await base.ReadAsync(buffer, offset, count, cancellationToken);
        BytesRead += read;
        return read;
    }
}
