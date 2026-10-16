using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Converter.Svg.Internal;

// A private segmented memory spool avoids placing document plaintext in a shared
// temp directory. Allocate one bounded block only after the complete write fits.
internal sealed class HtmlSpool : Stream
{
    private const int BlockSize = 16384;
    private readonly List<byte[]> _blocks = new();
    private readonly OfdHtmlOptions _options;
    private readonly CancellationToken _token;
    private long _length, _pageStart;
    private bool _page;
    internal HtmlSpool(OfdHtmlOptions options, CancellationToken token) { _options = options; _token = token; }
    internal void BeginPage() { _pageStart = _length; _page = true; }
    internal void EndPage() { _page = false; }
    public override void Write(byte[] buffer, int offset, int count)
    {
        _token.ThrowIfCancellationRequested();
        if (count < 0 || count > _options.MaxOutputBytes - _length ||
            (_page && count > _options.MaxPageBytes - (_length - _pageStart)))
            throw new InvalidDataException("HTML preview capacity exceeded: output/page bytes.");
        while (count > 0)
        {
            _token.ThrowIfCancellationRequested();
            var block = checked((int)(_length / BlockSize)); var start = (int)(_length % BlockSize);
            if (block == _blocks.Count) _blocks.Add(new byte[(int)Math.Min(BlockSize, _options.MaxOutputBytes - _length)]);
            var take = Math.Min(count, _blocks[block].Length - start);
            Buffer.BlockCopy(buffer, offset, _blocks[block], start, take);
            offset += take; count -= take; _length += take;
        }
    }
    internal async Task CopyResultAsync(Stream output, CancellationToken token)
    {
        var remaining = _length;
        foreach (var block in _blocks)
        {
            token.ThrowIfCancellationRequested();
            var count = (int)Math.Min(block.Length, remaining);
            await output.WriteAsync(block, 0, count, token).ConfigureAwait(false); remaining -= count;
        }
        await output.FlushAsync(token).ConfigureAwait(false);
    }
    protected override void Dispose(bool disposing) { if (disposing) _blocks.Clear(); base.Dispose(disposing); }
    public override void Flush() => _token.ThrowIfCancellationRequested();
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => _length;
    public override long Position { get => _length; set => throw new NotSupportedException(); }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
}
