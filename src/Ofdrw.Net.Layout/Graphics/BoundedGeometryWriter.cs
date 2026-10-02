using System;
using System.IO;
using System.Text;
using System.Threading;

namespace Ofdrw.Net.Layout.Graphics;

internal sealed class BoundedGeometryWriter : TextWriter
{
    private readonly StringBuilder _buffer = new();
    private readonly long _maximum;
    private readonly CancellationToken _token;
    internal BoundedGeometryWriter(long maximum, CancellationToken token)
    {
        if (maximum < 0) throw new InvalidOperationException("Graphics geometry budget exceeded.");
        _maximum = Math.Min(maximum, int.MaxValue); _token = token;
    }
    public override Encoding Encoding => Encoding.UTF8;
    internal int Length => _buffer.Length;
    private void Reserve(int count)
    {
        _token.ThrowIfCancellationRequested();
        if (count > _maximum - _buffer.Length) throw new InvalidOperationException("Graphics geometry budget exceeded.");
    }
    public override void Write(char value) { Reserve(1); _buffer.Append(value); }
    public override void Write(string? value)
    { if (value is null) return; Reserve(value.Length); _buffer.Append(value); }
    public override void Write(char[] buffer, int index, int count)
    { Reserve(count); _buffer.Append(buffer, index, count); }
    public override string ToString() => _buffer.ToString();
}
