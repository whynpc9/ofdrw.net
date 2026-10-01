using System;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// One decoded resource at a time. Eviction precedes construction so distinct payloads cannot accumulate decoded buffers.
internal sealed class SinglePayloadResource<T> : IDisposable where T : class, IDisposable
{
    private readonly Func<byte[], T> _create;
    private byte[]? _key;
    private T? _resource;
    internal SinglePayloadResource(Func<byte[], T> create) => _create = create;
    internal T Get(byte[] key)
    {
        if (_resource is not null && ReferenceEquals(key, _key)) return _resource;
        Dispose();
        var resource = _create(key);
        _resource = resource;
        _key = key;
        return resource;
    }
    public void Dispose()
    {
        _resource?.Dispose();
        _resource = null;
        _key = null;
    }
}
