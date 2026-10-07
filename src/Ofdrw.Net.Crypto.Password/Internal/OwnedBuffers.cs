using System.Security.Cryptography;

namespace Ofdrw.Net.Crypto.Password.Internal;

// Only arrays acquired or created by this extension enter this owner. Caller
// streams and the complete returned ZIP are never registered here.
internal sealed class OwnedBuffers : IDisposable
{
    private readonly List<byte[]> buffers = new();
    internal void Add(byte[] bytes) => buffers.Add(bytes);
    internal void AddRange(IEnumerable<byte[]> values) => buffers.AddRange(values);
    public void Dispose()
    {
        foreach (var bytes in buffers) CryptographicOperations.ZeroMemory(bytes);
        buffers.Clear();
    }
}
