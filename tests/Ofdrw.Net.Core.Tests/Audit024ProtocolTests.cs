using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using Ofdrw.Net.Core.Processes;

namespace Ofdrw.Net.Core.Tests;

public sealed class Audit024ProtocolTests
{
    [UnixProtocolFact]
    public void Audit024_StartupProtocolHonorsShortUserDeadline()
    {
        using var peer = new ProtocolPeer(TimeSpan.FromMilliseconds(350));
        var elapsed = Stopwatch.StartNew();
        var failure = Assert.Throws<TimeoutException>(() => peer.Receive(default));
        Assert.Contains("startup", failure.Message);
        Assert.InRange(elapsed.ElapsedMilliseconds, 250, 1500);
    }

    [UnixProtocolFact]
    public void Audit024_StartupProtocolCancellationIsBounded()
    {
        using var peer = new ProtocolPeer(TimeSpan.FromSeconds(10));
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(150));
        var elapsed = Stopwatch.StartNew();
        Assert.ThrowsAny<OperationCanceledException>(() => peer.Receive(cancellation.Token));
        Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(1));
    }

    [UnixProtocolFact]
    public void Audit024_PartialFrameAndEofCannotProduceSuccessfulExit()
    {
        using var peer = new ProtocolPeer(TimeSpan.FromSeconds(10));
        peer.Write(new byte[] { 0x4f, 0x46, 0x53, 1, 0, 0, 0, 1 });
        Assert.False(peer.HasExited());
        peer.ClosePeer();
        Assert.Throws<IOException>(() => peer.HasExited());
    }

    // Construct only the protocol receiver, with test-owned socket endpoints.
    // No process is launched and no runtime-private Process APIs are used. This
    // lets the real poll/read path see slow, partial and closed peer streams.
    private sealed class ProtocolPeer : IDisposable
    {
        private readonly Type _type = typeof(OwnedProcessGroup).GetNestedType("UnixGroup", BindingFlags.NonPublic)!;
        private readonly object _receiver;
        private readonly int _socket;
        private int _peer;
        private readonly IntPtr _frame = Marshal.AllocHGlobal(16);
        internal ProtocolPeer(TimeSpan timeout)
        {
            var sockets = new int[2];
            Assert.Equal(0, socketpair(1, 1, 0, sockets));
            _socket = sockets[0]; _peer = sockets[1];
            _receiver = RuntimeHelpers.GetUninitializedObject(_type);
            Set("_protocol", _socket); Set("_frame", _frame);
            Set("_pid", Environment.ProcessId); Set("_executable", "protocol-probe");
            Set("_startElapsed", Stopwatch.StartNew()); Set("_startTimeout", timeout); Set("_starting", true);
        }
        private void Set(string name, object value) => _type.GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(_receiver, value);
        internal void Receive(CancellationToken cancellationToken) => Invoke("ReceiveFrame", Stopwatch.StartNew(), 6000, cancellationToken, true);
        internal bool HasExited() => (bool)Unwrap(() => _type.GetProperty("HasExited", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(_receiver))!;
        private void Invoke(string method, params object[] arguments) => Unwrap(() => _type.GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(_receiver, arguments));
        private static object? Unwrap(Func<object?> action)
        {
            try { return action(); }
            catch (TargetInvocationException exception) { ExceptionDispatchInfo.Capture(exception.InnerException!).Throw(); throw; }
        }
        internal void Write(byte[] bytes) => Assert.Equal(bytes.Length, write(_peer, bytes, (UIntPtr)bytes.Length).ToInt32());
        internal void ClosePeer() { if (_peer >= 0) { close(_peer); _peer = -1; } }
        public void Dispose() { ClosePeer(); close(_socket); Marshal.FreeHGlobal(_frame); }
        [DllImport("libc")] private static extern int socketpair(int domain, int type, int protocol, int[] pair);
        [DllImport("libc")] private static extern IntPtr write(int fd, byte[] data, UIntPtr count);
        [DllImport("libc")] private static extern int close(int fd);
    }

    private sealed class UnixProtocolFactAttribute : FactAttribute
    {
        public UnixProtocolFactAttribute() { if (OperatingSystem.IsWindows()) Skip = "Unix socket protocol; Windows retains the Job backend."; }
    }
}
