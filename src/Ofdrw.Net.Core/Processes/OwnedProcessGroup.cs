using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;

namespace Ofdrw.Net.Core.Processes;

// Ownership is established before the executable can run, and outlives its exit.
// Never reconstruct ownership by enumerating process names or current parent IDs.
internal abstract class OwnedProcessGroup : IDisposable
{
    internal StreamReader StandardOutput { get; private protected set; } = null!;
    internal StreamReader StandardError { get; private protected set; } = null!;
    internal abstract bool HasExited { get; }
    internal abstract int ExitCode { get; }
    internal abstract void Terminate();
    public abstract void Dispose();

    internal static OwnedProcessGroup Start(ProcessStartInfo info, TimeSpan timeout, Stopwatch elapsed, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(info.FileName)) throw new ArgumentException("An executable is required.", nameof(info));
        if (!string.IsNullOrEmpty(info.UserName) || info.RedirectStandardInput)
            throw new NotSupportedException("Owned processes do not support alternate credentials or redirected input.");
        var arguments = GetArguments(info);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return new WindowsGroup(info, arguments);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX) || RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            if (IntPtr.Size != 8)
                throw new PlatformNotSupportedException("Owned Unix processes require a validated 64-bit native ABI.");
            return new UnixGroup(info, arguments, timeout, elapsed, cancellationToken);
        }
        throw new PlatformNotSupportedException("Owned process groups require Windows, macOS or Linux.");
    }

    private protected bool WaitAttempted { get; private set; }

    internal bool WaitForExit(int milliseconds)
    {
        WaitAttempted = true;
        var elapsed = Stopwatch.StartNew();
        while (!HasExited)
        {
            if (elapsed.ElapsedMilliseconds >= milliseconds) return false;
            Thread.Sleep(10);
        }
        return true;
    }

    private static List<string> GetArguments(ProcessStartInfo info)
    {
        var list = (IEnumerable?)typeof(ProcessStartInfo).GetProperty("ArgumentList")?.GetValue(info);
        var result = list?.Cast<string>().ToList() ?? new List<string>();
        if (result.Count != 0)
        {
            if (!string.IsNullOrEmpty(info.Arguments)) throw new ArgumentException("Arguments and ArgumentList cannot both be specified.");
            return result;
        }
        // Same CRT quote/backslash rules as ProcessStartInfo.Arguments on Unix.
        var text = info.Arguments;
        var position = 0;
        while (position < text.Length)
        {
            while (position < text.Length && (text[position] == ' ' || text[position] == '\t')) position++;
            if (position == text.Length) break;
            var value = new StringBuilder();
            var quoted = false;
            while (position < text.Length)
            {
                var slashes = 0;
                while (position < text.Length && text[position] == '\\') { slashes++; position++; }
                if (position == text.Length) { value.Append('\\', slashes); break; }
                var character = text[position];
                if (character == '"')
                {
                    value.Append('\\', slashes / 2);
                    if (slashes % 2 == 1) value.Append('"');
                    else if (quoted && position + 1 < text.Length && text[position + 1] == '"')
                    { value.Append('"'); position++; }
                    else quoted = !quoted;
                }
                else
                {
                    value.Append('\\', slashes);
                    if (!quoted && (character == ' ' || character == '\t')) break;
                    value.Append(character);
                }
                position++;
            }
            result.Add(value.ToString());
        }
        return result;
    }

    private static StreamReader Reader(IntPtr handle, Encoding? encoding)
    {
        var ownedHandle = new SafePipeHandle(handle, true);
        AnonymousPipeClientStream? stream = null;
        try
        {
            stream = new AnonymousPipeClientStream(PipeDirection.In, ownedHandle);
            return new StreamReader(stream, encoding ?? Console.OutputEncoding);
        }
        catch (Exception exception)
        {
            Exception? failure = exception;
            Cleanup(ref failure, "close failed reader stream", () => stream?.Dispose());
            Cleanup(ref failure, "close failed reader handle", ownedHandle.Dispose);
            throw;
        }
    }

    private static void Cleanup(ref Exception? failure, string operation, Action action)
    {
        try { action(); }
        catch (Exception exception)
        {
            if (failure == null) failure = exception;
            else failure.Data["OwnedProcessCleanup." + operation] = exception.ToString();
        }
    }

    private static Win32Exception NativeError(string operation) => new Win32Exception(Marshal.GetLastWin32Error(), operation);

    private sealed class NativeStrings : IDisposable
    {
        internal IntPtr Pointer { get; }
        private readonly List<IntPtr> _values = new List<IntPtr>();
        internal NativeStrings(IEnumerable<string> values)
        {
            try
            {
                foreach (var value in values)
                {
                    if (value.IndexOf('\0') >= 0) throw new ArgumentException("Process arguments/environment cannot contain NUL.");
                    var bytes = Encoding.UTF8.GetBytes(value + "\0");
                    var address = Marshal.AllocHGlobal(bytes.Length);
                    _values.Add(address);
                    Marshal.Copy(bytes, 0, address, bytes.Length);
                }
                Pointer = Marshal.AllocHGlobal((_values.Count + 1) * IntPtr.Size);
                for (var i = 0; i < _values.Count; i++) Marshal.WriteIntPtr(Pointer, i * IntPtr.Size, _values[i]);
                Marshal.WriteIntPtr(Pointer, _values.Count * IntPtr.Size, IntPtr.Zero);
            }
            catch { Dispose(); throw; }
        }
        public void Dispose()
        {
            foreach (var value in _values) Marshal.FreeHGlobal(value);
            if (Pointer != IntPtr.Zero) Marshal.FreeHGlobal(Pointer);
        }
    }

    private sealed class UnixGroup : OwnedProcessGroup
    {
        private static readonly object LaunchLock = new object();
        private static readonly bool Darwin = RuntimeInformation.IsOSPlatform(OSPlatform.OSX);
        private const int ProtocolTimeout = 6000;
        private readonly int _pid;
        private readonly string _executable;
        private readonly IntPtr _frame = Marshal.AllocHGlobal(16);
        private NativeSupervisorAssets.Extraction? _extraction;
        private int _protocol = -1;
        private int _frameLength;
        private Stopwatch? _partialFrame;
        private int _targetPid;
        private bool _started, _terminated, _disposed, _exited;
        private int _exitCode;
        private readonly Stopwatch _startElapsed;
        private readonly TimeSpan _startTimeout;
        private bool _starting = true, _targetStartError;
        private Win32Exception? _startFailure;

        internal UnixGroup(ProcessStartInfo info, List<string> arguments, TimeSpan timeout, Stopwatch elapsed, CancellationToken cancellationToken)
        {
            _executable = info.FileName;
            _startElapsed = elapsed; _startTimeout = timeout;
            var attributes = IntPtr.Zero;
            var actions = IntPtr.Zero;
            var attributesReady = false;
            var actionsReady = false;
            var descriptors = new List<int>();
            Exception? failure = null;
            var resolving = true;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                CheckStartBudget();
                var executable = ResolveExecutable(info.FileName);
                resolving = false;
                EnterLaunchLock(cancellationToken);
                try
                {
                    // No owned spawn may overlap another launch's
                    // temporary executable writer.
                    _extraction = NativeSupervisorAssets.Extract();
                    attributes = Marshal.AllocHGlobal(Darwin ? IntPtr.Size : 336);
                    actions = Marshal.AllocHGlobal(Darwin ? IntPtr.Size : 80);
                    using var argv = new NativeStrings(new[] { _extraction.ExecutablePath, executable, info.FileName }.Concat(arguments));
                    using var environment = new NativeStrings(info.Environment.Where(pair => pair.Value != null)
                        .Select(pair => pair.Key + "=" + pair.Value));
                    Check(posix_spawnattr_init(attributes), "posix_spawnattr_init"); attributesReady = true;
                    Check(posix_spawn_file_actions_init(actions), "posix_spawn_file_actions_init"); actionsReady = true;
                    Check(posix_spawnattr_setpgroup(attributes, 0), "posix_spawnattr_setpgroup");
                    Check(posix_spawnattr_setflags(attributes, (short)(2 | (Darwin ? 0x4000 : 0))), "posix_spawnattr_setflags");
                    if (!string.IsNullOrEmpty(info.WorkingDirectory))
                    {
                        try { Check(posix_spawn_file_actions_addchdir_np(actions, info.WorkingDirectory), "posix_spawn chdir"); }
                        catch (EntryPointNotFoundException exception)
                        {
                            throw new PlatformNotSupportedException("Owned processes with a working directory require posix_spawn_file_actions_addchdir_np (macOS 10.15+ or Linux libc providing this extension).", exception);
                        }
                    }
                    var control = CreateSocket(descriptors);
                    var output = CreatePipe(descriptors);
                    var error = CreatePipe(descriptors);
                    Check(posix_spawn_file_actions_adddup2(actions, control[1], 3), "posix_spawn protocol");
                    Check(posix_spawn_file_actions_adddup2(actions, output[1], 1), "posix_spawn stdout");
                    Check(posix_spawn_file_actions_adddup2(actions, error[1], 2), "posix_spawn stderr");
                    foreach (var descriptor in descriptors)
                        Check(posix_spawn_file_actions_addclose(actions, descriptor), "posix_spawn close");
                    cancellationToken.ThrowIfCancellationRequested();
                    CheckStartBudget();
                    var pid = 0;
                    int Spawn() => posix_spawn(out pid, _extraction.ExecutablePath, actions, attributes, argv.Pointer, environment.Pointer);
                    SupervisorSpawnAdmission? admission = null;
                    int spawnError;
                    if (Darwin) spawnError = Spawn();
                    else
                    {
                        var result = SupervisorExecutableBusyRetry.Start(Spawn, CheckStartBudget,
                            () => _startTimeout - _startElapsed.Elapsed, cancellationToken);
                        admission = result;
                        spawnError = result.Error;
                    }
                    if (spawnError != 0)
                    {
                        var spawnFailure = new IOException("Cannot start the owned-process supervisor (the temporary filesystem must permit execution).", new Win32Exception(spawnError));
                        admission?.AttachTo(spawnFailure);
                        throw spawnFailure;
                    }
                    _pid = pid;
                    descriptors.Remove(control[0]); _protocol = control[0];
                    SetNonblocking(_protocol);
                    foreach (var descriptor in new[] { control[1], output[1], error[1] })
                    { descriptors.Remove(descriptor); CloseDescriptor(descriptor); }
                    descriptors.Remove(output[0]); StandardOutput = Reader((IntPtr)output[0], info.StandardOutputEncoding);
                    descriptors.Remove(error[0]); StandardError = Reader((IntPtr)error[0], info.StandardErrorEncoding);
                }
                finally { Monitor.Exit(LaunchLock); }
                var protocolElapsed = Stopwatch.StartNew();
                while (!_started) ReceiveFrame(protocolElapsed, ProtocolTimeout, cancellationToken, true);
                cancellationToken.ThrowIfCancellationRequested();
                CheckStartBudget();
                _starting = false;
            }
            catch (Win32Exception exception) when (!resolving && !_targetStartError) { failure = ProtocolError("Cannot initialize the owned-process supervisor.", exception.NativeErrorCode); }
            catch (Exception exception)
            {
                failure = _startFailure ?? exception;
                if (_startFailure != null && !ReferenceEquals(exception, _startFailure))
                    failure.Data["OwnedProcessCleanup.start failure"] = exception.ToString();
            }
            finally
            {
                foreach (var descriptor in descriptors) Cleanup(ref failure, "close fd " + descriptor, () => CloseDescriptor(descriptor));
                if (actionsReady) Cleanup(ref failure, "destroy spawn actions", () => Check(posix_spawn_file_actions_destroy(actions), "destroy spawn actions"));
                if (attributesReady) Cleanup(ref failure, "destroy spawn attributes", () => Check(posix_spawnattr_destroy(attributes), "destroy spawn attributes"));
                Cleanup(ref failure, "free spawn actions", () => Marshal.FreeHGlobal(actions));
                Cleanup(ref failure, "free spawn attributes", () => Marshal.FreeHGlobal(attributes));
            }
            if (failure != null)
            {
                // EOF asks the native owner to clean up, including a launch that
                // has not yet reached exec. Never signal a cached numeric PID.
                Cleanup(ref failure, "close supervisor protocol", CloseProtocol);
                Cleanup(ref failure, "reap supervisor", () => ReapSupervisor(ProtocolTimeout));
                Cleanup(ref failure, "close stdout", () => StandardOutput?.Dispose());
                Cleanup(ref failure, "close stderr", () => StandardError?.Dispose());
                Cleanup(ref failure, "remove supervisor executable", () => _extraction?.Dispose());
                Cleanup(ref failure, "free process frame", () => Marshal.FreeHGlobal(_frame));
                ExceptionDispatchInfo.Capture(failure!).Throw();
            }
        }

        internal override bool HasExited
        {
            get
            {
                if (!_exited) ReceiveFrame(Stopwatch.StartNew(), 0, default, false);
                return _exited;
            }
        }
        internal override int ExitCode => HasExited ? _exitCode : throw new InvalidOperationException("Process has not exited.");

        internal override void Terminate()
        {
            if (_terminated) return;
            var elapsed = Stopwatch.StartNew();
            var command = new[] { (byte)'K' };
            while (true)
            {
                if (elapsed.ElapsedMilliseconds >= ProtocolTimeout) throw ProtocolError("Timed out requesting process cleanup.");
                var written = send(_protocol, command, (UIntPtr)1, Darwin ? 0 : 0x4000).ToInt64();
                if (written == 1) break;
                var error = Marshal.GetLastWin32Error();
                if (written >= 0 || (error != 4 && error != (Darwin ? 35 : 11)))
                    throw ProtocolError("Cannot request process cleanup.", error);
                Poll(4, elapsed, ProtocolTimeout, default, true);
            }
            while (!_terminated) ReceiveFrame(elapsed, ProtocolTimeout, default, true);
        }

        public override void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Exception? failure = null;
            Cleanup(ref failure, "terminate", Terminate);
            Cleanup(ref failure, "close supervisor protocol", CloseProtocol);
            Cleanup(ref failure, "reap supervisor", () => ReapSupervisor(failure == null ? 1000 : ProtocolTimeout));
            Cleanup(ref failure, "close stdout", StandardOutput.Dispose);
            Cleanup(ref failure, "close stderr", StandardError.Dispose);
            Cleanup(ref failure, "remove supervisor executable", () => _extraction?.Dispose());
            Cleanup(ref failure, "free process frame", () => Marshal.FreeHGlobal(_frame));
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }

        private void ReceiveFrame(Stopwatch elapsed, int milliseconds, CancellationToken cancellationToken, bool blocking)
        {
            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_partialFrame != null && _partialFrame.ElapsedMilliseconds >= ProtocolTimeout)
                    throw ProtocolError("Incomplete supervisor frame exceeded its deadline.");
                if (!Poll(1, elapsed, milliseconds, cancellationToken, blocking)) return;
                var received = read(_protocol, IntPtr.Add(_frame, _frameLength), (UIntPtr)(16 - _frameLength)).ToInt64();
                if (received == 0) throw ProtocolError("The owned-process supervisor closed its protocol unexpectedly; target cleanup could not be confirmed.");
                if (received < 0)
                {
                    var error = Marshal.GetLastWin32Error();
                    if (error == 4 || error == (Darwin ? 35 : 11)) continue;
                    throw ProtocolError("Cannot read owned-process supervisor protocol.", error);
                }
                if (_frameLength == 0) _partialFrame = Stopwatch.StartNew();
                _frameLength += (int)received;
                if (_frameLength != 16) continue;
                var bytes = new byte[16]; Marshal.Copy(_frame, bytes, 0, bytes.Length);
                _frameLength = 0; _partialFrame = null;
                int Word(int index) => unchecked((int)((uint)bytes[index] << 24 | (uint)bytes[index + 1] << 16 | (uint)bytes[index + 2] << 8 | bytes[index + 3]));
                var kind = Word(4); var target = Word(8); var value = Word(12);
                if (Word(0) != 0x4f465301 || target < 0 || target == _pid || (_targetPid != 0 && target != _targetPid))
                    throw ProtocolError("Invalid supervisor frame identity or version.");
                if (kind == 5 && value > 0) throw ProtocolError("Owned-process supervisor failed.", value);
                if (target == 0) throw ProtocolError("Supervisor did not identify its target.");
                if (_targetPid == 0) _targetPid = target;
                switch (kind)
                {
                    case 1 when !_started && _startFailure == null && value == 0: _started = true; return;
                    case 2 when _started && !_exited && value >= 0 && value <= 255: _exitCode = value; _exited = true; return;
                    case 3 when !_started && _startFailure == null && value > 0:
                        _targetStartError = true;
                        _startFailure = new Win32Exception(value, "execve " + _executable);
                        return;
                    case 4 when value == 0 && ((_started && _exited) || _startFailure != null):
                        _terminated = true;
                        if (_startFailure != null) throw _startFailure;
                        return;
                    default: throw ProtocolError("Invalid supervisor protocol state.");
                }
            } while (blocking);
        }

        private bool Poll(short events, Stopwatch elapsed, int milliseconds, CancellationToken cancellationToken, bool blocking)
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_starting) CheckStartBudget();
                var remaining = milliseconds - elapsed.ElapsedMilliseconds;
                if (blocking && remaining <= 0) throw ProtocolError("Owned-process supervisor exceeded its protocol deadline.");
                var descriptors = new[] { new PollDescriptor { Descriptor = _protocol, Events = events } };
                var result = poll(descriptors, (UIntPtr)1, blocking ? (int)Math.Min(25, remaining) : 0);
                if (result > 0) return true;
                if (result == 0) { if (!blocking) return false; continue; }
                var error = Marshal.GetLastWin32Error();
                if (error != 4) throw ProtocolError("Cannot poll owned-process supervisor.", error);
            }
        }
        private void EnterLaunchLock(CancellationToken cancellationToken)
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CheckStartBudget();
                var remaining = _startTimeout - _startElapsed.Elapsed;
                var wait = (int)Math.Max(1, Math.Min(25, Math.Ceiling(remaining.TotalMilliseconds)));
                if (!Monitor.TryEnter(LaunchLock, wait)) continue;
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    CheckStartBudget();
                    return;
                }
                catch { Monitor.Exit(LaunchLock); throw; }
            }
        }
        private void CheckStartBudget()
        {
            if (_startElapsed.Elapsed >= _startTimeout)
                throw new TimeoutException($"{_executable} exceeded {_startTimeout} during process startup.");
        }
        private static IOException ProtocolError(string message, int error = 0) =>
            error == 0 ? new IOException(message) : new IOException(message, new Win32Exception(error));
        private void CloseProtocol()
        {
            if (_protocol < 0) return;
            var descriptor = _protocol; _protocol = -1; CloseDescriptor(descriptor);
        }
        private void ReapSupervisor(int milliseconds)
        {
            if (_pid == 0) return;
            var elapsed = Stopwatch.StartNew();
            while (true)
            {
                var result = waitpid(_pid, out _, 1);
                if (result == _pid || (result < 0 && Marshal.GetLastWin32Error() == 10)) return;
                if (result < 0 && Marshal.GetLastWin32Error() != 4) throw ProtocolError("Cannot reap supervisor.", Marshal.GetLastWin32Error());
                if (elapsed.ElapsedMilliseconds >= milliseconds)
                {
                    throw ProtocolError("Supervisor cleanup did not finish within its deadline.");
                }
                Thread.Sleep(10);
            }
        }
        private static void CloseDescriptor(int descriptor) { if (close(descriptor) != 0) throw NativeError("close process descriptor"); }
        private static void SetNonblocking(int descriptor)
        {
            var flags = fcntl(descriptor, 3, 0);
            if (flags < 0 || fcntl(descriptor, 4, flags | (Darwin ? 4 : 0x800)) != 0) throw NativeError("set process protocol nonblocking");
        }
        private static int[] Normalize(int[] pair, List<int> descriptors)
        {
            descriptors.AddRange(pair);
            for (var index = 0; index < pair.Length; index++)
            {
                if (pair[index] >= 10) continue;
                var original = pair[index];
                var duplicate = fcntl(original, Darwin ? 67 : 1030, 10);
                if (duplicate < 0) throw NativeError("reserve process descriptor");
                descriptors.Add(duplicate); pair[index] = duplicate;
                descriptors.Remove(original); CloseDescriptor(original);
            }
            return pair;
        }
        private static int[] CreatePipe(List<int> descriptors)
        {
            var pair = new int[2];
            if ((Darwin ? pipe(pair) : pipe2(pair, 0x80000)) != 0) throw NativeError("pipe");
            Normalize(pair, descriptors);
            if (Darwin) foreach (var descriptor in pair) if (fcntl(descriptor, 2, 1) < 0) throw NativeError("pipe close-on-exec");
            return pair;
        }
        private static int[] CreateSocket(List<int> descriptors)
        {
            var pair = new int[2];
            if (socketpair(1, 1 | (Darwin ? 0 : 0x80000), 0, pair) != 0) throw NativeError("process socketpair");
            Normalize(pair, descriptors);
            if (Darwin)
            {
                foreach (var descriptor in pair) if (fcntl(descriptor, 2, 1) < 0) throw NativeError("socket close-on-exec");
                var enabled = 1;
                if (setsockopt(pair[0], 0xffff, 0x1022, ref enabled, sizeof(int)) != 0) throw NativeError("socket no SIGPIPE");
            }
            return pair;
        }
        private static string ResolveExecutable(string name)
        {
            if (Path.IsPathRooted(name)) return name;
            if (name.IndexOf(Path.DirectorySeparatorChar) >= 0) return Path.GetFullPath(name);
            foreach (var root in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() }
                .Concat((Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator)))
            {
                var candidate = Path.GetFullPath(Path.Combine(root, name));
                if (File.Exists(candidate) && access(candidate, 1) == 0) return candidate;
            }
            throw new Win32Exception(2, "Executable not found: " + name);
        }
        private static void Check(int error, string operation) { if (error != 0) throw new Win32Exception(error, operation); }
        [StructLayout(LayoutKind.Sequential)] private struct PollDescriptor { internal int Descriptor; internal short Events; internal short ReturnedEvents; }
        [DllImport("libc", SetLastError = true)] private static extern int pipe(int[] descriptors);
        [DllImport("libc", SetLastError = true)] private static extern int pipe2(int[] descriptors, int flags);
        [DllImport("libc", SetLastError = true)] private static extern int socketpair(int domain, int type, int protocol, int[] pair);
        [DllImport("libc", SetLastError = true)] private static extern int setsockopt(int socket, int level, int option, ref int value, int length);
        [DllImport("libc", SetLastError = true)] private static extern int fcntl(int fd, int command, int value);
        [DllImport("libc", SetLastError = true)] private static extern int close(int fd);
        [DllImport("libc", SetLastError = true)] private static extern int access(string path, int mode);
        [DllImport("libc", SetLastError = true)] private static extern IntPtr send(int fd, byte[] data, UIntPtr length, int flags);
        [DllImport("libc", SetLastError = true)] private static extern IntPtr read(int fd, IntPtr data, UIntPtr length);
        [DllImport("libc", SetLastError = true)] private static extern int poll(PollDescriptor[] descriptors, UIntPtr count, int milliseconds);
        [DllImport("libc", SetLastError = true)] private static extern int waitpid(int pid, out int status, int options);
        [DllImport("libc")] private static extern int posix_spawnattr_init(IntPtr attributes);
        [DllImport("libc")] private static extern int posix_spawnattr_destroy(IntPtr attributes);
        [DllImport("libc")] private static extern int posix_spawnattr_setpgroup(IntPtr attributes, int group);
        [DllImport("libc")] private static extern int posix_spawnattr_setflags(IntPtr attributes, short flags);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_init(IntPtr actions);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_destroy(IntPtr actions);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_addchdir_np(IntPtr actions, string directory);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_adddup2(IntPtr actions, int source, int target);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_addclose(IntPtr actions, int descriptor);
        [DllImport("libc")] private static extern int posix_spawn(out int pid, string path, IntPtr actions, IntPtr attributes, IntPtr argv, IntPtr envp);
    }

    private sealed class WindowsGroup : OwnedProcessGroup
    {
        private IntPtr _job;
        private IntPtr _process;
        private bool _terminated;
        private bool _disposed;

        internal WindowsGroup(ProcessStartInfo info, List<string> arguments)
        {
            var handles = new List<IntPtr>();
            var attributes = IntPtr.Zero;
            var attributeHandles = IntPtr.Zero;
            var environment = IntPtr.Zero;
            var thread = IntPtr.Zero;
            var initialized = false;
            Exception? failure = null;
            try
            {
                _job = CreateJobObjectW(IntPtr.Zero, null);
                if (_job == IntPtr.Zero) throw NativeError("CreateJobObject");
                var limits = new JobLimits(); limits.Basic.LimitFlags = 0x2000; // KILL_ON_JOB_CLOSE
                if (!SetInformationJobObject(_job, 9, ref limits, (uint)Marshal.SizeOf<JobLimits>())) throw NativeError("SetInformationJobObject");
                var security = new SecurityAttributes { Length = Marshal.SizeOf<SecurityAttributes>(), Inherit = 1 };
                var output = CreatePipe(handles, ref security);
                var error = CreatePipe(handles, ref security);
                var originalInput = GetStdHandle(-10);
                IntPtr input;
                if (originalInput == IntPtr.Zero || originalInput == new IntPtr(-1))
                {
                    input = CreateFileW("NUL", 0x80000000, 3, ref security, 3, 0, IntPtr.Zero);
                    if (input == new IntPtr(-1)) throw NativeError("Open NUL input");
                }
                else if (!DuplicateHandle(GetCurrentProcess(), originalInput, GetCurrentProcess(), out input, 0, true, 2))
                    throw NativeError("Duplicate standard input");
                handles.Add(input);
                var size = IntPtr.Zero;
                InitializeProcThreadAttributeList(IntPtr.Zero, 1, 0, ref size);
                attributes = Marshal.AllocHGlobal(size);
                if (!InitializeProcThreadAttributeList(attributes, 1, 0, ref size)) throw NativeError("InitializeProcThreadAttributeList");
                initialized = true;
                attributeHandles = Marshal.AllocHGlobal(IntPtr.Size * 3);
                Marshal.WriteIntPtr(attributeHandles, 0, input);
                Marshal.WriteIntPtr(attributeHandles, IntPtr.Size, output[1]);
                Marshal.WriteIntPtr(attributeHandles, IntPtr.Size * 2, error[1]);
                if (!UpdateProcThreadAttribute(attributes, 0, (IntPtr)0x20002, attributeHandles, (IntPtr)(IntPtr.Size * 3), IntPtr.Zero, IntPtr.Zero))
                    throw NativeError("Restrict inherited handles");
                var startup = new StartupInfoEx { Attributes = attributes };
                startup.Startup.Size = Marshal.SizeOf<StartupInfoEx>();
                startup.Startup.Flags = 0x100; // STARTF_USESTDHANDLES
                startup.Startup.Input = input; startup.Startup.Output = output[1]; startup.Startup.Error = error[1];
                var command = new StringBuilder(Quote(info.FileName));
                // Windows Process.Start sends Arguments verbatim to the target.
                if (!string.IsNullOrEmpty(info.Arguments)) command.Append(' ').Append(info.Arguments);
                else foreach (var argument in arguments) command.Append(' ').Append(Quote(argument));
                var environmentText = string.Join("\0", info.Environment.Where(pair => pair.Value != null)
                    .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase).Select(pair => pair.Key + "=" + pair.Value)) + "\0\0";
                environment = Marshal.StringToHGlobalUni(environmentText);
                if (!CreateProcessW(null, command, IntPtr.Zero, IntPtr.Zero, true, 0x80000 | 0x400 | 0x4 | 0x08000000,
                    environment, string.IsNullOrEmpty(info.WorkingDirectory) ? null : info.WorkingDirectory, ref startup, out var process))
                    throw NativeError("CreateProcess " + info.FileName);
                _process = process.Process; thread = process.Thread;
                if (!AssignProcessToJobObject(_job, _process)) throw NativeError("AssignProcessToJobObject");
                // CreateProcess was suspended: no target code/descendant can escape assignment.
                if (ResumeThread(thread) == uint.MaxValue) throw NativeError("ResumeThread");
                handles.Remove(output[1]); CloseNativeHandle(output[1]);
                handles.Remove(error[1]); CloseNativeHandle(error[1]);
                handles.Remove(output[0]); StandardOutput = Reader(output[0], info.StandardOutputEncoding);
                handles.Remove(error[0]); StandardError = Reader(error[0], info.StandardErrorEncoding);
            }
            catch (Exception exception) { failure = exception; }
            finally
            {
                foreach (var handle in handles) Cleanup(ref failure, "close launch handle " + handle, () => CloseNativeHandle(handle));
                if (thread != IntPtr.Zero) Cleanup(ref failure, "close thread", () => CloseNativeHandle(thread));
                if (initialized) Cleanup(ref failure, "delete launch attributes", () => DeleteProcThreadAttributeList(attributes));
                Cleanup(ref failure, "free launch attributes", () => Marshal.FreeHGlobal(attributes));
                Cleanup(ref failure, "free launch handle list", () => Marshal.FreeHGlobal(attributeHandles));
                Cleanup(ref failure, "free environment", () => Marshal.FreeHGlobal(environment));
            }
            if (failure != null)
            {
                if (_process != IntPtr.Zero)
                {
                    Cleanup(ref failure, "terminate failed launch", () =>
                    {
                        if (!HasExited && !TerminateProcess(_process, 1)) throw NativeError("TerminateProcess after failed launch");
                    });
                }
                // Closing the Job also cleans descendants if a failure occurred
                // after ResumeThread. Always attempt it, even if termination failed.
                if (_job != IntPtr.Zero) Cleanup(ref failure, "close failed launch job", () => CloseNativeHandle(_job));
                if (_process != IntPtr.Zero)
                {
                    Cleanup(ref failure, "wait failed launch", () =>
                    {
                        var status = WaitForSingleObject(_process, 5000);
                        if (status == uint.MaxValue) throw NativeError("wait failed launch");
                        if (status != 0) throw new TimeoutException("Failed launch did not terminate within the cleanup deadline.");
                    });
                    Cleanup(ref failure, "close failed launch process", () => CloseNativeHandle(_process));
                }
                Cleanup(ref failure, "close stdout", () => StandardOutput?.Dispose());
                Cleanup(ref failure, "close stderr", () => StandardError?.Dispose());
                ExceptionDispatchInfo.Capture(failure!).Throw();
            }
        }
        internal override bool HasExited
        {
            get
            {
                var status = WaitForSingleObject(_process, 0);
                if (status == uint.MaxValue) throw NativeError("WaitForSingleObject");
                return status == 0;
            }
        }
        internal override int ExitCode
        {
            get
            {
                if (!HasExited) throw new InvalidOperationException("Process has not exited.");
                if (!GetExitCodeProcess(_process, out var code)) throw NativeError("GetExitCodeProcess");
                return unchecked((int)code);
            }
        }
        internal override void Terminate()
        {
            if (_terminated) return;
            if (!TerminateJobObject(_job, 1)) throw NativeError("TerminateJobObject");
            _terminated = true;
        }
        public override void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Exception? failure = null;
            Cleanup(ref failure, "terminate", Terminate);
            Cleanup(ref failure, "wait", () =>
            {
                if (!(WaitAttempted ? HasExited : WaitForExit(5000)))
                    throw new TimeoutException("Owned process did not exit within the cleanup deadline.");
            });
            Cleanup(ref failure, "close job", () => CloseNativeHandle(_job));
            Cleanup(ref failure, "close process", () => CloseNativeHandle(_process));
            Cleanup(ref failure, "close stdout", StandardOutput.Dispose);
            Cleanup(ref failure, "close stderr", StandardError.Dispose);
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
        private static void CloseNativeHandle(IntPtr handle)
        {
            if (!CloseHandle(handle)) throw NativeError("CloseHandle for owned process");
        }
        private static IntPtr[] CreatePipe(List<IntPtr> handles, ref SecurityAttributes security)
        {
            if (!CreatePipe(out var read, out var write, ref security, 0)) throw NativeError("CreatePipe");
            handles.Add(read); handles.Add(write);
            if (!SetHandleInformation(read, 1, 0)) throw NativeError("SetHandleInformation");
            return new[] { read, write };
        }
        private static string Quote(string value)
        {
            if (value.IndexOf('\0') >= 0) throw new ArgumentException("Process arguments cannot contain NUL.");
            var quoted = new StringBuilder("\"");
            var slashes = 0;
            foreach (var character in value)
            {
                if (character == '\\') { slashes++; continue; }
                quoted.Append('\\', character == '"' ? slashes * 2 + 1 : slashes).Append(character);
                slashes = 0;
            }
            return quoted.Append('\\', slashes * 2).Append('"').ToString();
        }
        [StructLayout(LayoutKind.Sequential)] private struct SecurityAttributes { internal int Length; internal IntPtr Descriptor; internal int Inherit; }
        [StructLayout(LayoutKind.Sequential)] private struct BasicLimits
        { internal long ProcessTime, JobTime; internal uint LimitFlags; internal UIntPtr MinWorkingSet, MaxWorkingSet; internal uint ActiveProcessLimit; internal UIntPtr Affinity; internal uint Priority, Scheduling; }
        [StructLayout(LayoutKind.Sequential)] private struct JobLimits
        { internal BasicLimits Basic; internal ulong ReadOperations, WriteOperations, OtherOperations, ReadBytes, WriteBytes, OtherBytes; internal UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory; }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct StartupInfo
        { internal int Size; internal string? Reserved, Desktop, Title; internal uint X, Y, Width, Height, XChars, YChars, Fill, Flags; internal short Show, ReservedSize; internal IntPtr ReservedBytes, Input, Output, Error; }
        [StructLayout(LayoutKind.Sequential)] private struct StartupInfoEx { internal StartupInfo Startup; internal IntPtr Attributes; }
        [StructLayout(LayoutKind.Sequential)] private struct ProcessInformation { internal IntPtr Process, Thread; internal uint Pid, Tid; }
        [DllImport("kernel32")] private static extern IntPtr GetStdHandle(int type);
        [DllImport("kernel32")] private static extern IntPtr GetCurrentProcess();
        [DllImport("kernel32", SetLastError = true)] private static extern bool DuplicateHandle(IntPtr sourceProcess, IntPtr source, IntPtr targetProcess, out IntPtr target, uint access, bool inherit, uint options);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateJobObjectW(IntPtr security, string? name);
        [DllImport("kernel32", SetLastError = true)] private static extern bool SetInformationJobObject(IntPtr job, int type, ref JobLimits limits, uint length);
        [DllImport("kernel32", SetLastError = true)] private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32", SetLastError = true)] private static extern bool TerminateJobObject(IntPtr job, uint code);
        [DllImport("kernel32", SetLastError = true)] private static extern bool TerminateProcess(IntPtr process, uint code);
        [DllImport("kernel32", SetLastError = true)] private static extern bool CreatePipe(out IntPtr read, out IntPtr write, ref SecurityAttributes security, uint size);
        [DllImport("kernel32", SetLastError = true)] private static extern bool SetHandleInformation(IntPtr handle, uint mask, uint flags);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateFileW(string name, uint access, uint share, ref SecurityAttributes security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32", SetLastError = true)] private static extern bool InitializeProcThreadAttributeList(IntPtr attributes, int count, int flags, ref IntPtr size);
        [DllImport("kernel32", SetLastError = true)] private static extern bool UpdateProcThreadAttribute(IntPtr attributes, uint flags, IntPtr attribute, IntPtr value, IntPtr size, IntPtr previous, IntPtr returned);
        [DllImport("kernel32")] private static extern void DeleteProcThreadAttributeList(IntPtr attributes);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CreateProcessW(string? application, StringBuilder command, IntPtr processSecurity, IntPtr threadSecurity, bool inherit, uint flags, IntPtr environment, string? directory, ref StartupInfoEx startup, out ProcessInformation process);
        [DllImport("kernel32", SetLastError = true)] private static extern uint ResumeThread(IntPtr thread);
        [DllImport("kernel32", SetLastError = true)] private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
        [DllImport("kernel32", SetLastError = true)] private static extern bool GetExitCodeProcess(IntPtr process, out uint code);
        [DllImport("kernel32")] private static extern bool CloseHandle(IntPtr handle);
    }
}
