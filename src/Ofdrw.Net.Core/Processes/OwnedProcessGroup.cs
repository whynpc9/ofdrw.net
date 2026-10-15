using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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

    internal static OwnedProcessGroup Start(ProcessStartInfo info)
    {
        if (string.IsNullOrWhiteSpace(info.FileName)) throw new ArgumentException("An executable is required.", nameof(info));
        if (!string.IsNullOrEmpty(info.UserName) || info.RedirectStandardInput)
            throw new NotSupportedException("Owned processes do not support alternate credentials or redirected input.");
        var arguments = GetArguments(info);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return new WindowsGroup(info, arguments);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX) || RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            return new UnixGroup(info, arguments);
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

    private static StreamReader Reader(IntPtr handle, Encoding? encoding) => new StreamReader(
        new AnonymousPipeClientStream(PipeDirection.In, new SafePipeHandle(handle, true)), encoding ?? Console.OutputEncoding);

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
        private readonly int _pid;
        private readonly IntPtr _status = Marshal.AllocHGlobal(128);
        private bool _terminated;
        private bool _disposed;
        private bool _exited;
        private int _exitCode;

        internal UnixGroup(ProcessStartInfo info, List<string> arguments)
        {
            // Darwin's opaque types are pointers; Linux glibc/musl structures
            // fit the documented ABI storage (spawnattr 336, file actions 80).
            var attributes = Marshal.AllocHGlobal(Darwin ? IntPtr.Size : 336);
            var actions = Marshal.AllocHGlobal(Darwin ? IntPtr.Size : 80);
            var attributesReady = false;
            var actionsReady = false;
            var descriptors = new List<int>();
            try
            {
                using var argv = new NativeStrings(new[] { info.FileName }.Concat(arguments));
                using var environment = new NativeStrings(info.Environment.Where(pair => pair.Value != null)
                    .Select(pair => pair.Key + "=" + pair.Value));
                lock (LaunchLock)
                {
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
                    var output = CreatePipe(descriptors);
                    var error = CreatePipe(descriptors);
                    foreach (var descriptor in new[] { output[0], error[0] })
                        Check(posix_spawn_file_actions_addclose(actions, descriptor), "posix_spawn close");
                    Check(posix_spawn_file_actions_adddup2(actions, output[1], 1), "posix_spawn stdout");
                    Check(posix_spawn_file_actions_adddup2(actions, error[1], 2), "posix_spawn stderr");
                    foreach (var descriptor in new[] { output[1], error[1] })
                        Check(posix_spawn_file_actions_addclose(actions, descriptor), "posix_spawn close");
                    // Resolve like Process.Start, before the child's chdir. Passing
                    // an absolute path also preserves missing/invalid-executable errno.
                    var executable = ResolveExecutable(info.FileName);
                    Check(posix_spawn(out _pid, executable, actions, attributes, argv.Pointer, environment.Pointer), "posix_spawn " + info.FileName);
                    close(output[1]); descriptors.Remove(output[1]);
                    close(error[1]); descriptors.Remove(error[1]);
                    StandardOutput = Reader((IntPtr)output[0], info.StandardOutputEncoding); descriptors.Remove(output[0]);
                    StandardError = Reader((IntPtr)error[0], info.StandardErrorEncoding); descriptors.Remove(error[0]);
                }
            }
            catch
            {
                if (_pid != 0) { Terminate(); WaitForExit(5000); waitpid(_pid, out _, 1); }
                StandardOutput?.Dispose(); StandardError?.Dispose();
                Marshal.FreeHGlobal(_status);
                throw;
            }
            finally
            {
                foreach (var descriptor in descriptors) close(descriptor);
                if (actionsReady) posix_spawn_file_actions_destroy(actions);
                if (attributesReady) posix_spawnattr_destroy(attributes);
                Marshal.FreeHGlobal(actions); Marshal.FreeHGlobal(attributes);
            }
        }

        internal override bool HasExited
        {
            get
            {
                if (_exited) return true;
                for (var offset = 0; offset < 128; offset += 4) Marshal.WriteInt32(_status, offset, 0);
                // WNOWAIT reserves the leader PID until no further kill(-pgid)
                // can occur. Do not register this child with System.Diagnostics.Process.
                if (waitid(1, _pid, _status, 1 | 4 | (Darwin ? 0x20 : 0x1000000)) != 0)
                    throw NativeError("waitid on owned process");
                var code = Marshal.ReadInt32(_status, 8);
                if (Marshal.ReadInt32(_status, Darwin ? 12 : 16) != _pid || code < 1 || code > 3) return false;
                var status = Marshal.ReadInt32(_status, Darwin ? 20 : 24);
                _exitCode = code == 1 ? status : 128 + status;
                return _exited = true;
            }
        }
        internal override int ExitCode => HasExited ? _exitCode : throw new InvalidOperationException("Process has not exited.");
        internal override void Terminate()
        {
            if (_terminated) return;
            // Verify the waitable leader still belongs to us before signalling.
            _ = HasExited;
            if (kill(-_pid, 9) != 0)
            {
                var error = Marshal.GetLastWin32Error();
                // Darwin reports EPERM, rather than ESRCH, for an all-zombie
                // group. Prove that state before treating it as terminated.
                if (error != 3 && !(Darwin && error == 1 && HasExited && DarwinGroupHasOnlyZombies()))
                    throw new Win32Exception(error, "kill owned process group");
            }
            _terminated = true;
        }
        public override void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                Terminate();
                if (WaitAttempted ? HasExited : WaitForExit(5000))
                {
                    var result = waitpid(_pid, out _, 1);
                    if (result != _pid) throw NativeError("reap owned process");
                }
                else
                {
                    // A killed process stuck in uninterruptible kernel work may
                    // exit later. Reap only this reserved PID without holding up
                    // the caller or ever signalling the group again.
                    var pid = _pid;
                    _ = Task.Run(() => { while (waitpid(pid, out _, 0) < 0 && Marshal.GetLastWin32Error() == 4) { } });
                    throw new TimeoutException("Owned process did not exit within the cleanup deadline.");
                }
            }
            finally
            {
                StandardOutput.Dispose(); StandardError.Dispose(); Marshal.FreeHGlobal(_status);
            }
        }
        private bool DarwinGroupHasOnlyZombies()
        {
            var members = new int[4096];
            var bytes = proc_listpgrppids(_pid, members, members.Length * sizeof(int));
            if (bytes < 0 || bytes >= members.Length * sizeof(int)) return false;
            var status = new byte[64]; // proc_bsdshortinfo has a fixed 64-byte ABI.
            for (var index = 0; index < bytes / sizeof(int); index++)
            {
                if (members[index] == 0) continue;
                var length = proc_pidinfo(members[index], 13, 0, status, status.Length);
                if (length == 0 && Marshal.GetLastWin32Error() == 3) continue;
                if (length != status.Length || BitConverter.ToInt32(status, 8) != _pid ||
                    BitConverter.ToInt32(status, 12) != 5) return false;
            }
            return true;
        }
        private static int[] CreatePipe(List<int> descriptors)
        {
            var pair = new int[2];
            if ((Darwin ? pipe(pair) : pipe2(pair, 0x80000)) != 0) throw NativeError("pipe");
            descriptors.AddRange(pair);
            if (Darwin)
                foreach (var descriptor in pair)
                    if (fcntl(descriptor, 2, 1) < 0) throw NativeError("pipe close-on-exec");
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

        [DllImport("/usr/lib/libproc.dylib", SetLastError = true)] private static extern int proc_listpgrppids(int group, int[] members, int size);
        [DllImport("/usr/lib/libproc.dylib", SetLastError = true)] private static extern int proc_pidinfo(int pid, int flavor, ulong argument, byte[] buffer, int size);
        [DllImport("libc", SetLastError = true)] private static extern int pipe(int[] descriptors);
        [DllImport("libc", SetLastError = true)] private static extern int pipe2(int[] descriptors, int flags);
        [DllImport("libc", SetLastError = true)] private static extern int fcntl(int fd, int command, int value);
        [DllImport("libc")] private static extern int close(int fd);
        [DllImport("libc", SetLastError = true)] private static extern int access(string path, int mode);
        [DllImport("libc", SetLastError = true)] private static extern int kill(int pid, int signal);
        [DllImport("libc", SetLastError = true)] private static extern int waitid(int type, int pid, IntPtr info, int options);
        [DllImport("libc", SetLastError = true)] private static extern int waitpid(int pid, out int status, int options);
        [DllImport("libc")] private static extern int posix_spawnattr_init(IntPtr attributes);
        [DllImport("libc")] private static extern int posix_spawnattr_destroy(IntPtr attributes);
        [DllImport("libc")] private static extern int posix_spawnattr_setpgroup(IntPtr attributes, int group);
        [DllImport("libc")] private static extern int posix_spawnattr_setflags(IntPtr attributes, short flags);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_init(IntPtr actions);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_destroy(IntPtr actions);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_addchdir_np(IntPtr actions, string directory);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_addclose(IntPtr actions, int descriptor);
        [DllImport("libc")] private static extern int posix_spawn_file_actions_adddup2(IntPtr actions, int descriptor, int destination);
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
                CloseHandle(output[1]); handles.Remove(output[1]);
                CloseHandle(error[1]); handles.Remove(error[1]);
                StandardOutput = Reader(output[0], info.StandardOutputEncoding); handles.Remove(output[0]);
                StandardError = Reader(error[0], info.StandardErrorEncoding); handles.Remove(error[0]);
            }
            catch
            {
                if (_process != IntPtr.Zero) { TerminateProcess(_process, 1); WaitForSingleObject(_process, 5000); CloseHandle(_process); }
                if (_job != IntPtr.Zero) CloseHandle(_job);
                StandardOutput?.Dispose(); StandardError?.Dispose();
                throw;
            }
            finally
            {
                foreach (var handle in handles) CloseHandle(handle);
                if (thread != IntPtr.Zero) CloseHandle(thread);
                if (initialized) DeleteProcThreadAttributeList(attributes);
                if (attributes != IntPtr.Zero) Marshal.FreeHGlobal(attributes);
                if (attributeHandles != IntPtr.Zero) Marshal.FreeHGlobal(attributeHandles);
                if (environment != IntPtr.Zero) Marshal.FreeHGlobal(environment);
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
            try { Terminate(); if (!(WaitAttempted ? HasExited : WaitForExit(5000))) throw new TimeoutException("Owned process did not exit within the cleanup deadline."); }
            finally
            {
                CloseHandle(_job); CloseHandle(_process);
                StandardOutput.Dispose(); StandardError.Dispose();
            }
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
