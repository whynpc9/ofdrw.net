# Owned Unix process supervisor

`supervisor.c` is a first-party native executable embedded in both Core target
frameworks. Windows continues using its native Job backend. Linux x64/ARM64 assets
are statically linked against musl; macOS x64/ARM64 assets use system libraries.
No compiler, shell, Python or separately installed runtime is used by consumers.
See `assets/manifest.json` and `PROVENANCE.md` for exact build and license inputs.

The host starts a fresh supervisor with `posix_spawn`, a private socket at fd 3
and separate stdout/stderr pipes. The supervisor is outside the target group and
owns its direct child. That child establishes its group and reports readiness,
then waits for a grant before `execve`; native errors travel through a CLOEXEC
error pipe. Before the grant, cancellation can kill that exact waitable child;
after the grant, only the native owner may signal the reserved target group.
No code forks the multithreaded CLR host or runs managed code after a native fork.

The supervisor restores its own SIGCHLD to default while preserving the target's
original signal mask and ignored dispositions. It observes target exit using
`waitid(WNOWAIT)`, retains the leader until the final group signal, and reaps it
only afterward. This remains valid if the managed host is PID 1 or inherits
SIGCHLD=SIG_IGN: the host runtime never owns the supervisor's child wait status.

Protocol version 1 frames are 16 bytes: four network-order uint32 values,
`0x4f465301`, kind, target PID, value. Kinds: 1=exec succeeded (value 0), 2=actual
target exit (exit status, or 128+signal), 3=target exec failed (native errno),
4=cleanup acknowledged (0), 5=supervisor failure (native errno). Kind 3 is followed by a same-target cleanup acknowledgement or failure; the host
retains the original errno and attaches any unconfirmed cleanup diagnostic. Only
kind 3 maps to a target native start exception. Malformed frames, helper failures, missing or
nonexecutable extracted assets and unexpected EOF are I/O failures.

The host sends one `K` byte to request cleanup. EOF, invalid commands and catchable
supervisor termination also clean the owned group. Native writes and managed
reads/writes have monotonic deadlines; startup counts toward the caller's timeout
and checks cancellation. Cleanup is bounded and reports unconfirmed termination.
There is no host fallback that signals a cached PID/group or scans process names.

An uncatchable supervisor crash or SIGKILL can release its wait reservation before
cleanup. The host reports an I/O failure and cannot safely reconstruct ownership
or guarantee descendant removal in that case. Deliberate target process-group or
session escape remains outside this ownership guarantee. Kernel uninterruptible
states may also prevent confirmed exit within the cleanup deadline. None of these
failures produces a synthetic successful exit code.

Extraction uses a fresh native `mkdtemp` 0700 directory, verifies embedded size
and SHA256, creates the executable exclusively, closes the writer, then applies
0700. The protocol is closed before deleting that owned directory. A noexec temp
filesystem is reported explicitly. Assets are source-bound by SDK `GetFileHash`
checks on every Core compile/pack; packing additionally requires the complete
manifest's publication gates to be approved. Python verification/build tools are
for maintainers and CI, not normal Core builds or consumers.

32-bit Unix remains unsupported. Nonempty WorkingDirectory still requires the
existing `posix_spawn_file_actions_addchdir_np` extension (macOS 10.15+ or a
supporting libc). The manifest specifies native deployment minimums; successful
execution on one architecture is not evidence for the others.
