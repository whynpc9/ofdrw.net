# PR 23 P1: native supervisor delivery proposal

Status: **prototype verified; production implementation and release gates pending**.
Source baseline: `26dc12087ddff71af31c1b46d5eb216723a36fad`.
Worker: `codex/pr23-process-reaper` (separate managed worktree).

## Defect and evidence

The .NET Unix signal loop intentionally reaps all children when its PID is 1.
`ProcessWaitState.CheckChildren(reapAll: true)` saves status only for its managed
child table; `GetProcessById` creates a non-child entry and cannot repair this.
An unregistered native child therefore loses both its real exit status and the
waitable zombie that reserved its process-group number. Swallowing ECHILD cannot
recover either, and signalling a cached negative group ID after that loss is unsafe.

Official source inspected: dotnet/runtime tags v8.0.0, v10.0.0 and v10.0.13,
`src/native/libs/System.Native/pal_signal.c` and
`src/libraries/System.Diagnostics.Process/src/System/Diagnostics/ProcessWaitState.Unix.cs`.
The inspected versions leave an inherited SIG_IGN disposition installed; kernel
auto-reaping also breaks the existing raw child. An ordinary managed Process wait
under inherited SIG_IGN can itself hang. This does not invalidate the separate
PID 1 repro, where managed Process.Start works normally.

Evidence is in `artifacts/pr23-process-reaper/` of this worker checkout:

- Linux ARM64, Ubuntu 24.04, .NET 10.0.13, Docker PID 1: after a managed Process
  initializes SIGCHLD, the unchanged runner fails for target exit 0 and exit 7
  with ECHILD. Its parent-first wrapper leaves the published sleep descendant alive.
- Same PID 1 environment, native executable prototype: actual exit 0 and 7 are
  preserved; parent-first wrapper reaches the original timeout, and its known
  descendant is absent after cleanup. No synthetic success code.
- Linux x64 WSL, .NET 10.0.12: default disposition controls pass for both backends;
  inherited SIG_IGN raw backend fails for 0/7, supervisor passes for 0/7.
- `native-prototype/prototype-manifest.json` records source/binary SHA256 values.
  These binaries are explicitly **not publishable**: prototype toolchains were
  not pinned and the minimal managed adapter has incomplete failure handling.

## Choice: independent executable

Use one small first-party native executable, started with existing posix_spawn
file actions. It starts in a new group outside the target's group. Its fresh
single-threaded native process contains no CLR, runtime signal thread, managed
callbacks, or host pthread_atfork registrations. It then forks and execve's the
target after establishing the target group, and alone owns target waiting.

A shared library could fork an all-native supervisor branch without returning to
managed code. Its first fork would still occur in the multithreaded CLR host,
inherit host/native-library state and atfork registrations, and require a broader
fork-safety proof. It does not reduce immutable RID asset or license requirements.
The independent executable avoids that proof burden and is the proposed product.
No compiler, shell, Python or extra runtime is invoked on consumer machines.

## Process and protocol contract

- Host resolves executable and parses Arguments/ArgumentList as today. argv carries
  helper path, resolved target path, original argv[0], then exact target arguments.
  The same target environment and working directory are applied before launch.
- Private CLOEXEC stdout/stderr pipes retain existing capture and drain semantics.
  A private socketpair is mapped to supervisor fd 3. All source descriptors must
  be above fd 3 before constructing dup/close actions; this also covers hosts with
  closed stdin/stdout/stderr. The helper closes unrelated inherited descriptors.
- Helper restores its own SIGCHLD to SIG_DFL without SA_NOCLDWAIT and uses no
  signal reaper. It preserves the target's inherited ignored signal dispositions
  and mask across the supervisor-only adjustments. The fork-to-exec target path
  uses only async-signal-safe operations; no managed code can run there.
- Child setpgid precedes any target exec. Parent also establishes/verifies that
  group immediately after fork before processing commands. A CLOEXEC error pipe
  distinguishes target exec success from exact native errno (including ENOENT,
  EACCES and ENOEXEC), without shell fallback or parsing diagnostics.
- Helper reports Started, target Exited (real wait status), target StartFailed,
  Cleaned and protocol/native failure. Fixed framing has a version/magic and
  validates state ordering, target identity and payload lengths. Unknown/partial
  frames never imply success. Host initial/partial reads use poll with a fixed
  monotonic deadline and cancellation checks; there is no blocking unbounded read.
- Helper retains target leader with waitid(WNOWAIT) after target exit. An exited
  wrapper therefore does not release its group ID while descendants own pipes.
  Only helper, while it still has that waitable child, may kill(-targetGroup).
  Host never signals a numeric target or helper PID/group.
- Cleanup command, host socket EOF, corrupt protocol, helper control I/O failure,
  and catchable supervisor termination all use the same kill/reserved-wait path.
  Read exec-error and host control concurrently, so cancellation does not wait
  indefinitely for target exec. Group creation must be confirmed before any group
  signal; pending launch cancellation must not allow target code after ownership
  release. Helper sends real exit status before cleanup acknowledgement.
- Helper reaps target only after its final group signal and reports native cleanup
  failure. Cleanup has a monotonic deadline. If kernel uninterruptible state
  prevents exit, no further group signal occurs after the helper abandons waiting;
  the limitation is reported. Darwin EPERM is accepted only after a bounded query
  proves all members of this exact retained group are zombies/gone.
- On host construction failure, close protocol to trigger native cleanup, release
  every fd/stream independently, preserve original error, and attach diagnostics.
  The host may reap helper by exact child waitpid after protocol completion;
  ECHILD for this helper is acceptable because no helper exit code determines the
  target result and the host never signals that PID. This is not target ECHILD
  suppression. Unexpected helper/protocol failures are IOException, not target
  Win32Exception, so DOCX does not call them a missing LibreOffice installation.
- Ordinary completion, cancellation, timeout, stable-output cleanup, output floods,
  parent-first exit, no-collateral and disposal retain the existing runner contract.
  Deliberate target setsid/setpgid escape remains the existing explicit boundary.

## Proposed source and immutable asset files

Production files, after root design approval:

- `native/Ofdrw.ProcessSupervisor/supervisor.c` and `README.md`: source/protocol,
  platform branches, compile command and ABI/support boundary.
- `native/Ofdrw.ProcessSupervisor/assets/manifest.json`: protocol version, source
  SHA256, complete asset list, architecture/OS headers, minimum OS, toolchain image
  digest or exact Xcode/SDK build, compiler/link flags, binary SHA256 and sizes.
- Immutable committed executables in `assets/linux-x64`, `assets/linux-arm64`,
  `assets/osx-x64`, `assets/osx-arm64`. Linux is statically linked against pinned
  musl, and the same respective binary must execute on glibc and musl hosts.
  macOS could use a universal file, but two explicit assets simplify manifests.
- `scripts/build-process-supervisor.py` (build-time only) and
  `scripts/verify-process-supervisor-assets.py`: deterministic source build,
  complete-manifest validation, native-header/static-link checks, license checks.
  Linux build uses a pinned immutable toolchain image, pinned musl/source and
  deterministic linker flags. macOS fixes deployment target and supported SDK;
  reproducibility claims must match actual rebuild SHA evidence, including
  disabling timestamp/random build identifiers where applicable.
- `src/Ofdrw.Net.Core/Processes/NativeSupervisorAssets.cs`: select OS/process
  architecture using netstandard-compatible APIs, verify manifest hash and extract
  only selected embedded executable into mkdtemp-created 0700 private directory.
  Use exclusive new file, close writer, chmod 0700, then spawn; no shared writable
  cache or PATH helper lookup. Clean only owned extraction paths.
- `OwnedProcessGroup.cs`: retain Windows Job code; replace Unix raw target wait /
  negative-group signal with helper protocol; reuse argument/native stream helpers.
- `ExternalProcessRunner.cs`: pass cancellation to startup and preserve bounded
  cleanup/error attachment. No renderer changes or public API/dependency changes.
- Core csproj embeds **all four** immutable executables and manifest. Before build
  or pack, validate complete asset hashes/source binding; missing any asset is an
  error. Ubuntu pack therefore embeds macOS and both Linux assets too. There is
  no publishable current-RID-only mode. Normal builds do not require a compiler.
- `THIRD-PARTY-NOTICES.md` plus `docs/licenses/MUSL-COPYRIGHT.txt`: exact musl
  version/source/license/copyright and any statically linked compiler support
  attribution. Include retained license in Core package, and align existing
  package notice verification. No synthetic NuGet dependency.
- Tests: isolated probe host/fixtures, new process regressions in Core tests,
  native protocol fault tests and package-consumer extraction/start probe.
  Extend existing process CI filter so the new class is actually executed.

## Gates before production completion

1. Freeze reviewed C/protocol and establish exact source SHA. Build all four assets
   from pinned inputs; record architecture, dependency/static proof, license and
   reproducibility evidence. Do not promote prototype binaries into this set.
2. Source-level review of every native error / fork child path, descriptor ownership,
   signal changes, parent death, partial protocol and bounded deadline.
3. Native old tests plus new real 0/7, native exec failures, exact argv/cwd/env,
   flood, cancellation during startup, partial/EOF protocol and helper crash.
   Verify parent-first timeout/cancel/stable-output, no-collateral and known PIDs.
4. Linux x64 and ARM64, on glibc and musl, run target host as actual PID 1 and
   with inherited SIG_IGN in isolated subprocesses/containers. Do not infer one
   OS/ABI from another. macOS x64/ARM64 and Windows existing Job regressions
   require actual execution or explicit uncompleted evidence.
5. Rebuild Core netstandard2.0/2.1, full affected regressions and docs checks.
   Root owns shared validation lock, full solution/package/renderer gates.
6. From final Ubuntu-built local packages, extract/launch selected helper on each
   Unix native architecture/runtime consumer, with no runtime cc/Python/bash and
   read-only package directories. Verify SHA and private extraction permissions,
   cleanup, actual target exit and native errors. All 11 packages' retained notice
   checks must pass; a cross-platform package claim requires the complete assets.
7. Keep 32-bit Unix explicitly rejected. The existing nonempty-working-directory
   addchdir_np boundary remains (macOS 10.15+, glibc 2.29+ or supporting musl);
   absence must fail clearly before target executes. Exact supported Linux kernel
   and macOS deployment minimums follow the asset/toolchain manifest and tests.

Only source/prototype/design evidence exists at this checkpoint; this proposal
is not a claim that P1 is fixed in production or the PR is ready to merge.
