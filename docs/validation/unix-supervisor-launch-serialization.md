# Unix supervisor extraction and launch ownership

The first CI run for merged main `b1a8b6319fcc840d0c711569d8232990cdeef34a`
([41244813317](https://github.com/whynpc9/ofdrw.net/actions/runs/41244813317))
failed in `WholePageFallbackDiscardsPriorNativePaintAndThenNextPageIsNative`.
`posix_spawn` returned `ETXTBSY` (26) while starting the temporary native
supervisor. The package-producing job stopped; that main run is not accepted.
The native supervisor source was unchanged by PR80.

A local Linux ARM64 diagnostic using the actual original Core DLL observed
three matching failures in 512 `/bin/true` launches at concurrency 16. Subsequent
8,192-launch runs of both the original and the first candidate DLL had no
failures. Those equal-size runs do not establish a statistical old/new
improvement; all results, including the first failure, remain in ignored
`artifacts/`.

## Change

Supervisor extraction now takes place under the same ownership gate as every
owned `posix_spawn`. Previously another owned launch could spawn while the
extractor's executable writer was open. The writer was closed before its own
spawn, but that did not serialize it against other spawns in the process.
The source now guarantees that these owned write/spawn windows cannot overlap.
The private directory, immutable asset hashes, file permissions, per-launch
extraction, protocol and target-status handling are retained.

Acquiring ownership checks cancellation and the original startup deadline,
using waits of at most 25 ms and the remaining budget. It checks again after
acquisition and releases ownership on extraction, argument preparation or
spawn failure. No execution-error retry or test skip was introduced. Synchronous
filesystem calls already in progress are not claimed to be interruptible.

## Evidence boundaries

A separate deterministic Linux experiment opened an executable writer, forked a
child that inherited the descriptor, then closed the parent's writer. The
parent descriptor was confirmed closed (`EBADF`); the inherited descriptor had
`FD_CLOEXEC`. The next spawn returned 26 while the child held it. After releasing
and reaping that child, the same unchanged executable spawned and exited
successfully. A no-writer negative control also succeeded. This establishes the
OS admission mechanism, not the particular child/descriptor involved in the
first CI failure.

The shared fixture observes startup in an isolated process and temporary
directory while a separate thread owns the real launch gate. With the same
fixture DLL, the original Core DLL published an extraction directory before the
gate was released (exit 43), while the candidate observed none (exit 0). Both
subsequently executed the marker-writing target and removed extraction files.
The worker announcement precedes `RunAsync`, and `WaitSleepJoin` does not identify
which lock is awaited: this is finite controlled observation, not deterministic
proof of the sampling point. Source review separately establishes the lock
boundary.

Regression tests also cancel or expire startup while the gate remains held,
check that the target marker was never written, release the owner, and confirm
that a normal launch still succeeds. Windows retains its Job backend; these
Unix-specific cases are explicitly inapplicable there.

An arbitrary foreign fork/spawn in the same process does not participate in this
gate. The original CI event's precise causal chain and immunity to such callers
remain unproved. CI, package and main acceptance must be recorded for the final
immutable candidate; earlier candidate DLL observations do not transfer to it.
This process fix does not close PR78 quantity-counter readability, terminal-date
Find, general C3/C4 or complex-script document-output limitations.
