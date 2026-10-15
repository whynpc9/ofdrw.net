# Plan 024 process ownership design

Baseline: c03d669309d6e61444de9e9bfdde142537c56dd3. Process sources match audit; CI contains Plan 004 native path tests, which must remain. Plans are not tracked in this baseline, so this design is added explicitly.

## Staged evidence

Step 1 macOS: six tests discovered, normal exit/output and existing tree pass; four fail on child survival after wrapper exit. All finally blocks clean only fixture-published PIDs, including the unrelated control. Evidence: artifacts/plan-024/step1. Step 2 repeats tests before implementation and retains the same expected survival failures; design file existence is checked separately.

## Ownership and launch race

Process.Start followed by setpgid or Job assignment is racy. Ownership must exist before any target code executes; removing HasExited cannot fix orphan ownership.

Unix (macOS/Linux): posix_spawn with POSIX_SPAWN_SETPGROUP and pgroup zero atomically establishes the group before exec. File actions redirect private stdout/stderr pipes. Direct native exec preserves missing/invalid executable errors. A nonempty working directory uses posix_spawn_file_actions_addchdir_np; an absent entrypoint fails clearly and closed (macOS before 10.15 or older Linux libc, including glibc before 2.29). Empty WorkingDirectory does not require this extension. No shell bootstrap is used. ProcessStartInfo.Arguments uses .NET-compatible Unix quoting/backslash parsing; ArgumentList is read via reflection for netstandard2.0. Environment comes from ProcessStartInfo. waitid(WNOWAIT) observes exit without reaping, reserving the leader PID until group cleanup and avoiding group-ID reuse. kill(-pgid, SIGKILL) therefore works after wrapper exit. Bounded waitpid reaps the owned leader on disposal. Native opaque storage is allocated for Darwin/Linux ABI sizes; other Unix platforms fail explicitly.

Windows: create a private Job with KILL_ON_JOB_CLOSE. CreateProcessW starts suspended, with a restricted inherited-handle list for private pipes. Assign the suspended target to the Job before ResumeThread; assignment failure terminates that exact suspended target and reports the native error. Retain Job/process handles after wrapper exit. TerminateJobObject cleans descendants; close the Job on normal completion too. No job breakaway, privileges, process-name scans, or machine changes. Preserve Windows argument quoting, environment, working directory and executable lookup.

## Cleanup and compatibility boundary

Runner still caps both captures at 65,536 characters and includes inherited-pipe completion in its deadline. Cancellation/timeout/stable-output cleanup uses owned group/job independent of HasExited. Parent wait is bounded to five seconds, pipe drain to one second. Preserve original cancellation/timeout when cleanup fails and attach cleanup diagnostics. Dispose private handles/streams and reap the leader.

Unix descendants deliberately leaving their process group/session are outside the group guarantee. Preventing deliberate daemon escape needs a separate supervisor/container design. Office stable-output handling remains opt-in through expectedOutputPath. Existing Office instances activated through external GUI services are not owned and are never enumerated or killed.

No public API/dependency/TFM changes: internal helper only, netstandard2.0/2.1 compile. Preserve Arguments/ArgumentList, spaces/Unicode, environment/cwd, normal/nonzero/flood capture, missing executable and Office output stabilization. Verify timeout/cancellation after parent-first exit, old tree, unrelated same-name fixture and bounded cleanup. Native CI ORs ExternalProcessTests with Audit024RegressionTests while retaining Plan 004 subsets. Run solution restore/build, full Core/DOCX regressions and API docs gate. Process-only changes require no new visual artifact and claim no rendering acceptance.

Darwin refinement: kill on an all-zombie process group returns EPERM. A bounded libproc query checks only members of that exact owned group; EPERM is accepted only when all reported members are zombies/gone, never when a live member remains. The leader is reaped afterward. Tests recognize OS-proven zombie state only while waiting for the parent-before-cleanup event; after cleanup both known PIDs must disappear.
