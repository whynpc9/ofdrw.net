# PR19 R3 missing-parent repair — final static closure

No blocking finding in this two-file repair. Reviewed actual diff against current `e5102cc437a832f9c8079a7a5ceb06a992812a94`; no product edits, builds/tests, GUI, commits or owner09 contact performed.

Both Noop and Changed constructors invoke the same private destination-directory helper after actual bound scope admission and existing stage reservation, but before stage path generation, observation binding or FileStream creation. The helper computes checked path-prefix quantities and prepays Copy `4096+4*(n+1)^2+256*(n+1)` plus Work `128*(n+1)+4*(n+1)^2` on the same material context before new directory work. It then uses the existing fresh destination guard, creates the parent, rechecks context/token, and repeats the existing destination guard before stage creation. Overflow, cancellation, quota or IO errors propagate through the existing factory catch to terminal owner failure. These are logical bounded path/metadata charges, not a claim of interruptible OS IO or hard OS CPU/RSS limits.

The original adjacent CreateNew stage, observation accounting, stage-copy close, final source/raw/token checks, fresh commit destination guard, precommit cleanup, and single successful rename semantics are unchanged. No new authority, filesystem service, profile, ledger reset/cap increase or postcommit operation was added. A directory may remain if later preparation fails; the code does not perform unsafe recursive rollback of shared parents.

The five added actual CLI instances cover both positive branches (Noop byte-exact and Changed with fresh public Reader plus original-font Locator/neighbor position), a regular-file ancestor, a real directory-symlink ancestor preserving link and target sentinel, and invalid source with the entire missing-parent chain still absent. Source bytes and output-stage residue are asserted. The two positive instances ensure the new parent initially does not exist. The helper is exercised through actual qualified CLI save; it is not replaced with a mock or forged scope. Windows link-API unavailability is narrowly caught and explicitly logged as NOT EXECUTED; such a returned test must not be counted as executed symlink-rejection evidence. Linux/macOS failures in link creation still fail the test.

Before/after exact hash readback:

- `src/Ofdrw.Net.Packaging/OfdOwnedAtomicStage.cs`: `ace1aa2a686a8d8a0ea5d45268c99713d330fa92564eaee27aed34d8b34951c1`
- `tests/Ofdrw.Net.Cli.Tests/TextReplacementCommitNotificationTests.cs`: `644406d9775c9a9bcbb2f61de4d5120d47b6727eaa7f3b8b7ca6549f3a3bc0f5`

Eligible for the planned Low affected checks and the newly required full-original-CJK Changed run under the unchanged four default caps. This static report does not claim the new path charges fit that runtime case. New metrics/evidence must be attached to this candidate; historical ZIP/counter records remain unchanged. No page-content/rendering change is introduced, so the existing Preview content identity can be retained as authorized, without claiming a newly performed visual review.
