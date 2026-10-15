# PR19 missing output parent — static P2 and minimal repair

Baseline `e5102cc437a832f9c8079a7a5ceb06a992812a94`. Read-only source review; no build/test, product edits, GUI, owner09 contact or reference movement.

## Confirmed defect

`src/Ofdrw.Net.Packaging/OfdOwnedAtomicStage.cs:27–32` (Noop) and `:40–45` (Changed) extract the canonical destination parent and immediately open a neighboring CreateNew stage. Neither creates a missing parent. A valid command with destination `new/sub/output.ofd` fails with DirectoryNotFoundException when `new/sub` does not exist. The actual default CLI calls Layout `OfdPrivateSaveWorker`, then `OfdPrivateAtomicBridge.CreateNoop` / `Create`, then these actual Packaging constructors; the bridge is not the defect's implementation location.

This is a default-command behavior regression, not an explicit existing-directory profile. Existing `src/Ofdrw.Net.Cli/AtomicOutput.cs:12` creates the parent. The keyword command/help/tutorial do not document an existing-parent restriction. Core's `RequireCanonicalPath` deliberately tolerates missing files/directories while rejecting noncanonical paths and existing reparse points. The bound destination therefore passes initial path admission and only fails later at the stage open.

## Minimum owner-local implementation

Keep both constructors and their real scope admission. Immediately after extracting/validating the parent, before generating/opening the stage, invoke a small **private helper in the same Packaging class**:

1. Prepay the new parent-directory operation and any associated transient path metadata with the same `context.Material`, using checked arithmetic and the original token. Preserve the existing stage buffer/path reservation and CreateNew Work charge; do not silently spend them twice on unaccounted work.
2. Call existing `owner.RequireCurrentDestination(token)` immediately before directory creation. This already checks the live context/token and the bound canonical destination/ancestor reparse policy; no new path parser, filesystem service, issuer or caller authority parameter is needed.
3. `Directory.CreateDirectory(parent)` (also valid if the directory already exists), followed by context/token and fresh existing destination admission **before** creating the stage. Thus the newly created ancestor chain is checked too. This is cooperative boundary cancellation, not an interruptible-directory-syscall claim.
4. Continue the existing adjacent unique path, observation binding, Work reservation and FileMode.CreateNew open. All newly introduced fallible preparation stays before FileStream creation, so it cannot create a new constructor-cleanup leak. Static factory catches already fail the owner on any exception.

Directory.CreateDirectory can create multiple ancestor segments and materialize prefixes internally. Use a conservative finite, path-length-derived reservation rather than one unpriced fixed syscall. For example, with `n = parent.Length`, a checked quadratic character-prefix allowance (at most `n+1` prefixes, each at most `n` characters), plus bounded per-prefix shells/work, is a defensible logical-accounting model. A concrete conservative choice is `Copy = 4096 + 4*(n+1)^2 + 256*(n+1)` and `Work = 128*(n+1) + 4*(n+1)^2`, prepaid before the BCL call; existing fresh canonical checks charge their own walks. An equivalent justified tighter accounting is acceptable. This prices logical path/metadata work; it must not claim a hard OS filesystem CPU/RSS bound. No four-ledger cap increase/reset/refund or alternate context is needed.

Do not move directory creation into CLI parsing/bootstrap: malformed input or failed readback/equivalence must not create output directories. Keep creation behind actual Noop verification or successful Changed readback/geometry scope. Do not recursively remove newly created parents on failure: directories can become shared, and the required atomic guarantee is no published target/no residual stage, not transactional rollback of directory metadata. Preserve final raw-input/source/token checks, precommit throwing cleanup, fresh destination admission at commit, one real rename, and no fallible postcommit cleanup.

## Necessary actual regressions and affected checks

- Actual `Cli.RunAsync`, valid OFD/explicit deterministic font, nested previously absent parent, **Noop and Changed**. Assert parent absence before invocation, exit0, unchanged source bytes/hash, Noop byte-exact target, Changed fresh public Reader/locator replacement and untouched neighbor evidence, and no adjacent stage residue. These cover both actual constructors; a direct helper or mocked stage test alone is insufficient.
- Existing ancestor is a regular file: nested output fails, the blocker file and source remain byte-identical, no output publication or stage residue.
- Actual symlink ancestor into an existing directory: reject under the unchanged reparse policy, preserve source, link identity and any pre-existing target sentinel, and create no stage. Use the real link API; record unsupported-platform setup separately, with real Linux/macOS execution rather than a fabricated link.
- Invalid input with missing output parent: failure leaves source bytes intact and does not create the destination/parent. The existing bad-input/existing-target sentinel case remains relevant. Do not create directories before source verification merely to satisfy the positive tests.
- Keep existing same-path/case-only/alias tests intact. A quota/cancellation admission test for the new directory preparation should establish rejection before filesystem mutation if exposed through existing internal test access; do not add public hooks or fake issuing scopes just for it.

Low should run full CLI tests, affected Packaging owned-stage/boundary/session/observation tests, and netstandard2.0/2.1 Packaging builds plus affected downstream Layout/Reader/CLI compilation; retain the established shared Core checks in the planned affected gate. Verify actual platform branches in Ubuntu CI/local macOS. This path-only correction does not change generated page content, fonts, budget limits or profiles, and does not require regenerating unchanged Preview assets. No runtime or overall acceptance is asserted by this static review.
