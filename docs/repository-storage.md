# Repository storage

## Default checkout

Binary archives, document/image fixtures and fonts are stored in Git LFS.
`.lfsconfig` downloads only `e2e/**` and `tests/**` by default. CI skips automatic
smudging and explicitly fetches those inputs; the public-table evidence job
also fetches its own retained bundle. This avoids downloading all historic
visual evidence on every build.

Use the [clone instructions](../README.md#clone-and-large-files) for a shallow
checkout before the history cutover, or to fetch one evidence directory.
`git archive` exports LFS pointers rather than hydrated payloads: use an LFS
checkout when running tests or reviewing binary inputs from a source snapshot.

## What to retain

- Keep deterministic, shared input fixtures and regression tests that exercise
  distinct behavior. Store binary fixtures in LFS, and reuse them across tests.
- Keep a compact report of the source commit, environment, commands, actual
  results, inspected pages and remaining limitations.
- Keep only the input/output pages, screenshots or failed-run records required
  to reproduce a defect or assess the reported acceptance scope. Put large
  retained payloads in one optional LFS bundle with a size/SHA-256 manifest.
- Leave build output, caches, repeated source trees and routine full test logs
  in ignored `artifacts/` or CI artifacts. A new source commit is not a reason
  to retain another copy of every previous artifact.

Existing validation history is preserved. The 1,171 directly tracked PR23
artifact files were consolidated without changing their bytes: 42.4 MiB became
one 7.6 MiB compressed LFS bundle. Its [index and manifest](evidence/pr23-integration/README.md)
restore the original paths. No existing .NET regression tests were removed;
the large footprint came from retained evidence rather than test source.
Routine solution-test console verbosity was reduced while retaining the tests
and existing TRX reporting.

Run `python3 scripts/check-repository-storage.py` after staging changes. CI
checks the staged/checked-out Git objects, rejects tracked `artifacts/` files
and ordinary blobs above 5 MiB, and verifies required fixture hashes. This
check works with optional evidence left as pointer files.

## History cutover

Adding pointers in a new commit does not remove the old binary blobs from
ordinary Git history. The separate local rehearsal used `git lfs migrate import`
with all repository binary patterns, then restored the eight published release
tags to their original object IDs. It verified 1,493 conversions across 28
branch tips against the original payload hashes. A full local transport clone
of the migrated history transferred a 7,940,417-byte Git pack, plus 50,406,029
bytes of explicitly fetched development LFS payloads. Optional evidence stayed
as pointers. This is a local migration measurement, not proof that GitHub's
remote history has already changed.

The cutover must use the reviewed storage branch as the new `main` tip so the
default download policy and artifact consolidation take effect together. Other
existing remote branches can use their mapped migrated tips. Published release
tags stay unchanged. The prepared local backup, original remote-ref snapshot,
old/new commit map and branch map are outside the tracked source tree under the
migration worktree's ignored `artifacts/lfs-migration/` directory.

Before publishing rewritten branch history:

1. Upload every referenced LFS object and verify a fresh remote clone can fetch
   the required development payloads.
2. Re-read remote branch/tag IDs and stop if they differ from the snapshot.
3. Push only explicitly mapped branch refs, using an exact `--force-with-lease`
   for each old ID. Do not mirror-push, delete branches or move release tags.
4. Verify the new remote `main`, unchanged tags, clone size and LFS retrieval.

This changes historical commit IDs and requires updating existing worktrees and
other machines. Prefer a fresh clone after saving any local work. Do not merge
an old, unmigrated branch back into the migrated repository; that would restore
the original large Git history. Old acceptance records still name their actual
historical source commits; the saved commit map relates them to migrated IDs.
GitHub may retain old objects for pull requests, so the repository's displayed
server-side size can lag behind the size transferred by a fresh normal clone.
